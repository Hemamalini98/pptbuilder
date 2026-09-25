"""Automatic template PPTX repair.

Patches common defects in the template file so downstream tools (extract,
convert, master_swap) see a well-formed template. The primary defect
observed in customer-supplied templates is `<p:ph>` elements missing the
`idx` attribute — PowerPoint still opens the file, but python-pptx and
most style-inheritance chains rely on `idx` to key placeholders to their
layout counterparts. Repair rewrites the template on disk so no in-memory
workaround is needed.

Repairs performed:

1. **Missing `<p:ph>` `idx` attribute** — inferred from `type` (title=0,
   ftr=10, sldNum=11, others assigned first free idx per slide/layout).
2. **Missing `<p:ph>` `type` attribute** — inferred from shape name
   (`Title 1` → `title`, `Content Placeholder 2` → `body`) when the shape
   is clearly a placeholder by name convention.
3. **Duplicate `idx` within one layout/master** — later duplicates are
   reassigned to the first free integer.
4. **Missing `<a:extLst>/<a:ext>` `<a16:creationId>`** — a fresh GUID is
   injected. PowerPoint uses these for internal tracking; without them,
   some features (comments, review) misbehave.

Not repaired (out of scope):
- Broken relationships between slides and layouts (structural)
- Missing theme parts (would break entire deck)
"""
from __future__ import annotations

import shutil
import uuid
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

from lxml import etree

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A16_NS = "http://schemas.microsoft.com/office/drawing/2014/main"

NS = {"a": A_NS, "p": P_NS, "a16": A16_NS}


def _qn(prefix: str, tag: str) -> str:
    return "{" + NS[prefix] + "}" + tag


# Standard idx assignments for known placeholder types (per OOXML convention).
STANDARD_IDX = {
    "title": 0,
    "ctrTitle": 0,
    "subTitle": 1,
    "ftr": 10,
    "sldNum": 11,
    "dt": 12,
}


@dataclass
class RepairReport:
    """Summary of repairs applied to one template file."""

    file: str
    fixes: list[dict] = field(default_factory=list)

    def add(self, part: str, defect: str, resolution: str) -> None:
        self.fixes.append({"part": part, "defect": defect, "resolution": resolution})

    @property
    def count(self) -> int:
        return len(self.fixes)


# ---------------------------------------------------------------------------
# Placeholder repair
# ---------------------------------------------------------------------------


def _infer_type_from_name(shape_name: str) -> str | None:
    """Best-effort mapping from shape name → placeholder type.

    Templates authored in PowerPoint use conventional names like
    `Title 1`, `Content Placeholder 2`, `Footer Placeholder 3` which
    reliably indicate the intended type.
    """
    if not shape_name:
        return None
    lower = shape_name.lower()
    if "title" in lower:
        if "sub" in lower:
            return "subTitle"
        if "center" in lower:
            return "ctrTitle"
        return "title"
    if "content" in lower or "object" in lower:
        return "body"
    if "footer" in lower:
        return "ftr"
    if "slide number" in lower or "sldnum" in lower:
        return "sldNum"
    if "date" in lower:
        return "dt"
    if "picture" in lower or "image" in lower:
        return "pic"
    if "text" in lower:
        return "body"
    return None


def _first_free_idx(used: set[int], starting_from: int = 1) -> int:
    """Smallest integer ≥ `starting_from` not already in `used`."""
    i = starting_from
    while i in used:
        i += 1
    return i


def _shape_name(sp: etree._Element) -> str:
    """Return the shape's cNvPr name, or empty string."""
    cNvPr = sp.find(_qn("p", "nvSpPr") + "/" + _qn("p", "cNvPr"))
    return cNvPr.get("name", "") if cNvPr is not None else ""


def _add_creation_id_if_missing(sp: etree._Element) -> bool:
    """Ensure cNvPr/extLst/ext/a16:creationId exists. Returns True if added."""
    cNvPr = sp.find(_qn("p", "nvSpPr") + "/" + _qn("p", "cNvPr"))
    if cNvPr is None:
        return False
    extLst = cNvPr.find(_qn("a", "extLst"))
    if extLst is not None and extLst.find(_qn("a", "ext")) is not None:
        return False
    if extLst is None:
        extLst = etree.SubElement(cNvPr, _qn("a", "extLst"))
    ext = etree.SubElement(extLst, _qn("a", "ext"))
    ext.set("uri", "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}")
    creation_id = etree.SubElement(ext, _qn("a16", "creationId"))
    creation_id.set("id", "{" + str(uuid.uuid4()).upper() + "}")
    return True


def _repair_placeholders_in_part(
    xml_bytes: bytes,
    part_name: str,
    report: RepairReport,
) -> bytes | None:
    """Repair all placeholder defects in one slideMaster/slideLayout XML.

    Returns the repaired bytes if anything changed, None otherwise.
    """
    root = etree.fromstring(xml_bytes)
    changed = False

    used_idx: set[int] = set()

    # First pass — collect existing valid idx values so we don't reassign
    # duplicates onto used ones.
    for ph in root.iter(_qn("p", "ph")):
        idx_str = ph.get("idx")
        if idx_str is not None and idx_str.isdigit():
            used_idx.add(int(idx_str))

    for sp in root.iter(_qn("p", "sp")):
        ph = sp.find(_qn("p", "nvSpPr") + "/" + _qn("p", "nvPr") + "/" + _qn("p", "ph"))
        if ph is None:
            continue

        ph_type = ph.get("type")
        ph_idx = ph.get("idx")
        name = _shape_name(sp)

        # Defect 1: missing type — infer from name.
        if ph_type is None:
            inferred = _infer_type_from_name(name)
            if inferred and inferred not in ("body",):
                # 'body' is the *implied* default when type is absent, so
                # don't add it unless there's a specific type to declare.
                ph.set("type", inferred)
                report.add(
                    part_name,
                    f"placeholder {name!r} missing type attribute",
                    f"added type={inferred!r} (inferred from shape name)",
                )
                ph_type = inferred
                changed = True

        # Defect 2: missing idx — infer from type or assign next free int.
        if ph_idx is None:
            if ph_type in STANDARD_IDX:
                new_idx = STANDARD_IDX[ph_type]
                # If the standard slot is already taken, pick next free.
                if new_idx in used_idx:
                    new_idx = _first_free_idx(used_idx, starting_from=new_idx + 1)
            else:
                # Body/object placeholders start at idx 1 (title owns 0).
                new_idx = _first_free_idx(used_idx, starting_from=1)
            ph.set("idx", str(new_idx))
            used_idx.add(new_idx)
            report.add(
                part_name,
                f"placeholder {name!r} missing idx attribute",
                f"added idx={new_idx} (inferred from type={ph_type!r})",
            )
            changed = True

        # Defect 3: creationId GUID missing (harmless but affects some
        # PowerPoint tracking features).
        if _add_creation_id_if_missing(sp):
            report.add(
                part_name,
                f"shape {name!r} missing creationId GUID",
                "generated fresh GUID",
            )
            changed = True

    if not changed:
        return None
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


# ---------------------------------------------------------------------------
# Package I/O
# ---------------------------------------------------------------------------


def _target_parts(zf: zipfile.ZipFile) -> list[str]:
    """Return the slideMaster and slideLayout XML parts to repair."""
    return [
        n
        for n in zf.namelist()
        if (n.startswith("ppt/slideMasters/") and n.endswith(".xml") and "_rels" not in n)
        or (n.startswith("ppt/slideLayouts/") and n.endswith(".xml") and "_rels" not in n)
    ]


def repair_template(
    input_path: str | Path,
    output_path: str | Path | None = None,
) -> RepairReport:
    """Repair a template PPTX in place (or into `output_path` if given).

    Returns a report of what was fixed. If nothing was defective, the file
    is still copied/rewritten but the report has no entries.
    """
    input_path = Path(input_path)
    if output_path is None:
        output_path = input_path
    else:
        output_path = Path(output_path)

    report = RepairReport(file=str(input_path.name))

    # Read all parts into memory (small enough for a template), apply
    # repairs, write out fresh zip.
    parts: dict[str, bytes] = {}
    with zipfile.ZipFile(input_path, "r") as zf:
        for name in zf.namelist():
            parts[name] = zf.read(name)
        target_names = _target_parts(zf)

    for name in target_names:
        try:
            repaired = _repair_placeholders_in_part(parts[name], name, report)
        except etree.XMLSyntaxError as exc:
            report.add(name, f"XML parse error ({exc})", "left unchanged")
            continue
        if repaired is not None:
            parts[name] = repaired

    # Write repaired package. Use ZIP_DEFLATED to match PowerPoint's
    # default. Preserve file order.
    tmp_out = output_path.with_suffix(output_path.suffix + ".tmp")
    with zipfile.ZipFile(tmp_out, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in parts.items():
            zf.writestr(name, data)
    tmp_out.replace(output_path)

    return report


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import argparse
    import json

    ap = argparse.ArgumentParser(description="Auto-repair a template PPTX file.")
    ap.add_argument("--input", required=True, help="path to template pptx")
    ap.add_argument("--output", default=None, help="write repaired pptx here (default: in place)")
    ap.add_argument("--report", default=None, help="write JSON report to this path")
    args = ap.parse_args()

    report = repair_template(args.input, args.output)
    print(f"Repaired {report.file}: {report.count} fix(es)")
    for fix in report.fixes:
        print(f"  [{fix['part']}] {fix['defect']} → {fix['resolution']}")
    if args.report:
        Path(args.report).write_text(json.dumps({"file": report.file, "fixes": report.fixes}, indent=2))
        print(f"Report → {args.report}")
