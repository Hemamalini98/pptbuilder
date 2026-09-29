"""Parse client-supplied alt-text Excel workbooks into normalized entries.

The frontend accepts an accessibility Excel that different customers ship in
similar-but-not-identical shapes. This parser is deliberately tolerant:

* Picks the sheet whose name contains "import" first, then falls back to the
  first sheet whose header row exposes the fields we care about.
* Locates columns by header name (case-insensitive substring match) so the
  Excel author can reorder columns without breaking us.
* Normalizes chapter number and element number into a lowercase
  ``figure_key`` (e.g. ``figure 2.1``, ``table 10.3``, ``unnumbered 4.2a``)
  that the frontend compares against its own normalized figure names.

Rows that cannot be mapped to a numeric-chapter + numeric-element pair are
skipped. Front matter, back matter, and other non-numeric chapter prefixes
are preserved literally (``figure fm.1``) so downstream can still surface
them if a figure with a matching name exists.
"""
from __future__ import annotations

import re
import warnings
from typing import Any

import openpyxl


# Usage code (Excel column E) → figure_key prefix used by the frontend.
# Only the three types that carry a numeric element identifier are mapped;
# every other usage (design element, cover, front matter feature) still
# yields an entry when its element number is parseable but keeps the raw
# usage string as its prefix so nothing is silently discarded.
_USAGE_TO_KIND = {
    "FIG": "figure",
    "TAB": "table",
    "UNN": "unnumbered",
}

_HEADER_TOKENS = {
    "chapter": ("chapter number", "chapter no", "chapter"),
    "element": ("element number", "element no", "element"),
    "usage": ("usage",),
    "decorative": ("decorative",),
    "alt_short": ("alt text - short", "alt text short", "alt short", "short alt"),
    "alt_long": ("alt text - long", "alt text long", "alt long", "long alt"),
}


def _norm_header(cell: Any) -> str:
    if cell is None:
        return ""
    return str(cell).strip().lower()


def _find_header_row(rows: list[list[Any]]) -> tuple[int, dict[str, int]] | None:
    """Return (row_index, column_map) for the first row that looks like a header.

    A row qualifies when it exposes at least a chapter, an element, and an
    alt-text-short column — the three fields we need to build a usable entry.

    Column matching is two-pass so specific tokens beat generic ones: the
    exact string "chapter number" wins the ``chapter`` slot even when
    "Chapter Prefix" is also present (which the loose "chapter" substring
    would otherwise capture first). Substring fallback only fires when no
    header cell matched a token exactly.
    """
    for r_idx, row in enumerate(rows):
        headers = [_norm_header(v) for v in row]
        col_map: dict[str, int] = {}
        used_columns: set[int] = set()
        # Pass 1: exact matches (specific > generic thanks to token order).
        for key, tokens in _HEADER_TOKENS.items():
            for tok in tokens:
                for c_idx, h in enumerate(headers):
                    if c_idx in used_columns:
                        continue
                    if tok == h:
                        col_map[key] = c_idx
                        used_columns.add(c_idx)
                        break
                if key in col_map:
                    break
        # Pass 2: substring fallback for keys still unresolved.
        for key, tokens in _HEADER_TOKENS.items():
            if key in col_map:
                continue
            for tok in tokens:
                for c_idx, h in enumerate(headers):
                    if c_idx in used_columns:
                        continue
                    if tok in h:
                        col_map[key] = c_idx
                        used_columns.add(c_idx)
                        break
                if key in col_map:
                    break
        if all(k in col_map for k in ("chapter", "element", "alt_short")):
            return r_idx, col_map
    return None


def _normalize_chapter(raw: Any) -> str | None:
    """Return the chapter identifier as a lowercase string, or None if blank.

    Numeric chapters lose leading zeros (``02`` → ``2``) so keys align with
    figure names produced elsewhere in the app. Non-numeric chapter codes
    (``FM``, ``AP``, ``APA``) are lowercased and kept literal — the entry is
    still emitted so the caller can surface it if a matching figure exists.
    """
    if raw is None:
        return None
    s = str(raw).strip()
    if not s:
        return None
    if s.isdigit():
        return str(int(s))
    return s.lower()


_ELEMENT_RE = re.compile(r"(\d+)([a-zA-Z]*)")


def _normalize_element(raw: Any) -> str | None:
    """Extract the element identifier from an Excel value.

    Strings like ``F01``, ``T02``, ``UN03``, ``FM11`` yield ``1``, ``2``,
    ``3``, ``11``. Element parts (``F01A``, ``F01B``) yield ``1a``, ``1b``.
    Returns None when no digit is present.
    """
    if raw is None:
        return None
    s = str(raw).strip()
    if not s:
        return None
    m = _ELEMENT_RE.search(s)
    if not m:
        return None
    number = str(int(m.group(1)))
    suffix = (m.group(2) or "").lower()
    return f"{number}{suffix}"


def _kind_for_row(usage_raw: Any, element_raw: Any) -> str | None:
    """Choose the figure_key prefix for a row.

    Priority: usage column (FIG/TAB/UNN → figure/table/unnumbered). When the
    usage cell is blank we infer from the element-number prefix (``F`` → figure,
    ``T`` → table, ``UN`` → unnumbered) so decks that ship a sparser Excel
    still produce usable keys.
    """
    if usage_raw is not None:
        usage = str(usage_raw).strip().upper()
        if usage in _USAGE_TO_KIND:
            return _USAGE_TO_KIND[usage]
        if usage:
            return usage.lower()
    if element_raw is not None:
        s = str(element_raw).strip().upper()
        if s.startswith("UN"):
            return "unnumbered"
        if s.startswith("F"):
            return "figure"
        if s.startswith("T"):
            return "table"
    return None


def _decorative_from_cell(raw: Any) -> bool:
    if raw is None:
        return False
    return str(raw).strip().lower() in ("yes", "y", "true", "1")


def _pick_sheet(wb) -> Any:
    """Return the preferred sheet: first name containing 'import', else first sheet."""
    for name in wb.sheetnames:
        if "import" in name.lower():
            return wb[name]
    return wb[wb.sheetnames[0]]


def parse_alttext_excel(path: str) -> list[dict]:
    """Parse the alt-text workbook at ``path`` and return normalized entries.

    Each returned dict:
        {
          "figure_key":     "figure 2.1",
          "element":        "F01",      # raw as it appeared in the sheet
          "chapter":        "02",       # raw as it appeared in the sheet
          "decorative":     False,
          "alt_text_short": "…",
          "alt_text_long":  "…",
        }

    Rows without a chapter/element/short-alt are skipped silently — the same
    Excel usually carries instruction/legend rows the customer doesn't want
    surfaced.
    """
    with warnings.catch_warnings():
        # openpyxl warns about unknown extensions on customer-shipped files;
        # they don't affect the values we read.
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)

    ws = _pick_sheet(wb)
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    header = _find_header_row(rows)
    if header is None:
        return []
    header_idx, cols = header

    def cell(row: list[Any], key: str) -> Any:
        idx = cols.get(key)
        if idx is None or idx >= len(row):
            return None
        return row[idx]

    entries: list[dict] = []
    seen_keys: set[str] = set()
    for row in rows[header_idx + 1:]:
        chapter_raw = cell(row, "chapter")
        element_raw = cell(row, "element")
        alt_short_raw = cell(row, "alt_short")
        alt_long_raw = cell(row, "alt_long")
        usage_raw = cell(row, "usage")
        decorative_raw = cell(row, "decorative")

        # A row must have at least chapter + element to build a figure_key.
        # The short-alt is what makes the entry useful; blank rows are skipped.
        chapter = _normalize_chapter(chapter_raw)
        element = _normalize_element(element_raw)
        if chapter is None or element is None:
            continue
        kind = _kind_for_row(usage_raw, element_raw)
        if kind is None:
            continue

        alt_short = (str(alt_short_raw).strip() if alt_short_raw is not None else "")
        alt_long = (str(alt_long_raw).strip() if alt_long_raw is not None else "")
        if not alt_short and not alt_long:
            continue

        figure_key = f"{kind} {chapter}.{element}"
        # De-duplicate: first non-blank row per key wins. Customer Excels
        # sometimes repeat an element on the "Examples" and "For Import" tabs.
        if figure_key in seen_keys:
            continue
        seen_keys.add(figure_key)

        entries.append({
            "figure_key": figure_key,
            "element": str(element_raw).strip() if element_raw is not None else "",
            "chapter": str(chapter_raw).strip() if chapter_raw is not None else "",
            "decorative": _decorative_from_cell(decorative_raw),
            "alt_text_short": alt_short,
            "alt_text_long": alt_long,
        })

    return entries
