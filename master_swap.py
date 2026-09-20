"""Master-swap PPTX converter.

Rebuilds an input deck inside a fresh copy of a template so the output inherits
the template's slide master, layouts, and theme (unlike convert.py which
restyles shapes in place and keeps the input's master).

Selection of a template layout per input slide is driven by rules_default.json.

Contract:
    apply_template(input_path, template_path, output_path, rules_path=None) -> Path
"""
from __future__ import annotations

import copy
import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.slide import Slide, SlideLayout
from pptx.util import Emu, Pt

FIG_MARKER_RE = re.compile(r"insert\s+(figure|table|unnumbered)", re.IGNORECASE)


# ---------------------------------------------------------------------------
# Input analysis
# ---------------------------------------------------------------------------

@dataclass
class SlideSignature:
    idx: int  # 1-based
    input_layout_name: str
    title: str
    title_is_upper: bool
    body_line_count: int
    subheading_hits: int
    has_figure_marker: bool
    image_count: int
    body_paragraphs: list[dict]  # see _extract_paragraph
    notes_text: str
    title_size_pt: float | None = None  # input's title font size


def _extract_run(run) -> dict:
    """Capture the run-level properties that carry over into master swap.

    Only *semantic* properties are preserved:
    - bold / italic / underline (emphasis)
    - RGB color (only when explicitly set — markers with FF0000, etc.)

    Font size and font family are intentionally NOT captured: master swap
    means the template's master/layout drives typography. Preserving the
    input's `Arial 24pt` would force those values onto every output run and
    override the template's Archivo Medium at its master-defined size.
    """
    f = run.font
    out: dict = {"text": run.text}
    if f.bold is not None:
        out["bold"] = bool(f.bold)
    if f.italic is not None:
        out["italic"] = bool(f.italic)
    if f.underline is not None:
        try:
            out["underline"] = bool(f.underline)
        except Exception:
            pass
    try:
        if f.color and f.color.type is not None and hasattr(f.color, "rgb") and f.color.rgb:
            out["color_rgb"] = str(f.color.rgb)
    except Exception:
        pass
    return out


def _paragraph_has_no_bullet(p) -> bool:
    """Return True if the paragraph's pPr explicitly requests no bullet."""
    pPr = p._pPr
    if pPr is None:
        return False
    return pPr.find(qn("a:buNone")) is not None


def _extract_paragraph(p) -> dict:
    """Capture paragraph structure with per-run properties for master swap.

    Returned dict: {level, has_no_bullet, runs: [{text, size_pt, bold, ...}, ...]}
    """
    runs = [_extract_run(r) for r in p.runs]
    return {
        "level": p.level or 0,
        "has_no_bullet": _paragraph_has_no_bullet(p),
        "runs": runs,
    }


def _para_text(para: dict) -> str:
    """Full text of a captured paragraph, across all runs."""
    return "".join(r.get("text", "") for r in para.get("runs", []))


def _para_is_all_bold(para: dict) -> bool:
    runs = [r for r in para.get("runs", []) if r.get("text", "").strip()]
    return bool(runs) and all(r.get("bold") is True for r in runs)


def analyze_input_slide(slide: Slide, idx: int) -> SlideSignature:
    title = ""
    title_size_pt: float | None = None
    for ph in slide.placeholders:
        try:
            if ph.placeholder_format.idx == 0 and ph.has_text_frame:
                title = ph.text_frame.text.strip()
                # Capture the input title's font size from its first run so we
                # can re-apply it in the output — the template's layout title
                # placeholder defaults to a very large font (Archivo Black
                # ~48pt) that wraps long titles. Preserving the input's size
                # keeps titles on one line and matches the human's manual work.
                for para in ph.text_frame.paragraphs:
                    for r in para.runs:
                        if r.font.size is not None:
                            try:
                                title_size_pt = float(r.font.size.pt)
                            except Exception:
                                pass
                            break
                    if title_size_pt is not None:
                        break
                break
        except Exception:
            pass

    body_paras: list[dict] = []
    subheading_hits = 0
    has_fig = False
    image_count = 0

    # Iterate shapes in reading order (top-to-bottom, then left-to-right)
    # so paragraphs that appear above others on the slide land earlier in
    # the body list. Input decks frequently put a heading like
    # "YOU LEARNED" in a separate placeholder above the bullet content;
    # z-order sorting alone would preserve creation order, not layout.
    def _shape_reading_key(sh):
        top = sh.top if sh.top is not None else 0
        left = sh.left if sh.left is not None else 0
        return (top, left)

    ordered_shapes = sorted(slide.shapes, key=_shape_reading_key)

    for shape in ordered_shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            image_count += 1
        if not shape.has_text_frame:
            continue
        # All non-title, non-marker paragraphs go into body_paras — including
        # paragraphs from custom shapes (Rectangle, TextBox). Input decks
        # frequently place bullet lists inside Rectangle shapes rather than
        # in the content placeholder itself, so treating them as loose
        # would drop the main body content on slides 20, 34, 35, etc.
        for p in shape.text_frame.paragraphs:
            txt = p.text.strip()
            if not txt or txt == title:
                continue
            para = _extract_paragraph(p)
            if FIG_MARKER_RE.search(txt):
                # Figure markers are collected here; the fill step splits
                # them out into their own red TextBox.
                body_paras.append(para)
                has_fig = True
                continue
            body_paras.append(para)
            if _para_is_all_bold(para) and len(txt) < 60 and not txt.endswith("."):
                subheading_hits += 1

    notes_text = ""
    if slide.has_notes_slide:
        try:
            notes_text = slide.notes_slide.notes_text_frame.text
        except Exception:
            notes_text = ""

    return SlideSignature(
        idx=idx,
        input_layout_name=slide.slide_layout.name,
        title=title,
        title_is_upper=bool(title)
        and title == title.upper()
        and any(c.isalpha() for c in title),
        body_line_count=len(body_paras),
        subheading_hits=subheading_hits,
        has_figure_marker=has_fig,
        image_count=image_count,
        body_paragraphs=body_paras,
        notes_text=notes_text,
        title_size_pt=title_size_pt,
    )


# ---------------------------------------------------------------------------
# Rule engine
# ---------------------------------------------------------------------------

def _rule_matches(when: dict, sig: SlideSignature) -> bool:
    for key, expected in when.items():
        if key == "input_layout":
            if sig.input_layout_name != expected:
                return False
        elif key == "title_is_upper":
            if sig.title_is_upper != expected:
                return False
        elif key == "has_figure_marker":
            if sig.has_figure_marker != expected:
                return False
        elif key == "body_line_count_lte":
            if sig.body_line_count > expected:
                return False
        elif key == "body_line_count_gte":
            if sig.body_line_count < expected:
                return False
        elif key == "subheading_hits_gte":
            if sig.subheading_hits < expected:
                return False
        elif key == "image_count_gte":
            if sig.image_count < expected:
                return False
        else:
            raise ValueError(f"unknown rule term {key!r}")
    return True


def _pick_layouts(rules: dict, sig: SlideSignature) -> list[str]:
    for rule in rules.get("rules", []):
        if _rule_matches(rule.get("when", {}), sig):
            action = rule.get("action")
            if action == "single":
                return [rule["layout"]]
            if action == "split_title_slide":
                return list(rule["layouts"])
            raise ValueError(f"unknown action {action!r}")
    raise LookupError(
        f"no rule matched input slide {sig.idx}: layout={sig.input_layout_name!r}"
    )


def _layout_by_name(prs, name: str) -> SlideLayout:
    for L in prs.slide_layouts:
        if L.name == name:
            return L
    raise LookupError(
        f"template has no layout named {name!r}; available: {[L.name for L in prs.slide_layouts]}"
    )


# ---------------------------------------------------------------------------
# Slide construction
# ---------------------------------------------------------------------------

def _delete_all_slides(prs) -> None:
    """Remove every slide from a presentation (leaves master + layouts intact)."""
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        rid = sldId.get(qn("r:id"))
        sldIdLst.remove(sldId)
        prs.part.drop_rel(rid)


def _clear_placeholder_text(ph) -> None:
    """Reset a placeholder's text frame to one empty paragraph, keeping template styling."""
    if not ph.has_text_frame:
        return
    tf = ph.text_frame
    txBody = tf._txBody
    for p in txBody.findall(qn("a:p")):
        txBody.remove(p)
    p = etree.SubElement(txBody, qn("a:p"))


def _enable_autofit(
    txBody,
    *,
    zero_insets: bool = False,
    anchor_center: bool = False,
    font_scale_pct: int | None = None,
    ln_spc_reduction_pct: int | None = None,
) -> None:
    """Force `<a:normAutofit/>` on a txBody's bodyPr.

    Template layouts declare `<a:noAutofit/>` on title placeholders, which
    lets long titles overflow / wrap. Overriding with normAutofit tells
    PowerPoint to shrink the text to fit.

    - `zero_insets` sets lIns/tIns/rIns/bIns to 0 (titles use this).
    - `anchor_center` sets `anchor="ctr"` so text centers vertically.
    - `font_scale_pct` / `ln_spc_reduction_pct` (0-100) emit
      `<a:normAutofit fontScale="XXXXX" lnSpcReduction="XXXXX"/>`. Needed
      for titles when we know PowerPoint's auto-shrink won't run on
      file-open — an explicit value forces the shrink deterministically.
      A value of 80 means 80% of the inherited font size.
    """
    bodyPr = txBody.find(qn("a:bodyPr"))
    if bodyPr is None:
        bodyPr = etree.SubElement(txBody, qn("a:bodyPr"))
        txBody.insert(0, bodyPr)
    if zero_insets:
        for attr in ("lIns", "tIns", "rIns", "bIns"):
            bodyPr.set(attr, "0")
    if anchor_center:
        bodyPr.set("anchor", "ctr")
    # remove any existing autofit setting
    for tag in ("a:noAutofit", "a:normAutofit", "a:spAutoFit"):
        for existing in bodyPr.findall(qn(tag)):
            bodyPr.remove(existing)
    autofit = etree.SubElement(bodyPr, qn("a:normAutofit"))
    if font_scale_pct is not None:
        autofit.set("fontScale", str(int(font_scale_pct) * 1000))
    if ln_spc_reduction_pct is not None:
        autofit.set("lnSpcReduction", str(int(ln_spc_reduction_pct) * 1000))


def _title_font_scale_pct(text: str) -> int | None:
    """Pick a fontScale% for a title based on character count.

    Master title defRPr is 40pt Archivo Black (all-caps). Long strings
    overflow a 12.5"-wide title placeholder even at zero padding. These
    thresholds match what the human's manual work produces when saving
    long titles that don't fit at 40pt.
    """
    n = len(text or "")
    if n <= 25:
        return None       # fits at 40pt, no shrink
    if n <= 35:
        return 85         # ~34pt
    if n <= 45:
        return 70         # ~28pt (matches expected slide 6's rendered size)
    return 60             # ~24pt for very long titles


def _layout_has_dark_bg(layout) -> bool:
    """Detect whether a layout has a dark background (MidnightSky-style).

    Uses the layout's `<p:clrMapOvr><a:overrideClrMapping/>` — if bg1 is
    mapped to dk1 (dark), the layout has a dark background. This is
    template-agnostic; any dark-background layout gets the same treatment.
    """
    try:
        el = layout.element
        clrMapOvr = el.find(qn("p:clrMapOvr"))
        if clrMapOvr is None:
            return False
        override = clrMapOvr.find(qn("a:overrideClrMapping"))
        if override is None:
            return False
        return override.get("bg1", "") == "dk1"
    except Exception:
        return False


def _set_placeholder_text(
    ph,
    text: str,
    *,
    autofit: bool = False,
    size_pt: float | None = None,
    is_title: bool = False,
    title_color_scheme: str | None = None,
) -> None:
    """Write a single line of text into a placeholder.

    - `autofit=True` forces `<a:normAutofit/>` so PowerPoint shrinks the
      text if it overflows.
    - `size_pt` emits an explicit `sz` on the run.
    - `is_title=True` applies the template's title conventions: zeroes the
      body-frame insets so long titles have room to fit or shrink.
    - `title_color_scheme` picks the schemeClr for title runs:
      * `"tx2"` (default for light layouts) — resolves to dk2 (dark navy)
        via the master's color map.
      * `"tx1"` — resolves to lt1 (white) on layouts whose
        `overrideClrMapping` swaps text/background (MidnightSky variants).
    """
    _clear_placeholder_text(ph)
    txBody = ph.text_frame._txBody
    if autofit:
        font_scale = _title_font_scale_pct(text) if is_title else None
        # Do NOT combine lnSpcReduction with the master's already-tight 72%
        # line spacing — the two multiply and cause wrapped title lines to
        # visually overlap. fontScale alone gives the shrink we need.
        _enable_autofit(
            txBody,
            zero_insets=is_title,
            font_scale_pct=font_scale,
        )
    p = txBody.find(qn("a:p"))
    if text:
        r = etree.SubElement(p, qn("a:r"))
        rPr = etree.SubElement(r, qn("a:rPr"))
        rPr.set("lang", "en-US")
        rPr.set("dirty", "0")
        if size_pt is not None:
            rPr.set("sz", str(int(round(float(size_pt) * 100))))
        if is_title and title_color_scheme:
            fill = etree.SubElement(rPr, qn("a:solidFill"))
            scheme = etree.SubElement(fill, qn("a:schemeClr"))
            scheme.set("val", title_color_scheme)
        t = etree.SubElement(r, qn("a:t"))
        t.text = text


def _apply_bullet(pPr, level: int, *, style: str = "dot") -> None:
    """Emit a standard bullet on a pPr element.

    Required because the NASM template master defines `<a:buNone/>` at every
    body level, so paragraphs inherit "no bullet" unless we override.
    marL / indent (0.375") produce PowerPoint's default hanging bullet.
    Level shifts the margin so nested bullets indent visibly.

    style="dot" → Arial round bullet `•`  (used by 2-column and image layouts)
    style="wing" → Wingdings section symbol `§` (used by 1-column layouts)
    """
    base_marL = 342900  # 0.375" — PowerPoint's standard hanging bullet indent
    per_level_step = 342900
    pPr.set("marL", str(base_marL + level * per_level_step))
    pPr.set("indent", "-342900")
    buFont = etree.SubElement(pPr, qn("a:buFont"))
    buChar = etree.SubElement(pPr, qn("a:buChar"))
    if style == "wing":
        buFont.set("typeface", "Wingdings")
        buFont.set("panose", "05000000000000000000")
        buFont.set("pitchFamily", "2")
        buFont.set("charset", "2")
        buChar.set("char", "§")
    else:
        buFont.set("typeface", "Arial")
        buFont.set("panose", "020B0604020202020204")
        buFont.set("pitchFamily", "34")
        buFont.set("charset", "0")
        buChar.set("char", "•")


def _write_run(p_elem, run: dict, *, force_body_color: bool = False) -> None:
    """Emit an <a:r> element from a captured run dict.

    - Preserves size (sz), bold (b), italic (i), underline, and rgb color.
    - Font family is preserved only if the input had it explicitly.
    - When `force_body_color` is True and no run color is set, injects
      `<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>` so body text
      renders as the template's dark text color, overriding the master's
      tx2 default.
    - Emits `b="0"` explicitly for runs where the input's bold was absent
      or False, since the template master's body defRPr has `b="1"` — we
      must override to avoid every bullet becoming bold.
    """
    r = etree.SubElement(p_elem, qn("a:r"))
    rPr = etree.SubElement(r, qn("a:rPr"))
    rPr.set("lang", "en-US")
    rPr.set("dirty", "0")
    bold = run.get("bold")
    if bold is True:
        # Let the master defRPr's b="1" drive — omit explicit b so the run
        # inherits (matches the expected file's XML pattern).
        pass
    elif force_body_color:
        # Override the master's body defRPr b="1" so plain body text is
        # not bold. Emitting b="0" explicitly is what the expected file
        # does; without it, every bullet renders bold because the master
        # body style is bold by default.
        rPr.set("b", "0")
    if run.get("italic") is True:
        rPr.set("i", "1")
    elif run.get("italic") is False:
        rPr.set("i", "0")
    if run.get("underline") is True:
        rPr.set("u", "sng")
    # Color: only emit if the input had an explicit RGB (e.g. FF0000
    # markers). For body content, emit tx1 explicitly so the template's
    # master body color (tx2 = dark blue) doesn't kick in.
    if run.get("color_rgb"):
        fill = etree.SubElement(rPr, qn("a:solidFill"))
        clr = etree.SubElement(fill, qn("a:srgbClr"))
        clr.set("val", run["color_rgb"])
    elif force_body_color:
        fill = etree.SubElement(rPr, qn("a:solidFill"))
        scheme = etree.SubElement(fill, qn("a:schemeClr"))
        scheme.set("val", "tx1")
    # Font family and size are NOT re-emitted: master swap wants the
    # template's theme font at the master-defined size to drive typography.
    t = etree.SubElement(r, qn("a:t"))
    t.text = run.get("text", "")


def _write_paragraphs(
    ph,
    paragraphs: Iterable[dict],
    *,
    is_body: bool = True,
    bullet_style: str = "dot",
) -> None:
    """Write captured paragraphs into a placeholder.

    Each paragraph is: {level, has_no_bullet, runs: [{text, size_pt, bold, ...}, ...]}
    When `is_body` is True, emits a bullet on every paragraph unless the
    input paragraph explicitly declared `<a:buNone/>`, forces the template's
    tx1 body color on runs without an explicit color, and enables
    `<a:normAutofit/>` + `anchor="ctr"` so text shrinks and centers within
    the placeholder — matching how expected slides handle body content.
    """
    _clear_placeholder_text(ph)
    txBody = ph.text_frame._txBody
    if is_body:
        _enable_autofit(txBody, anchor_center=True)
    existing = txBody.find(qn("a:p"))
    if existing is not None:
        txBody.remove(existing)
    for para in paragraphs:
        p = etree.SubElement(txBody, qn("a:p"))
        pPr = etree.SubElement(p, qn("a:pPr"))
        level = para.get("level", 0) or 0
        if level:
            pPr.set("lvl", str(level))
        want_bullet = (
            is_body
            and not para.get("has_no_bullet", False)
            and any(r.get("text", "").strip() for r in para.get("runs", []))
        )
        if want_bullet:
            _apply_bullet(pPr, level, style=bullet_style)
        # If we set nothing on pPr, drop it so PowerPoint uses layout defaults.
        if not pPr.attrib and len(pPr) == 0:
            p.remove(pPr)
        for run in para.get("runs", []):
            _write_run(p, run, force_body_color=is_body)


def _find_placeholder(slide: Slide, *, idx: int | None = None, type_name: str | None = None):
    """Find a placeholder on a slide by idx or type."""
    for ph in slide.placeholders:
        try:
            fmt = ph.placeholder_format
            if idx is not None and fmt.idx == idx:
                return ph
            if type_name is not None and str(fmt.type).endswith(type_name):
                return ph
        except Exception:
            continue
    return None


def _find_title_placeholder(slide: Slide):
    for name in ("TITLE", "CENTER_TITLE"):
        ph = _find_placeholder(slide, type_name=name)
        if ph is not None:
            return ph
    return _find_placeholder(slide, idx=0)


def _find_body_placeholders(slide: Slide) -> list:
    """Return all body/object placeholders on the slide, sorted by idx."""
    out = []
    for ph in slide.placeholders:
        try:
            fmt = ph.placeholder_format
            if fmt.idx == 0:
                continue
            type_str = str(fmt.type)
            if any(k in type_str for k in ("BODY", "OBJECT", "SUBTITLE")):
                out.append((fmt.idx, ph))
        except Exception:
            continue
    return [ph for _, ph in sorted(out, key=lambda t: t[0])]


def _split_out_figure_markers(
    body_paragraphs: list[dict],
) -> tuple[list[dict], list[dict]]:
    """Separate `<Insert Figure/Table/…>` marker paragraphs from real body content.

    Markers are directive TextBoxes ("<Insert Table 2.1: …>") that the input
    keeps in a dedicated shape. When we flatten the input, they end up in the
    same paragraph list as the real bullets; the human-produced output keeps
    them as a standalone red TextBox on the slide, so we must extract them
    before writing paragraphs into the body placeholder.
    """
    content, markers = [], []
    for p in body_paragraphs:
        if FIG_MARKER_RE.search(_para_text(p)):
            markers.append(p)
        else:
            content.append(p)
    return content, markers


def _target_area_for_marker(layout: SlideLayout) -> tuple[int, int, int, int] | None:
    """Return (left, top, width, height) of the area to drop marker TextBox in.

    Prefers layout's ph15 (image slot / right column). Falls back to ph12
    (body area) for single-column layouts that have no ph15.
    """
    ph15 = ph12 = None
    for lp in layout.placeholders:
        try:
            idx = lp.placeholder_format.idx
        except Exception:
            continue
        if idx == 15:
            ph15 = lp
        elif idx == 12:
            ph12 = lp
    target = ph15 or ph12
    if target is None:
        return None
    if None in (target.left, target.top, target.width, target.height):
        return None
    return target.left, target.top, target.width, target.height


def _add_marker_textboxes(
    slide: Slide,
    marker_paras: list[dict],
    layout: SlideLayout,
) -> None:
    """Emit each figure/table marker as its own red non-placeholder TextBox.

    Position: centered inside the layout's ph15 (image/right column) or, if
    no ph15 exists, centered inside ph12.
    """
    if not marker_paras:
        return
    area = _target_area_for_marker(layout)
    if area is None:
        return
    area_l, area_t, area_w, area_h = area

    # Box is narrower than the target area (mimics the hand-authored output,
    # which uses ~3.3-5.3" wide boxes inside a 6"-wide column).
    box_w = min(area_w, int(area_w * 0.60) + Emu(int(0.3 * 914400)))
    line_h = Emu(int(0.85 * 914400))
    total_h = line_h * len(marker_paras)
    box_left = area_l + (area_w - box_w) // 2
    box_top = area_t + max(0, (area_h - total_h) // 2)

    for i, p in enumerate(marker_paras):
        tb = slide.shapes.add_textbox(
            box_left, box_top + i * line_h, box_w, line_h
        )
        tb.text_frame.word_wrap = True
        para = tb.text_frame.paragraphs[0]
        run = para.add_run()
        run.text = _para_text(p)
        run.font.size = Pt(20)
        run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)


def _copy_notes(new_slide: Slide, notes_text: str) -> None:
    if not notes_text.strip():
        return
    ns = new_slide.notes_slide
    ns.notes_text_frame.text = notes_text


# ---------------------------------------------------------------------------
# Title-slide split (JBL chapter cover convention)
# ---------------------------------------------------------------------------

def _extract_cover_parts(sig: SlideSignature) -> tuple[str, str, str]:
    """From a Title-Slide input, produce (book_title, chapter_line, part_title).

    JBL convention: input title = book title, body[0] = 'Chapter N', body[1] = PART TITLE.
    Falls back gracefully if the structure differs.
    """
    book_title = sig.title
    chapter_line = ""
    part_title = ""
    body_lines = [_para_text(p).strip() for p in sig.body_paragraphs if _para_text(p).strip()]
    for line in body_lines:
        if re.match(r"^chapter\s+\d+", line, re.IGNORECASE):
            chapter_line = line
        elif line == line.upper() and len(line) > 3:
            part_title = line
    if not part_title and body_lines:
        part_title = body_lines[-1]
    return book_title, chapter_line, part_title


# ---------------------------------------------------------------------------
# Main entrypoint
# ---------------------------------------------------------------------------

def apply_template(
    input_path: str | Path,
    template_path: str | Path,
    output_path: str | Path,
    rules_path: str | Path | None = None,
) -> tuple[Path, list[dict]]:
    """Rebuild the input deck inside a fresh copy of the template.

    Returns (output_path, report) where report is a per-slide record of the
    layout chosen and rule that matched.
    """
    input_path = Path(input_path)
    template_path = Path(template_path)
    output_path = Path(output_path)
    if rules_path is None:
        rules_path = Path(__file__).parent / "rules_default.json"
    rules = json.loads(Path(rules_path).read_text())
    overrides = {int(k): v for k, v in rules.get("overrides_by_output_index", {}).items() if k.isdigit()}

    input_prs = Presentation(str(input_path))
    input_slides = list(input_prs.slides)

    # Build output presentation from a fresh copy of the template
    output_prs = Presentation(str(template_path))
    _delete_all_slides(output_prs)

    report: list[dict] = []
    output_slide_counter = 0

    for i, in_slide in enumerate(input_slides, start=1):
        sig = analyze_input_slide(in_slide, i)
        layout_names = _pick_layouts(rules, sig)

        if len(layout_names) == 2 and sig.input_layout_name == "Title Slide":
            book_title, chapter_line, part_title = _extract_cover_parts(sig)

            layout1 = _layout_by_name(output_prs, layout_names[0])
            s1 = output_prs.slides.add_slide(layout1)
            output_slide_counter += 1
            _emit_report(report, output_slide_counter, i, sig, layout_names[0], "split[0]")
            _fill_cover_slide(s1, book_title, chapter_line, sig.title_size_pt)

            layout2 = _layout_by_name(output_prs, layout_names[1])
            s2 = output_prs.slides.add_slide(layout2)
            output_slide_counter += 1
            _emit_report(report, output_slide_counter, i, sig, layout_names[1], "split[1]")
            _fill_cover_slide(s2, part_title, chapter_line, sig.title_size_pt)
            continue

        layout_name = layout_names[0]
        # Apply per-slide override if configured (matches on the *next* output index)
        prospective_idx = output_slide_counter + 1
        if prospective_idx in overrides:
            layout_name = overrides[prospective_idx]

        layout = _layout_by_name(output_prs, layout_name)
        new_slide = output_prs.slides.add_slide(layout)
        output_slide_counter += 1
        _emit_report(report, output_slide_counter, i, sig, layout_name, "rule" if prospective_idx not in overrides else "override")
        _fill_body_slide(new_slide, sig, layout_name)

    output_prs.save(str(output_path))
    return output_path, report


def _emit_report(report, out_idx, in_idx, sig, layout_name, reason):
    report.append({
        "out_idx": out_idx,
        "in_idx": in_idx,
        "input_layout": sig.input_layout_name,
        "title": sig.title[:70],
        "chosen_layout": layout_name,
        "reason": reason,
        "body_lines": sig.body_line_count,
        "subheadings": sig.subheading_hits,
        "figure_marker": sig.has_figure_marker,
    })


def _remove_placeholder(slide: Slide, ph) -> None:
    """Detach a placeholder shape from the slide's spTree.

    Used to drop body placeholders we did not populate (e.g. an unused ph15
    on a 2-column layout), so PowerPoint does not render its edit-mode
    "Click to add text" prompt.
    """
    try:
        sp = ph._element
        sp.getparent().remove(sp)
    except Exception:
        pass


def _fill_cover_slide(slide: Slide, main_text: str, chapter_line: str, title_size_pt: float | None = None) -> None:
    title_ph = _find_title_placeholder(slide)
    if title_ph is not None:
        color = "tx1" if _layout_has_dark_bg(slide.slide_layout) else "tx2"
        _set_placeholder_text(
            title_ph, main_text,
            autofit=True, is_title=True, title_color_scheme=color,
        )
    bodies = _find_body_placeholders(slide)
    if bodies and chapter_line:
        _set_placeholder_text(bodies[0], chapter_line)
    # Drop any body placeholder we didn't populate so PowerPoint's "Click to
    # add text" prompt doesn't appear in the exported deck.
    filled = {bodies[0]._element} if (bodies and chapter_line) else set()
    for ph in bodies:
        if ph._element not in filled:
            _remove_placeholder(slide, ph)


def _bullet_style_for_layout(layout_name: str) -> str:
    """Pick the bullet character to match the human's expected output.

    2-column and image-flavored layouts use Arial `•`; 1-column layouts
    (and their MidnightSky variants) use Wingdings `§`. Rule derived from
    the expected file's per-slide bullet definitions.
    """
    if "2 Column" in layout_name or "3 Column" in layout_name:
        return "dot"
    if "Image" in layout_name:
        return "dot"
    return "wing"


def _add_loose_textbox(slide: Slide, paragraphs: list[dict], layout: SlideLayout) -> None:
    """Place standalone TextBox paragraphs (quotes/callouts) into their own
    non-placeholder TextBox on the output slide.

    Positioned in the lower half of the layout's ph15 area (right column /
    image slot) so it sits below any figure marker. Preserves the input's
    per-run font size and bold so the quote stays visually distinct from
    bullets.
    """
    if not paragraphs:
        return
    area = _target_area_for_marker(layout)
    if area is None:
        return
    area_l, area_t, area_w, area_h = area
    # Use most of the ph15 area — the human's manual output uses a similar
    # size for quote/callout boxes (~5" × 2").
    tb_left = area_l
    tb_top = area_t + area_h // 3
    tb_w = area_w
    tb_h = area_h - area_h // 3
    tb = slide.shapes.add_textbox(tb_left, tb_top, tb_w, tb_h)
    tb.text_frame.word_wrap = True
    # Reuse _write_paragraphs to preserve per-run styling; no bullet.
    for para in paragraphs:
        para["has_no_bullet"] = True
    _write_paragraphs(tb, paragraphs, is_body=True, bullet_style="dot")


def _copy_geometry_from_layout(slide_ph, layout) -> None:
    """Copy the layout placeholder's `<a:xfrm>` onto the slide placeholder.

    Placeholders without an explicit `<a:xfrm>` inherit geometry from the
    layout at render time. Copying the geometry onto the slide gives
    PowerPoint concrete dimensions to compute normAutofit against —
    matching how the expected file's body placeholders are structured.
    """
    try:
        idx = slide_ph.placeholder_format.idx
    except Exception:
        return
    for lp in layout.placeholders:
        try:
            if lp.placeholder_format.idx != idx:
                continue
        except Exception:
            continue
        layout_sp = lp._element
        layout_xfrm = layout_sp.find(qn("p:spPr") + "/" + qn("a:xfrm"))
        if layout_xfrm is None:
            return
        slide_sp = slide_ph._element
        slide_spPr = slide_sp.find(qn("p:spPr"))
        if slide_spPr is None:
            return
        # Remove any existing xfrm on the slide
        for existing in slide_spPr.findall(qn("a:xfrm")):
            slide_spPr.remove(existing)
        # Deep-copy the layout's xfrm as the first child of spPr
        new_xfrm = copy.deepcopy(layout_xfrm)
        slide_spPr.insert(0, new_xfrm)
        return


def _fill_body_slide(slide: Slide, sig: SlideSignature, layout_name: str) -> None:
    title_ph = _find_title_placeholder(slide)
    if title_ph is not None:
        color = "tx1" if _layout_has_dark_bg(slide.slide_layout) else "tx2"
        _set_placeholder_text(
            title_ph, sig.title,
            autofit=True, is_title=True, title_color_scheme=color,
        )

    # Peel off figure/table markers so they land in a dedicated red TextBox
    # instead of being lumped into the body placeholder.
    content_paras, marker_paras = _split_out_figure_markers(sig.body_paragraphs)

    bullet_style = _bullet_style_for_layout(layout_name)

    bodies = _find_body_placeholders(slide)
    filled_body: set = set()
    if bodies and content_paras:
        # All real body content goes into the primary body placeholder (ph12).
        # For "2 Column" and "1 Column + Image" layouts, ph15 is intentionally
        # left empty — that region is reserved for a figure/table (rendered
        # either as the image itself or as the red marker TextBox emitted
        # below). Splitting bullets across ph12/ph15 is not the convention
        # this template uses.
        _write_paragraphs(bodies[0], content_paras, bullet_style=bullet_style)
        # Give ph12 an explicit xfrm so PowerPoint's normAutofit measures the
        # exact box; matches the expected file's structure.
        _copy_geometry_from_layout(bodies[0], slide.slide_layout)
        filled_body.add(bodies[0]._element)

    # Drop any body placeholder we didn't fill so PowerPoint's edit-mode
    # "Click to add text" prompt does not appear (e.g. empty ph15 on
    # 2-column slides).
    for ph in bodies:
        if ph._element not in filled_body:
            _remove_placeholder(slide, ph)

    _add_marker_textboxes(slide, marker_paras, slide.slide_layout)
    _copy_notes(slide, sig.notes_text)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import argparse
    import sys

    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--template", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--rules", default=None)
    ap.add_argument("--report", default=None)
    args = ap.parse_args()

    out_path, report = apply_template(args.input, args.template, args.output, args.rules)
    print(f"wrote {out_path}")
    if args.report:
        Path(args.report).write_text(json.dumps(report, indent=2))
        print(f"report -> {args.report}")
    else:
        for r in report:
            print(
                f"  out{r['out_idx']:>2d} <- in{r['in_idx']:>2d}  "
                f"{r['input_layout']:20s} -> {r['chosen_layout']:28s} ({r['reason']})  "
                f"{r['title'][:60]}"
            )
