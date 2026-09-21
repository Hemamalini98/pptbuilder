"""Master-swap PPTX converter.

Rebuilds an input deck inside a fresh copy of a template so the output inherits
the template's slide master, layouts, and theme (unlike convert.py which
restyles shapes in place and keeps the input's master).

Layout selection is driven by scoring every template layout against the
source slide's placeholder signature — no hardcoded layout names, so the
same code works with any uploaded template.

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
# Generic layout matcher
#
# Instead of hardcoding rules by template layout name (which only work for
# one specific template), we score every template layout against the input
# slide's placeholder signature and pick the best-scoring one. This means
# the same code path works whether the uploaded template is NASM, Wolters
# Kluwer, or anything else — the template is the source of truth.
# ---------------------------------------------------------------------------

# Placeholder-type categories the matcher reasons about. Every template
# placeholder collapses into one of these, ignoring exact type strings
# so we can compare "OBJECT" ↔ "BODY" ↔ "TEXT" as equivalent body slots.
_CAT_TITLE = "title"
_CAT_BODY = "body"
_CAT_PICTURE = "picture"
_CAT_CHART = "chart"
_CAT_TABLE = "table"
_CAT_FOOTER = "footer"
_CAT_SLIDE_NUMBER = "sldnum"
_CAT_OTHER = "other"


def _categorize_placeholder_type(type_str: str) -> str:
    """Map a placeholder type string to a category for signature matching."""
    t = (type_str or "").upper()
    if "TITLE" in t:
        return _CAT_TITLE
    if any(k in t for k in ("BODY", "OBJECT", "SUBTITLE", "TEXT")):
        return _CAT_BODY
    if "PICTURE" in t or "MEDIA" in t or "CLIP_ART" in t:
        return _CAT_PICTURE
    if "CHART" in t:
        return _CAT_CHART
    if "TABLE" in t:
        return _CAT_TABLE
    if "FOOTER" in t:
        return _CAT_FOOTER
    if "SLIDE_NUMBER" in t or "SLIDENUM" in t:
        return _CAT_SLIDE_NUMBER
    return _CAT_OTHER


def _layout_placeholder_categories(layout) -> list[str]:
    """Return the categories of a template layout's placeholders."""
    cats: list[str] = []
    for ph in layout.placeholders:
        try:
            cats.append(_categorize_placeholder_type(str(ph.placeholder_format.type)))
        except Exception:
            continue
    return cats


def _layout_body_area_ratio(layout) -> float:
    """Return the fraction of the slide occupied by body-type placeholders.

    Content layouts (`1 Column`, `2 Column`, `Content`) devote most of
    the slide to body area. Title-slide / cover layouts (`Title Slide`,
    `1_Title`) have a small subtitle instead. Used to distinguish them
    when both look identical by category count.
    """
    try:
        prs = layout.part.package.presentation_part.presentation
        slide_area = prs.slide_width * prs.slide_height
    except Exception:
        return 0.0
    if not slide_area:
        return 0.0
    body_area = 0
    for ph in layout.placeholders:
        try:
            cat = _categorize_placeholder_type(str(ph.placeholder_format.type))
        except Exception:
            continue
        if cat not in (_CAT_BODY, _CAT_PICTURE, _CAT_CHART, _CAT_TABLE):
            continue
        if ph.width and ph.height:
            body_area += int(ph.width) * int(ph.height)
    return body_area / slide_area


def _layout_title_at_top(layout) -> bool:
    """True when the layout's title placeholder is in the upper third.

    Content layouts put the title near the top; cover/intro layouts
    place it lower on the slide.
    """
    try:
        prs = layout.part.package.presentation_part.presentation
        slide_h = prs.slide_height
    except Exception:
        return True
    for ph in layout.placeholders:
        try:
            if "TITLE" not in str(ph.placeholder_format.type):
                continue
        except Exception:
            continue
        if ph.top is None:
            return True
        return ph.top < slide_h * 0.33
    return True


def _score_layout_for_source(
    layout,
    source_cats: list[str],
    *,
    prefer_dark_bg: bool = False,
    prefer_content_layout: bool = True,
) -> int:
    """Score how well `layout` fits a source slide with `source_cats`.

    Higher is better. Ignores footer/slide-number placeholders (every
    layout has them; they're irrelevant to structural matching).
    """
    layout_cats = [
        c for c in _layout_placeholder_categories(layout)
        if c not in (_CAT_FOOTER, _CAT_SLIDE_NUMBER)
    ]
    source_relevant = [
        c for c in source_cats
        if c not in (_CAT_FOOTER, _CAT_SLIDE_NUMBER)
    ]

    score = 0
    from collections import Counter
    lc = Counter(layout_cats)
    sc = Counter(source_relevant)

    # Match each source category against layout's supply. Give bigger
    # weight to structural roles (title/body) than to accessories.
    weights = {
        _CAT_TITLE: 20,
        _CAT_BODY: 12,
        _CAT_PICTURE: 8,
        _CAT_CHART: 8,
        _CAT_TABLE: 6,
        _CAT_OTHER: 2,
    }
    for cat, count in sc.items():
        provided = lc.get(cat, 0)
        matched = min(count, provided)
        score += weights.get(cat, 1) * matched

    # Penalise mismatched surplus body slots — a source with one body
    # doesn't belong on a 3-column layout when a 1-column exists.
    surplus_body = max(0, lc.get(_CAT_BODY, 0) - sc.get(_CAT_BODY, 0))
    score -= 3 * surplus_body

    # Prefer layouts that match dark-background preference (for section
    # headers whose title is all-caps).
    if prefer_dark_bg and _layout_has_dark_bg(layout):
        score += 10
    elif not prefer_dark_bg and _layout_has_dark_bg(layout):
        score -= 5

    # Tie-breaker: prefer content-flavored layouts (title near top,
    # large body area) when the source has body content. Distinguishes
    # `1 Column` from `1_Title` etc. — both have title+body, but
    # `1 Column` devotes 55% of the slide to body while `1_Title` only
    # has a small subtitle.
    if prefer_content_layout and _CAT_BODY in source_relevant:
        if _layout_title_at_top(layout):
            score += 6
        body_ratio = _layout_body_area_ratio(layout)
        # Scale so a full-body layout adds ~6, a title-subtitle adds ~1
        score += int(body_ratio * 12)

    return score


def _pick_layout_for_slide(prs, sig: SlideSignature):
    """Score every template layout for the source slide and pick the best.

    Purely signature-driven — no hardcoded layout names. Works with any
    template the user uploads.
    """
    # Source placeholder categories:
    source_cats: list[str] = []
    if sig.title:
        source_cats.append(_CAT_TITLE)
    if sig.body_paragraphs:
        source_cats.append(_CAT_BODY)
    if sig.image_count > 0 or sig.has_figure_marker:
        source_cats.append(_CAT_PICTURE)

    # ALL-CAPS titles usually indicate a section header — bias toward
    # dark-background layouts when the template offers them.
    prefer_dark = sig.title_is_upper
    # Prefer content-flavored layouts unless the source is essentially
    # empty (just a title, no body — that's a real cover slide).
    prefer_content = bool(sig.body_paragraphs)

    best_layout = None
    best_score = None
    for layout in prs.slide_layouts:
        score = _score_layout_for_source(
            layout,
            source_cats,
            prefer_dark_bg=prefer_dark,
            prefer_content_layout=prefer_content,
        )
        if best_score is None or score > best_score:
            best_score = score
            best_layout = layout
    return best_layout


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


def _title_font_scale_pct(text: str, layout=None) -> int | None:
    """Pick a fontScale% for a title based on how much room the layout has.

    Uses the layout's title placeholder width and the master's title
    defRPr size to estimate whether the title needs to shrink to fit on
    one line. Purely template-driven: whatever font the master defines
    at whatever size, the estimate adapts.

    Returns None when no shrink is needed. Returns 60-95 when we need
    PowerPoint to render smaller at load time (the empty `<a:normAutofit/>`
    doesn't reliably trigger during headless export).
    """
    if not text or layout is None:
        return None
    # Locate the layout's title placeholder to read its width.
    title_width_emu = None
    for ph in layout.placeholders:
        try:
            if "TITLE" in str(ph.placeholder_format.type):
                title_width_emu = ph.width
                break
        except Exception:
            continue
    if title_width_emu is None or title_width_emu <= 0:
        return None
    title_width_in = title_width_emu / 914400

    # Read the master's title defRPr size (defaults to 40pt if unspecified).
    try:
        prs_part = layout.part.package.presentation_part
        prs = prs_part.presentation
    except Exception:
        prs = None
    title_size_pt = 40.0
    if prs is not None:
        for master in prs.slide_masters:
            try:
                txStyles = master.element.find(qn("p:txStyles"))
                titleStyle = txStyles.find(qn("p:titleStyle")) if txStyles is not None else None
                if titleStyle is None:
                    continue
                lvl1 = titleStyle.find(qn("a:lvl1pPr"))
                defRPr = lvl1.find(qn("a:defRPr")) if lvl1 is not None else None
                sz = defRPr.get("sz") if defRPr is not None else None
                if sz:
                    title_size_pt = float(sz) / 100.0
                break
            except Exception:
                continue

    # Rough char-width factor for a heavy display font (bold caps ~0.65em).
    approx_char_pt = title_size_pt * 0.65
    text_width_pt = len(text) * approx_char_pt
    available_pt = title_width_in * 72

    if text_width_pt <= available_pt:
        return None  # fits at native size

    scale = available_pt / text_width_pt
    pct = int(round(scale * 100))
    # Clamp so we don't shrink into unreadable territory.
    return max(50, min(95, pct))


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
    layout=None,
) -> None:
    """Write a single line of text into a placeholder.

    - `autofit=True` forces `<a:normAutofit/>` so PowerPoint shrinks the
      text if it overflows.
    - `size_pt` emits an explicit `sz` on the run.
    - `is_title=True` applies the template's title conventions: zeroes
      insets and, when `layout` is given, computes a fontScale from the
      layout's title-placeholder width vs. the master's title-size so
      long titles fit without hardcoding thresholds.
    - `title_color_scheme` picks the schemeClr for title runs — chosen
      by the caller based on the target layout's color map.
    """
    _clear_placeholder_text(ph)
    txBody = ph.text_frame._txBody
    if autofit:
        font_scale = (
            _title_font_scale_pct(text, layout=layout) if is_title else None
        )
        # No lnSpcReduction — multiplying with a tight master lnSpc causes
        # wrapped title lines to visually overlap.
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


@dataclass
class BulletDefinition:
    """A bullet inferred from the template (or a generic fallback)."""

    char: str
    font: str
    marL: int  # left margin at level 0, EMU
    indent: int  # first-line indent (usually negative for hanging)
    panose: str | None = None
    pitchFamily: str | None = None
    charset: str | None = None


# PowerPoint's default fallback bullet — used when the template's master
# and layout neither define a bullet nor inherit one. Kept as a *fallback*,
# not a hardcoded default: `_infer_bullet_from_template` is consulted first.
_DEFAULT_BULLET = BulletDefinition(
    char="•",
    font="Arial",
    marL=342900,
    indent=-342900,
    panose="020B0604020202020204",
    pitchFamily="34",
    charset="0",
)


def _infer_bullet_from_layout(layout, level: int = 0) -> BulletDefinition | None:
    """Read the level-N bullet from the layout's ph12 lstStyle if present.

    Only returns a bullet if the layout's placeholder explicitly defines
    one (buChar). Layouts that inherit `<a:buNone/>` from the master
    return None — the caller then decides whether to force a default.
    """
    for ph in layout.placeholders:
        try:
            fmt_idx = ph.placeholder_format.idx
        except Exception:
            continue
        if fmt_idx not in (12, 15, 1, 2):
            continue  # not a body-ish placeholder
        try:
            txBody = ph._element.find(qn("p:txBody"))
        except Exception:
            continue
        if txBody is None:
            continue
        lstStyle = txBody.find(qn("a:lstStyle"))
        if lstStyle is None:
            continue
        lvl_tag = f"a:lvl{level + 1}pPr"
        lvl = lstStyle.find(qn(lvl_tag))
        if lvl is None:
            continue
        buChar = lvl.find(qn("a:buChar"))
        if buChar is None:
            continue
        buFont = lvl.find(qn("a:buFont"))
        marL = int(lvl.get("marL", 342900))
        indent = int(lvl.get("indent", -342900))
        return BulletDefinition(
            char=buChar.get("char", "•"),
            font=buFont.get("typeface", "Arial") if buFont is not None else "Arial",
            marL=marL,
            indent=indent,
            panose=buFont.get("panose") if buFont is not None else None,
            pitchFamily=buFont.get("pitchFamily") if buFont is not None else None,
            charset=buFont.get("charset") if buFont is not None else None,
        )
    return None


def _infer_bullet_from_master(prs, level: int = 0) -> BulletDefinition | None:
    """Read the level-N body bullet from the master's txStyles bodyStyle."""
    for master in prs.slide_masters:
        try:
            txStyles = master.element.find(qn("p:txStyles"))
        except Exception:
            continue
        if txStyles is None:
            continue
        bodyStyle = txStyles.find(qn("p:bodyStyle"))
        if bodyStyle is None:
            continue
        lvl = bodyStyle.find(qn(f"a:lvl{level + 1}pPr"))
        if lvl is None:
            continue
        buChar = lvl.find(qn("a:buChar"))
        if buChar is None:
            continue
        buFont = lvl.find(qn("a:buFont"))
        marL = int(lvl.get("marL", 342900)) or 342900
        indent = int(lvl.get("indent", -342900)) or -342900
        return BulletDefinition(
            char=buChar.get("char", "•"),
            font=buFont.get("typeface", "Arial") if buFont is not None else "Arial",
            marL=marL,
            indent=indent,
            panose=buFont.get("panose") if buFont is not None else None,
            pitchFamily=buFont.get("pitchFamily") if buFont is not None else None,
            charset=buFont.get("charset") if buFont is not None else None,
        )
    return None


def _pick_bullet_for_slide(prs, layout, level: int = 0) -> BulletDefinition:
    """Pick the bullet definition to emit for a given layout + level.

    Tries in order: layout's own placeholder lstStyle → master's bodyStyle
    → generic Arial `•` fallback. This is fully template-driven: change
    the template and the bullet changes automatically.
    """
    return (
        _infer_bullet_from_layout(layout, level)
        or _infer_bullet_from_master(prs, level)
        or _DEFAULT_BULLET
    )


def _apply_bullet(pPr, level: int, bullet: BulletDefinition) -> None:
    """Emit a bullet on a pPr element using the given definition."""
    pPr.set("marL", str(bullet.marL + level * abs(bullet.marL)))
    pPr.set("indent", str(bullet.indent))
    buFont = etree.SubElement(pPr, qn("a:buFont"))
    buFont.set("typeface", bullet.font)
    if bullet.panose:
        buFont.set("panose", bullet.panose)
    if bullet.pitchFamily:
        buFont.set("pitchFamily", bullet.pitchFamily)
    if bullet.charset:
        buFont.set("charset", bullet.charset)
    buChar = etree.SubElement(pPr, qn("a:buChar"))
    buChar.set("char", bullet.char)


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
    bullet: "BulletDefinition | None" = None,
) -> None:
    """Write captured paragraphs into a placeholder.

    Each paragraph: {level, has_no_bullet, runs: [{text, bold, italic, ...}]}
    When `is_body` is True, emits `<a:normAutofit/>` + `anchor="ctr"` so
    body text shrinks and centers, and emits a bullet on every paragraph
    (unless the input declared `<a:buNone/>`) using the supplied
    `BulletDefinition`. The bullet definition is inferred from the target
    template, so different templates yield different bullet styles
    automatically.
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
            and bullet is not None
            and not para.get("has_no_bullet", False)
            and any(r.get("text", "").strip() for r in para.get("runs", []))
        )
        if want_bullet:
            _apply_bullet(pPr, level, bullet)
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
# Main entrypoint
# ---------------------------------------------------------------------------

def apply_template(
    input_path: str | Path,
    template_path: str | Path,
    output_path: str | Path,
    rules_path: str | Path | None = None,  # deprecated — kept for API compat
) -> tuple[Path, list[dict]]:
    """Rebuild the input deck inside a fresh copy of the template.

    Fully template-driven — no layout-name hardcoding. For each source slide
    we score every template layout against the source's placeholder
    signature (title/body/picture) and pick the highest scorer.

    Returns (output_path, report). `rules_path` is ignored and only kept
    to avoid breaking existing callers.
    """
    input_path = Path(input_path)
    template_path = Path(template_path)
    output_path = Path(output_path)

    # Auto-repair the template first so downstream picks / add_slide see
    # a well-formed file (missing <p:ph> idx/type, creationId GUIDs).
    try:
        from template_repair import repair_template
        repair_template(template_path)
    except Exception:
        # Repair failure is non-fatal — proceed with the template as-is.
        pass

    input_prs = Presentation(str(input_path))
    input_slides = list(input_prs.slides)

    # Build output presentation from a fresh copy of the template
    output_prs = Presentation(str(template_path))
    _delete_all_slides(output_prs)

    report: list[dict] = []
    output_slide_counter = 0

    for i, in_slide in enumerate(input_slides, start=1):
        sig = analyze_input_slide(in_slide, i)
        layout = _pick_layout_for_slide(output_prs, sig)
        if layout is None:
            # Template has no layouts at all — degenerate case.
            continue

        new_slide = output_prs.slides.add_slide(layout)
        output_slide_counter += 1
        _emit_report(
            report, output_slide_counter, i, sig, layout.name, "signature-match"
        )
        _fill_body_slide(new_slide, sig, layout.name)

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
    _write_paragraphs(tb, paragraphs, is_body=True, bullet=None)


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
            layout=slide.slide_layout,
        )

    # Peel off figure/table markers so they land in a dedicated red TextBox
    # instead of being lumped into the body placeholder.
    content_paras, marker_paras = _split_out_figure_markers(sig.body_paragraphs)

    # Bullet definition is inferred from the *target template* (its
    # layout ph lstStyle, else the master's bodyStyle, else a generic
    # PowerPoint default). Uploading a different template automatically
    # yields different bullets — no hardcoding required.
    prs = slide.part.package.presentation_part.presentation
    bullet = _pick_bullet_for_slide(prs, slide.slide_layout, level=0)

    bodies = _find_body_placeholders(slide)
    filled_body: set = set()
    if bodies and content_paras:
        # Primary body content goes into the first body placeholder. Any
        # secondary body placeholders (e.g. a right column) are dropped
        # unless the source deck has content for them — the template
        # decides how many columns exist; we don't over-fill.
        _write_paragraphs(bodies[0], content_paras, bullet=bullet)
        # Give the placeholder an explicit xfrm so PowerPoint's
        # normAutofit measures the exact box.
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
