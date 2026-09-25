"""Template-level layout preferences for the Master Swap engine.

Purpose
-------
The generic structural scorer in `master_swap.py` picks a template layout for
each source slide by placeholder-signature matching alone. That works well as
a default, but some templates encode stylistic choices the source can't
imply on its own — e.g. "when the deck has many bullets, split across two
columns rather than stacking in one", or "reserve the dark layout for
section dividers". Those are TEMPLATE decisions, not customer decisions.

This module lets a template ship a small JSON file that adds score bonuses
to specific target layouts when the source signature matches given
conditions. The engine remains generic — templates without a preferences
file fall through to the structural scorer unchanged.

Preferences file location
-------------------------
    <repo>/template_preferences/<template_stem>.json

The stem is the template filename without extension (e.g. `Template.pptx`
looks up `template_preferences/Template.json`). Missing file → no rules →
generic behavior. No customer name is ever consulted.

Schema
------
    {
      "description": "human-readable note (ignored by the engine)",
      "typography": {
        # Optional. When set, master-swap writes this size onto every body
        # run explicitly, overriding the template master's default body
        # size. Leave unset to inherit whatever the master defines.
        "body_font_size_pt": float
      },
      "rules": [
        {
          "name": "rule identifier for debugging",
          "when": {
            # Any subset of these; all present conditions must hold.
            "min_body_paragraphs": int,
            "max_body_paragraphs": int,
            "min_subheading_hits": int,
            "max_subheading_hits": int,
            "title_is_upper": bool,
            "has_body_placeholder": bool,
            "has_picture_placeholder": bool,
            # Matches when the source slide's own layout name is in this
            # list. Uses PowerPoint's built-in layout names (e.g.
            # "Section Header", "Title Slide") — structural signal from
            # the source deck, not customer-specific.
            "input_layout_name_in": ["Section Header", ...]
          },
          "prefer_layouts": ["layout name as it appears in the template"],
          "score_bonus": int  # points added to any layout in prefer_layouts
        }
      ]
    }

The bonus stacks with the structural score; it does not replace it. If no
rule matches, the layout gets zero bonus and the generic score wins.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

_PREFERENCES_DIR = Path(__file__).parent / "template_preferences"


def load_config(template_path: str | Path) -> dict | None:
    """Return the full preferences dict for a template, or None."""
    try:
        stem = Path(template_path).stem
    except Exception:
        return None
    prefs_file = _PREFERENCES_DIR / f"{stem}.json"
    if not prefs_file.exists():
        return None
    try:
        data = json.loads(prefs_file.read_text())
    except Exception:
        return None
    if not isinstance(data, dict):
        return None
    return data


def get_rules(config: dict | None) -> list[dict] | None:
    """Return the rule list from a config, or None."""
    if not isinstance(config, dict):
        return None
    rules = config.get("rules")
    return rules if isinstance(rules, list) else None


def get_body_font_size_pt(
    config: dict | None,
    layout_name: str | None = None,
) -> float | None:
    """Return the body font size override for the given layout, or None.

    Resolution order (first hit wins):
    1. `typography.body_font_size_pt_by_layout[<layout_name>]` — per-layout
       override for templates whose narrow multi-column layouts need a
       smaller body size than the wide single-column ones.
    2. `typography.body_font_size_pt` — flat default for all layouts.
    3. None — inherit the template master's default body size.
    """
    if not isinstance(config, dict):
        return None
    typo = config.get("typography")
    if not isinstance(typo, dict):
        return None
    by_layout = typo.get("body_font_size_pt_by_layout")
    if isinstance(by_layout, dict) and layout_name and layout_name in by_layout:
        try:
            return float(by_layout[layout_name])
        except Exception:
            pass
    size = typo.get("body_font_size_pt")
    try:
        return float(size) if size is not None else None
    except Exception:
        return None


def get_cover_chapter_divider(config: dict | None) -> dict | None:
    """Return the cover_chapter_divider settings, or None.

    Enables an extra slide to be synthesized after a source cover slide
    when the cover carries additional decorative text (e.g. a "Chapter N"
    label alongside a chapter title). Purely opt-in per template — no
    trigger unless the preferences file sets `enabled: true`.

    Schema:
        {"enabled": true, "layout": "Title - 2"}
    """
    if not isinstance(config, dict):
        return None
    ccd = config.get("cover_chapter_divider")
    if not isinstance(ccd, dict):
        return None
    if not ccd.get("enabled"):
        return None
    return ccd


def _rule_matches(rule: dict, sig: Any) -> bool:
    when = rule.get("when") or {}
    if not isinstance(when, dict):
        return False
    try:
        body_count = len(sig.body_paragraphs)
    except Exception:
        return False
    subheading_hits = getattr(sig, "subheading_hits", 0) or 0
    title_is_upper = bool(getattr(sig, "title_is_upper", False))
    has_body = bool(getattr(sig, "has_body_placeholder", False))
    has_picture = bool(getattr(sig, "has_picture_placeholder", False))

    if "min_body_paragraphs" in when and body_count < int(when["min_body_paragraphs"]):
        return False
    if "max_body_paragraphs" in when and body_count > int(when["max_body_paragraphs"]):
        return False
    if "min_subheading_hits" in when and subheading_hits < int(when["min_subheading_hits"]):
        return False
    if "max_subheading_hits" in when and subheading_hits > int(when["max_subheading_hits"]):
        return False
    if "title_is_upper" in when and bool(when["title_is_upper"]) != title_is_upper:
        return False
    if "has_body_placeholder" in when and bool(when["has_body_placeholder"]) != has_body:
        return False
    if "has_picture_placeholder" in when and bool(when["has_picture_placeholder"]) != has_picture:
        return False
    if "input_layout_name_in" in when:
        allowed = when["input_layout_name_in"]
        try:
            input_layout = getattr(sig, "input_layout_name", "") or ""
        except Exception:
            input_layout = ""
        if not isinstance(allowed, list) or input_layout not in allowed:
            return False
    return True


def preference_bonus(
    rules: list[dict] | None, sig: Any, layout_name: str
) -> int:
    """Sum bonuses from every matching rule that names this layout."""
    if not rules:
        return 0
    total = 0
    for rule in rules:
        try:
            if not _rule_matches(rule, sig):
                continue
            prefer = rule.get("prefer_layouts") or []
            if isinstance(prefer, list) and layout_name in prefer:
                total += int(rule.get("score_bonus", 0))
        except Exception:
            continue
    return total
