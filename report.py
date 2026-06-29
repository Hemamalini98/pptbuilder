"""
Generate a human-readable HTML change report comparing input vs output PPTX.
Shows actual rendered values (resolves inherited theme/master styles) so
"Theme default" never appears — only real before/after differences are listed.
Also lists figures that were requested in slide placeholders but skipped/missing,
or cropped but unused.
Usage: python3 report.py [input.pptx] [output.pptx] [report.html]
"""
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from lxml import etree
import json, sys, os, re

INPUT_PATH  = "input.pptx"
OUTPUT_PATH = "output.pptx"
REPORT_PATH = "change_report.html"

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS   = {"a": A_NS, "p": P_NS}

ALIGN_NAMES = {
    PP_ALIGN.LEFT: "Left", PP_ALIGN.CENTER: "Center",
    PP_ALIGN.RIGHT: "Right", PP_ALIGN.JUSTIFY: "Justify",
    PP_ALIGN.DISTRIBUTE: "Distribute",
}
SCHEME_MAP = {
    "tx1": "dk1", "bg1": "lt1", "tx2": "dk2", "bg2": "lt2",
    "accent1": "accent1", "accent2": "accent2", "accent3": "accent3",
    "accent4": "accent4", "accent5": "accent5", "accent6": "accent6",
    "hlink": "hlink", "folHlink": "folHlink",
}
BODY_LEVELS = ["lvl1pPr","lvl2pPr","lvl3pPr","lvl4pPr","lvl5pPr",
               "lvl6pPr","lvl7pPr","lvl8pPr","lvl9pPr"]
TITLE_TYPES = {"TITLE (1)", "CENTER_TITLE (3)"}
THEME_FONTS = {"+mn-lt","+mj-lt","+mn-ea","+mj-ea","+mn-cs","+mj-cs"}


# ── RESOLVE EFFECTIVE STYLES FROM SLIDE MASTER ──────────────────────────────

def get_theme_colors(prs):
    """Read the actual hex colors from the file's theme."""
    master_part = prs.slide_master.part
    for rel in master_part.rels.values():
        if "theme" in rel.reltype:
            try:
                theme_el = etree.fromstring(rel.target_part.blob)
                colors = {}
                clr = theme_el.find(".//a:clrScheme", NS)
                if clr is not None:
                    for child in clr:
                        name = child.tag.split("}")[-1]
                        srgb   = child.find(".//a:srgbClr", NS)
                        sysClr = child.find(".//a:sysClr",  NS)
                        if srgb is not None:
                            colors[name] = "#" + srgb.get("val", "").upper()
                        elif sysClr is not None:
                            v = sysClr.get("lastClr") or sysClr.get("val", "")
                            if v: colors[name] = "#" + v.upper()
                return colors
            except Exception:
                pass
    return {}


def resolve_color_el(el, colors):
    """Resolve a solidFill / srgbClr / schemeClr element to a hex string."""
    if el is None:
        return None
    srgb   = el.find(".//a:srgbClr", NS)
    scheme = el.find(".//a:schemeClr", NS)
    sysClr = el.find(".//a:sysClr",   NS)
    if srgb is not None:
        return "#" + srgb.get("val", "").upper()
    if scheme is not None:
        token = scheme.get("val", "")
        key   = SCHEME_MAP.get(token, token)
        return colors.get(key)
    if sysClr is not None:
        v = sysClr.get("lastClr") or sysClr.get("val", "")
        return ("#" + v.upper()) if v else None
    return None


def parse_level_style(lvl_el, colors):
    """Extract style dict from a lvlNpPr XML element."""
    s = {}
    algn = lvl_el.get("algn")
    if algn:
        s["alignment"] = {"l":"Left","ctr":"Center","r":"Right",
                          "just":"Justify","dist":"Distribute"}.get(algn, algn)
    defRPr = lvl_el.find("a:defRPr", NS)
    if defRPr is not None:
        sz = defRPr.get("sz")
        if sz:
            s["fontSize_pt"] = str(round(int(sz) / 100, 1)) + "pt"
        b = defRPr.get("b")
        if b is not None:
            s["bold"] = "Yes" if b == "1" else "No"
        i = defRPr.get("i")
        if i is not None:
            s["italic"] = "Yes" if i == "1" else "No"
        latin = defRPr.find("a:latin", NS)
        if latin is not None:
            f = latin.get("typeface", "")
            if f and f not in THEME_FONTS:
                s["fontFamily"] = f
        fill = defRPr.find("a:solidFill", NS)
        clr = resolve_color_el(fill, colors)
        if clr:
            s["color"] = clr
    return s


def get_master_styles(prs, colors):
    """Return {styleName: {levelKey: {prop: value}}} from the slide master."""
    txStyles = prs.slide_master.element.find("p:txStyles", NS)
    if txStyles is None:
        return {}
    out = {}
    for style_el in txStyles:
        name = style_el.tag.split("}")[-1]
        levels = {}
        for lvl_el in style_el:
            tag   = lvl_el.tag.split("}")[-1]
            style = parse_level_style(lvl_el, colors)
            if style:
                levels[tag] = style
        if levels:
            out[name] = levels
    return out


def effective_style(master_styles, ph_type_str, level):
    """Return the cascade-resolved style dict for a given placeholder type and bullet level."""
    style_name = "titleStyle" if any(t in ph_type_str for t in TITLE_TYPES) else "bodyStyle"
    level_key  = BODY_LEVELS[min(level, len(BODY_LEVELS) - 1)]
    lvl = master_styles.get(style_name, {}).get(level_key, {})
    # Fall back to level1 for inherited properties
    if not lvl:
        lvl = master_styles.get(style_name, {}).get("lvl1pPr", {})
    return lvl


# ── VALUE HELPERS ─────────────────────────────────────────────────────────────

def pt_str(size):
    try:
        return str(round(size.pt, 1)) + "pt" if size else None
    except Exception:
        return None


def rgb_str(font):
    try:
        if font.color and font.color.type is not None:
            return "#" + str(font.color.rgb).upper()
    except Exception:
        pass
    return None


def lnspc_str(para):
    v = para.line_spacing
    if v is None:
        return None
    if isinstance(v, float):
        return f"{round(v * 100)}%"
    try:
        return f"{v.pt}pt"
    except Exception:
        return str(v)


def align_str(para):
    return ALIGN_NAMES.get(para.alignment)


# ── DIFF ─────────────────────────────────────────────────────────────────────

def collect_changes(input_path, output_path):
    prs_in  = Presentation(input_path)
    prs_out = Presentation(output_path)

    # Resolve the input file's effective (inherited) styles
    colors_in      = get_theme_colors(prs_in)
    master_styles  = get_master_styles(prs_in, colors_in)

    slides_data = []

    for si, (sl_in, sl_out) in enumerate(zip(prs_in.slides, prs_out.slides)):
        shapes_in  = {s.shape_id: s for s in sl_in.shapes
                      if s.is_placeholder and s.has_text_frame}
        shapes_out = {s.shape_id: s for s in sl_out.shapes
                      if s.is_placeholder and s.has_text_frame}

        placeholders = []

        for sid in sorted(set(shapes_in) & set(shapes_out)):
            sh_in, sh_out = shapes_in[sid], shapes_out[sid]
            ph_type = str(sh_in.placeholder_format.type)

            paras_changed = []

            for pi, (p_in, p_out) in enumerate(
                    zip(sh_in.text_frame.paragraphs,
                        sh_out.text_frame.paragraphs)):

                level = p_in.level or 0
                eff   = effective_style(master_styles, ph_type, level)
                changes = []

                # ── PARAGRAPH LEVEL ──────────────────────────
                # Alignment
                a_in  = align_str(p_in)  or eff.get("alignment")
                a_out = align_str(p_out) or eff.get("alignment")
                if a_in != a_out:
                    changes.append({"prop":"Alignment",
                                    "before": a_in or "—", "after": a_out or "—"})

                # Line spacing
                ls_in  = lnspc_str(p_in)
                ls_out = lnspc_str(p_out)
                if ls_in != ls_out and ls_out:
                    changes.append({"prop":"Line spacing",
                                    "before": ls_in or "—", "after": ls_out})

                # Space before
                sb_in  = pt_str(p_in.space_before)
                sb_out = pt_str(p_out.space_before)
                if sb_in != sb_out and sb_out:
                    changes.append({"prop":"Space before",
                                    "before": sb_in or "—", "after": sb_out})

                # ── RUN LEVEL ────────────────────────────────
                for ri, (r_in, r_out) in enumerate(zip(p_in.runs, p_out.runs)):
                    fi, fo = r_in.font, r_out.font

                    # Font size
                    fs_in  = pt_str(fi.size) or eff.get("fontSize_pt")
                    fs_out = pt_str(fo.size)
                    if fs_in != fs_out and fs_out:
                        changes.append({"prop":"Font size",
                                        "before": fs_in or "—", "after": fs_out})

                    # Font family
                    ff_in  = fi.name if fi.name and fi.name not in THEME_FONTS else eff.get("fontFamily")
                    ff_out = fo.name if fo.name and fo.name not in THEME_FONTS else None
                    if ff_in != ff_out and ff_out:
                        changes.append({"prop":"Font family",
                                        "before": ff_in or "—", "after": ff_out})

                    # Bold
                    b_in  = ("Yes" if fi.bold else "No") if fi.bold is not None else eff.get("bold")
                    b_out = ("Yes" if fo.bold else "No") if fo.bold is not None else None
                    if b_in != b_out and b_out:
                        changes.append({"prop":"Bold",
                                        "before": b_in or "—", "after": b_out})

                    # Color
                    c_in  = rgb_str(fi) or eff.get("color")
                    c_out = rgb_str(fo)
                    if c_in != c_out and c_out:
                        changes.append({"prop":"Color", "is_color": True,
                                        "before": c_in or "—", "after": c_out})

                if changes:
                    paras_changed.append({
                        "text":    p_in.text[:55] if p_in.text else "",
                        "idx":     pi,
                        "changes": changes,
                    })

            if paras_changed:
                placeholders.append({
                    "name":  sh_in.name,
                    "type":  ph_type,
                    "idx":   sh_in.placeholder_format.idx,
                    "paras": paras_changed,
                })

        if placeholders:
            slides_data.append({"slide": si + 1, "placeholders": placeholders})

    return slides_data


def collect_figure_diagnostics(input_path, extracts_dir):
    """Scan input PPT for requested figures, compare against PDF crops."""
    if not os.path.exists(input_path):
        return [], []

    requested_figs = set()
    try:
        prs = Presentation(input_path)
        _FIG_PAT = re.compile(r'(figure|table)\s+([\d.]+)', re.IGNORECASE)
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for m in _FIG_PAT.finditer(shape.text_frame.text):
                        requested_figs.add(f"{m.group(1).lower()} {m.group(2)}")
    except Exception as e:
        print("Error checking requested figures:", e)

    cropped_figs = set()
    if os.path.isdir(extracts_dir):
        for fname in os.listdir(extracts_dir):
            if fname.lower().endswith(('.png', '.jpg', '.jpeg')):
                key = os.path.splitext(fname)[0].strip().lower()
                cropped_figs.add(key)

    missing_figs = sorted(list(requested_figs - cropped_figs))
    unplaced_figs = sorted(list(cropped_figs - requested_figs))
    return missing_figs, unplaced_figs


# ── HTML TEMPLATE ─────────────────────────────────────────────────────────────

HTML = '''<!DOCTYPE html>
<html>
<head>
<title>PPT Style & Figure Placement Report</title>
<style>
:root{--bg:#F8FAFC;--surface:#FFFFFF;--border:#E2E8F0;--ink:#0F172A;--muted:#64748B;--accent:#EA580C;--green:#16A34A;--tag-font:#1D4ED8;--tag-font-bg:#DBEAFE;--tag-size:#B45309;--tag-size-bg:#FEF3C7;--tag-color:#6D28D9;--tag-color-bg:#EDE9FE;--tag-bold:#374151;--tag-bold-bg:#F3F4F6;--tag-align:#0E7490;--tag-align-bg:#CFFAFE;--tag-spacing:#475569;--tag-spacing-bg:#F1F5F9}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{font-family:system-ui,-apple-system,"Segoe UI",sans-serif;background:var(--bg);color:var(--ink);font-size:13.5px;line-height:1.55}
.hd{background:var(--ink);color:#fff;padding:2.25rem 2.5rem 2rem}
.hd-top{display:flex;align-items:flex-start;justify-content:space-between;flex-wrap:wrap;gap:1rem}
.hd h1{font-size:1.6rem;font-weight:700;letter-spacing:-.01em}
.hd-source{font-size:.75rem;color:rgba(255,255,255,.45);margin-top:.3rem}
.hd-badge{background:var(--accent);color:#fff;font-size:.65rem;font-weight:700;text-transform:uppercase;letter-spacing:.1em;padding:.25rem .6rem;border-radius:3px;align-self:flex-start;white-space:nowrap}
.stats{display:flex;gap:2.5rem;margin-top:1.75rem;padding-top:1.5rem;border-top:1px solid rgba(255,255,255,.1)}
.stat-n{font-size:2rem;font-weight:700;font-variant-numeric:tabular-nums;line-height:1;letter-spacing:-.02em}
.stat-l{font-size:.68rem;color:rgba(255,255,255,.45);text-transform:uppercase;letter-spacing:.09em;margin-top:.3rem}

.diag-grid{display:grid;grid-template-columns:1fr 1fr;gap:1.5rem;margin-top:1.5rem}
.diag-card{background:var(--surface);border:1px solid var(--border);border-radius:7px;padding:1.25rem;overflow:hidden}
.diag-card-title{font-size:.85rem;font-weight:700;text-transform:uppercase;letter-spacing:.05em;margin-bottom:.75rem;display:flex;align-items:center;gap:.5rem}
.diag-card-title.missing{color:#E11D48}
.diag-card-title.unplaced{color:#475569}
.diag-list{list-style:none;padding:0;margin:0;font-size:.8rem;display:flex;flex-wrap:wrap;gap:.4rem}
.diag-item{background:#F1F5F9;color:#334155;padding:.2rem .5rem;border-radius:4px;font-family:monospace;font-weight:bold}
.diag-item.missing{background:#FFE4E6;color:#9F1239}
.diag-empty{color:var(--muted);font-style:italic;font-size:.8rem}

.legend{background:var(--surface);border-bottom:1px solid var(--border);padding:.7rem 2.5rem;display:flex;align-items:center;gap:.5rem;flex-wrap:wrap}
.legend-lbl{font-size:.65rem;text-transform:uppercase;letter-spacing:.1em;color:var(--muted);margin-right:.25rem}
.content{max-width:960px;margin:0 auto;padding:2rem 2.5rem 4rem;display:flex;flex-direction:column;gap:1.25rem}
.no-changes{text-align:center;padding:3rem;color:var(--muted);font-size:.9rem}
.card{background:var(--surface);border:1px solid var(--border);border-radius:7px;overflow:hidden}
.card-hd{background:var(--ink);color:#fff;padding:.65rem 1.25rem;display:flex;align-items:center;gap:.75rem}
.slide-pill{font-size:.62rem;font-weight:700;text-transform:uppercase;letter-spacing:.1em;border:1px solid rgba(255,255,255,.25);border-radius:3px;padding:.15rem .5rem;font-variant-numeric:tabular-nums;white-space:nowrap}
.card-hd-txt{font-size:.82rem;color:rgba(255,255,255,.65)}
.card-hd-count{margin-left:auto;font-size:.72rem;color:rgba(255,255,255,.4);font-variant-numeric:tabular-nums;white-space:nowrap}
.ph-section{border-top:1px solid var(--border)}
.ph-section:first-of-type{border-top:none}
.ph-hd{padding:.5rem 1.25rem;background:#F8FAFC;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:.5rem}
.ph-badge{font-size:.6rem;font-weight:700;text-transform:uppercase;letter-spacing:.1em;color:var(--accent);border:1px solid var(--accent);padding:.1rem .4rem;border-radius:3px}
.ph-name{font-size:.75rem;color:var(--muted)}
.para-block{padding:.75rem 1.25rem .85rem;border-bottom:1px solid #F1F5F9}
.para-block:last-child{border-bottom:none}
.para-quote{font-size:.78rem;color:var(--muted);font-style:italic;margin-bottom:.55rem;display:flex;align-items:baseline;gap:.35rem}
.para-quote::before{content:'"';font-family:Georgia,serif;font-size:1.3rem;color:var(--border);line-height:.8;flex-shrink:0}
.changes{display:flex;flex-direction:column;gap:.3rem}
.chg{display:grid;grid-template-columns:80px 1fr 14px 1fr;align-items:center;gap:.5rem;font-size:.8rem}
.tag{display:inline-flex;align-items:center;justify-content:center;font-size:.6rem;font-weight:700;text-transform:uppercase;letter-spacing:.08em;padding:.15rem .45rem;border-radius:3px;width:80px;cursor:pointer;user-select:none;transition:transform 0.15s, box-shadow 0.15s}
.tag:hover{transform:scale(1.05)}
.tag.active{outline:2px solid var(--ink);outline-offset:1px;box-shadow:0 0 8px rgba(0,0,0,0.15);font-weight:800}
.tag.inactive{opacity:0.35;transform:scale(0.95)}
.t-font{background:var(--tag-font-bg);color:var(--tag-font)}.t-size{background:var(--tag-size-bg);color:var(--tag-size)}.t-color{background:var(--tag-color-bg);color:var(--tag-color)}.t-bold{background:var(--tag-bold-bg);color:var(--tag-bold)}.t-align{background:var(--tag-align-bg);color:var(--tag-align)}.t-spacing{background:var(--tag-spacing-bg);color:var(--tag-spacing)}
.chg-before{color:var(--muted);text-decoration:line-through;text-decoration-color:rgba(0,0,0,.2);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.chg-arrow{color:#CBD5E1;font-size:.7rem;text-align:center}
.chg-after{color:var(--green);font-weight:600;display:flex;align-items:center;gap:.35rem;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.swatch{display:inline-block;width:11px;height:11px;border-radius:2px;border:1px solid rgba(0,0,0,.12);flex-shrink:0}
.footer{text-align:center;padding:1rem 0 2rem;font-size:.7rem;color:var(--muted)}
@media(max-width:600px){.hd,.legend,.content{padding-left:1rem;padding-right:1rem}.stats{gap:1.5rem}.chg{grid-template-columns:70px 1fr 12px 1fr}.diag-grid{grid-template-columns:1fr}}
</style>
</head>
<body>
<div class="hd">
  <div class="hd-top">
    <div><h1>Style & Figure Placement Report</h1><div class="hd-source">__INPUT__ &nbsp;→&nbsp; __OUTPUT__</div></div>
    <span class="hd-badge">Template Applied</span>
  </div>
  <div class="stats" id="stats"></div>
</div>
<div class="legend" id="legend">
  <span class="legend-lbl">Filter by change key:</span>
  <span class="tag t-font" data-filter="Font">Font</span>
  <span class="tag t-size" data-filter="Size">Size</span>
  <span class="tag t-color" data-filter="Color">Color</span>
  <span class="tag t-bold" data-filter="Bold">Bold</span>
  <span class="tag t-align" data-filter="Alignment">Alignment</span>
  <span class="tag t-spacing" data-filter="Spacing">Spacing</span>
</div>
<div class="content" id="content">
  <div id="cards-container"></div>
</div>
<div class="footer">Generated by pptxBuilder · report.py</div>
<script>
const DATA=__DATA__;
const TAG={'Font family':['t-font','Font'],'Font size':['t-size','Size'],'Color':['t-color','Color'],'Bold':['t-bold','Bold'],'Alignment':['t-align','Align'],'Line spacing':['t-spacing','Line Spc'],'Space before':['t-spacing','Spc Before'],'Space after':['t-spacing','Spc After']};
function phLabel(t){if(/center.title/i.test(t))return'Center Title';if(/subtitle/i.test(t))return'Subtitle';if(/title/i.test(t))return'Title';if(/body|object|text/i.test(t))return'Body/Content';return t;}
let tc=0,ps=new Set();DATA.forEach(s=>s.placeholders.forEach(p=>p.paras.forEach(q=>{tc+=q.changes.length;q.changes.forEach(c=>ps.add(c.prop));})));
const ct=document.getElementById('cards-container');
document.getElementById('stats').innerHTML=`<div><div class="stat-n">${DATA.length}</div><div class="stat-l">Slides Changed</div></div><div><div class="stat-n">${tc}</div><div class="stat-l">Real Differences</div></div><div><div class="stat-n">${ps.size}</div><div class="stat-l">Properties</div></div>`;

let activeFilter = null;
const FILTER_MAP = {
  'Font': ['Font family'],
  'Size': ['Font size'],
  'Color': ['Color'],
  'Bold': ['Bold'],
  'Alignment': ['Alignment'],
  'Spacing': ['Line spacing', 'Space before', 'Space after']
};

function render() {
  ct.innerHTML = '';
  
  document.querySelectorAll('#legend .tag').forEach(tag => {
    const f = tag.getAttribute('data-filter');
    tag.classList.remove('active', 'inactive');
    if (activeFilter) {
      if (f === activeFilter) {
        tag.classList.add('active');
      } else {
        tag.classList.add('inactive');
      }
    }
  });

  // Filter slides
  const filtered = DATA.map(slide => {
    if (!activeFilter) return slide;
    const allowedProps = FILTER_MAP[activeFilter];
    
    const newPlaceholders = slide.placeholders.map(ph => {
      const newParas = ph.paras.map(para => {
        const newChanges = para.changes.filter(c => allowedProps.includes(c.prop));
        if (newChanges.length > 0) {
          return { ...para, changes: newChanges };
        }
        return null;
      }).filter(Boolean);

      if (newParas.length > 0) {
        return { ...ph, paras: newParas };
      }
      return null;
    }).filter(Boolean);

    if (newPlaceholders.length > 0) {
      return { ...slide, placeholders: newPlaceholders };
    }
    return null;
  }).filter(Boolean);

  if (filtered.length === 0) {
    ct.innerHTML = '<div class="no-changes">No visual style differences match the selected filter.</div>';
    return;
  }

  filtered.forEach(slide => {
    const sc = slide.placeholders.reduce((a, p) => a + p.paras.reduce((b, q) => b + q.changes.length, 0), 0);
    const card = document.createElement('div');
    card.className = 'card';
    card.style.marginBottom = '1.25rem';
    card.innerHTML = `<div class="card-hd"><span class="slide-pill">Slide ${slide.slide}</span><span class="card-hd-txt">Formatting updated</span><span class="card-hd-count">${sc} change${sc !== 1 ? 's' : ''}</span></div>`;
    
    slide.placeholders.forEach(ph => {
      const sec = document.createElement('div');
      sec.className = 'ph-section';
      sec.innerHTML = `<div class="ph-hd"><span class="ph-badge">${phLabel(ph.type)}</span><span class="ph-name">${ph.name}</span></div>`;
      
      ph.paras.forEach(para => {
        const blk = document.createElement('div');
        blk.className = 'para-block';
        const dt = para.text ? `${para.text}${para.text.length >= 55 ? '…' : ''}` : '(empty paragraph)';
        const ch = para.changes.map(c => {
          const [cls, lbl] = TAG[c.prop] || ['t-spacing', c.prop];
          let ae = c.is_color && c.after && c.after.startsWith('#') ? `<span class="swatch" style="background:${c.after}"></span>${c.after}` : c.after;
          let be = c.is_color && c.before && c.before.startsWith('#') ? `<span class="swatch" style="background:${c.before}"></span>${c.before}` : c.before;
          return `<div class="chg"><span class="tag ${cls}">${lbl}</span><span class="chg-before">${be}</span><span class="chg-arrow">→</span><span class="chg-after">${ae}</span></div>`;
        }).join('');
        blk.innerHTML = `<div class="para-quote">${dt}</div><div class="changes">${ch}</div>`;
        sec.appendChild(blk);
      });
      card.appendChild(sec);
    });
    ct.appendChild(card);
  });
}

document.querySelectorAll('#legend .tag').forEach(tag => {
  tag.addEventListener('click', () => {
    const f = tag.getAttribute('data-filter');
    if (activeFilter === f) {
      activeFilter = null;
    } else {
      activeFilter = f;
    }
    render();
  });
});

render();
</script>
</body>
</html>'''


def generate_report(input_path, output_path, report_path):
    print(f"Comparing {input_path} → {output_path} ...")
    data = collect_changes(input_path, output_path)
    total = sum(len(c['changes']) for s in data for p in s['placeholders'] for c in p['paras'])
    print(f"  {len(data)} slides with real changes, {total} actual visual differences")

    input_dir = os.path.dirname(os.path.abspath(input_path))
    extracts_dir = os.path.join(input_dir, "pdf_extracts")
    missing, unplaced = collect_figure_diagnostics(input_path, extracts_dir)

    # Build missing html
    if missing:
        missing_html = '<ul class="diag-list">' + ''.join(f'<li class="diag-item missing">{m.upper()}</li>' for m in missing) + '</ul>'
    else:
        missing_html = '<div class="diag-empty">No skipped/missing figures. All requests placed.</div>'

    # Build unplaced html
    if unplaced:
        unplaced_html = '<ul class="diag-list">' + ''.join(f'<li class="diag-item">{u.upper()}</li>' for u in unplaced) + '</ul>'
    else:
        unplaced_html = '<div class="diag-empty">No unused cropped figures.</div>'

    html = (HTML
        .replace('__INPUT__',  os.path.basename(input_path))
        .replace('__OUTPUT__', os.path.basename(output_path))
        .replace('__DATA__',   json.dumps(data))
    )
    with open(report_path, "w") as f:
        f.write(html)
    print(f"  Report saved: {report_path}")


if __name__ == "__main__":
    ip = sys.argv[1] if len(sys.argv) > 1 else INPUT_PATH
    op = sys.argv[2] if len(sys.argv) > 2 else OUTPUT_PATH
    rp = sys.argv[3] if len(sys.argv) > 3 else REPORT_PATH
    generate_report(ip, op, rp)
