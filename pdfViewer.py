#!/usr/bin/env python3
"""
pdfViewer.py — Interactive PDF viewer with drag-to-extract (PDF.js rendering)
Usage:  python3 pdfViewer.py <file.pdf>
"""
import sys, os, json, socket, threading, webbrowser, tempfile
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse

try:
    import fitz
    fitz.TOOLS.mupdf_display_errors(False)
    fitz.TOOLS.mupdf_display_warnings(False)
except ImportError:
    sys.exit("PyMuPDF not found.  Run: pip3 install pymupdf --break-system-packages")

EXTRACT_SCALE = 3.0    # PyMuPDF render scale for saved extractions (high quality)

PDF_PATH    = ""
doc         = None
EXTRACT_DIR = ""


def crop_png(page_num: int, clip: "fitz.Rect") -> bytes:
    page = doc[page_num]
    pix  = page.get_pixmap(
        matrix=fitz.Matrix(EXTRACT_SCALE, EXTRACT_SCALE),
        clip=clip, alpha=False,
    )
    return pix.tobytes("png")


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a): pass

    def reply(self, code, ctype, body, extra=None):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        if extra:
            for k, v in extra.items(): self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path = urlparse(self.path).path.rstrip("/") or "/"

        if path == "/":
            self.reply(200, "text/html; charset=utf-8", HTML.encode())

        elif path == "/api/info":
            self.reply(200, "application/json", json.dumps({
                "filename": os.path.basename(PDF_PATH),
                "count":    len(doc),
            }).encode())

        elif path == "/api/pdf":
            with open(PDF_PATH, "rb") as f:
                data = f.read()
            self.reply(200, "application/pdf", data,
                       {"Content-Disposition": "inline"})

        elif path.startswith("/api/img/"):
            fname = path.rsplit("/", 1)[-1]
            fpath = os.path.join(EXTRACT_DIR, fname)
            if os.path.isfile(fpath):
                self.reply(200, "image/png", open(fpath, "rb").read())
            else:
                self.send_error(404)

        else:
            self.send_error(404)

    def do_POST(self):
        if urlparse(self.path).path != "/api/extract":
            return self.send_error(404)
        length = int(self.headers.get("Content-Length", 0))
        body   = json.loads(self.rfile.read(length))
        n      = body["page"]
        scale  = body["scale"]   # PDF.js renderScale: CSS px per PDF pt
        clip   = fitz.Rect(
            min(body["x0"], body["x1"]) / scale,
            min(body["y0"], body["y1"]) / scale,
            max(body["x0"], body["x1"]) / scale,
            max(body["y0"], body["y1"]) / scale,
        )
        png    = crop_png(n, clip)
        existing = [f for f in os.listdir(EXTRACT_DIR) if f.endswith(".png")]
        idx    = len(existing) + 1
        fname  = f"p{n+1}_extract{idx:03d}.png"
        with open(os.path.join(EXTRACT_DIR, fname), "wb") as f:
            f.write(png)
        self.reply(200, "application/json", json.dumps({
            "filename": fname, "url": f"/api/img/{fname}", "page": n + 1,
        }).encode())

    def do_DELETE(self):
        path = urlparse(self.path).path
        if path.startswith("/api/img/"):
            fname = path.rsplit("/", 1)[-1]
            fpath = os.path.join(EXTRACT_DIR, fname)
            if os.path.isfile(fpath): os.remove(fpath)
            self.reply(200, "application/json", b'{"ok":true}')
        else:
            self.send_error(404)


HTML = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>PDF Viewer</title>
<script src="https://unpkg.com/pdfjs-dist@3.11.174/build/pdf.min.js"></script>
<style>
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0 }
body {
  font-family: system-ui, -apple-system, sans-serif;
  background: #0f172a; color: #e2e8f0;
  display: flex; height: 100vh; overflow: hidden; font-size: 13px
}

/* ── Sidebar ───────────────────────────────────────── */
.sidebar {
  width: 176px; flex-shrink: 0; background: #1e293b;
  display: flex; flex-direction: column; border-right: 1px solid #334155
}
.sidebar-hdr { padding: 12px 10px 10px; border-bottom: 1px solid #334155 }
.sidebar-hdr .fname {
  display: block; font-weight: 600; font-size: 12px; color: #f1f5f9;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis
}
.sidebar-hdr .pcount { color: #64748b; font-size: 11px }
.thumbs {
  flex: 1; overflow-y: auto; padding: 8px;
  display: flex; flex-direction: column; gap: 6px
}
.thumbs::-webkit-scrollbar { width: 4px }
.thumbs::-webkit-scrollbar-thumb { background: #334155; border-radius: 2px }
.thumb {
  position: relative; border-radius: 3px; overflow: hidden;
  cursor: pointer; border: 2px solid transparent; transition: border-color .15s;
  background: #0f172a; flex-shrink: 0
}
.thumb:hover { border-color: #38bdf860 }
.thumb.active { border-color: #38bdf8 }
.thumb canvas { width: 100%; display: block }
.thumb .pg {
  position: absolute; bottom: 0; left: 0; right: 0;
  background: rgba(0,0,0,.65); text-align: center;
  padding: 2px 0; font-size: 10px; color: #94a3b8
}

/* ── Main ──────────────────────────────────────────── */
.main { flex: 1; display: flex; flex-direction: column; min-width: 0; overflow: hidden }

.toolbar {
  background: #1e293b; border-bottom: 1px solid #334155;
  padding: 7px 12px; display: flex; align-items: center; gap: 8px; flex-shrink: 0
}
.toolbar button {
  border: none; border-radius: 5px; padding: 5px 13px;
  font-size: 12px; font-weight: 600; cursor: pointer; transition: background .15s
}
#btn-extract { background: #0284c7; color: #fff }
#btn-extract:hover:not(:disabled) { background: #0369a1 }
#btn-extract:disabled { background: #1e3a4f; color: #475569; cursor: default }
#btn-clear { background: #334155; color: #94a3b8 }
#btn-clear:hover:not(:disabled) { background: #475569; color: #e2e8f0 }
#btn-clear:disabled { opacity: .4; cursor: default }
.coords { font-size: 11px; color: #64748b; flex: 1; font-variant-numeric: tabular-nums }
.page-info { font-size: 11px; color: #475569; white-space: nowrap }

/* ── Zoom controls ─────────────────────────────────── */
.sep { width: 1px; background: #334155; height: 20px; margin: 0 2px; flex-shrink: 0 }
.zoom-grp { display: flex; align-items: center; gap: 3px; flex-shrink: 0 }
.zoom-grp .zbtn {
  background: #334155; border: none; color: #e2e8f0;
  width: 24px; height: 24px; border-radius: 4px; cursor: pointer;
  font-size: 15px; font-weight: 700; padding: 0; line-height: 1;
  display: flex; align-items: center; justify-content: center
}
.zoom-grp .zbtn:hover { background: #475569 }
#zoom-sel {
  background: #334155; border: 1px solid #475569; color: #e2e8f0;
  padding: 2px 2px 2px 6px; border-radius: 4px; font-size: 11px;
  cursor: pointer; height: 24px; min-width: 70px
}

/* ── Page scroll area ──────────────────────────────── */
.page-area {
  flex: 1; overflow: auto; background: #374151;
  padding: 20px 24px; min-height: 0
}
.page-area::-webkit-scrollbar { width: 6px; height: 6px }
.page-area::-webkit-scrollbar-thumb { background: #475569; border-radius: 3px }
.pages-inner {
  display: flex; flex-direction: column; align-items: center;
  gap: 16px; min-width: fit-content
}
.page-wrap {
  position: relative; flex-shrink: 0;
  box-shadow: 0 6px 24px rgba(0,0,0,.55);
  cursor: crosshair; user-select: none; background: white; overflow: hidden
}
.page-wrap .pdf-canvas { display: block; pointer-events: none }
.page-wrap .sel-canvas { position: absolute; inset: 0; pointer-events: none }
.page-wrap.loading { background: #1e293b }
.page-wrap.loading::after {
  content: ''; position: absolute; inset: 0;
  background: linear-gradient(90deg, transparent 0%, #334155 50%, transparent 100%);
  background-size: 200% 100%; animation: shimmer 1.4s infinite
}
@keyframes shimmer { 0%{background-position:200% 0} 100%{background-position:-200% 0} }

/* ── Extractions bar ───────────────────────────────── */
.ext-bar {
  background: #1e293b; border-top: 1px solid #334155;
  padding: 8px 12px; flex-shrink: 0; max-height: 200px; overflow: hidden
}
.ext-bar-hdr { display: flex; align-items: center; gap: 8px; margin-bottom: 6px }
.ext-bar-hdr h4 {
  font-size: 10px; font-weight: 600; letter-spacing: .08em;
  color: #475569; text-transform: uppercase; flex: 1
}
.ext-bar-hdr .btn-clear-all {
  border: none; background: #334155; color: #94a3b8;
  font-size: 10px; padding: 2px 8px; border-radius: 3px; cursor: pointer
}
.ext-bar-hdr .btn-clear-all:hover { background: #ef4444; color: #fff }
.ext-list { display: flex; gap: 8px; overflow-x: auto; padding-bottom: 4px }
.ext-list::-webkit-scrollbar { height: 4px }
.ext-list::-webkit-scrollbar-thumb { background: #334155; border-radius: 2px }
.ext-card {
  flex-shrink: 0; border-radius: 5px; border: 1px solid #334155;
  background: #0f172a; display: flex; flex-direction: column;
  overflow: hidden; width: 160px
}
.ext-card-img { height: 130px; background: #0f172a;
  display: flex; align-items: center; justify-content: center }
.ext-card-img img { max-width: 100%; max-height: 130px; display: block }
.ext-card-foot {
  display: flex; align-items: center; padding: 3px 5px;
  background: #1e293b; gap: 4px; flex-shrink: 0
}
.ext-card-foot .lbl { font-size: 10px; color: #64748b; flex: 1;
  overflow: hidden; text-overflow: ellipsis; white-space: nowrap }
.ext-card-foot button { border: none; background: none; cursor: pointer;
  font-size: 13px; padding: 0 2px; line-height: 1 }
.ext-card-foot .btn-dl  { color: #38bdf8 }
.ext-card-foot .btn-del { color: #64748b }
.ext-card-foot .btn-dl:hover  { color: #7dd3fc }
.ext-card-foot .btn-del:hover { color: #ef4444 }
</style>
</head>
<body>

<aside class="sidebar">
  <div class="sidebar-hdr">
    <span class="fname" id="fname">Loading…</span>
    <span class="pcount" id="pcount"></span>
  </div>
  <div class="thumbs" id="thumbs"></div>
</aside>

<div class="main">
  <div class="toolbar">
    <button id="btn-extract" disabled>⬇ Extract</button>
    <button id="btn-clear"   disabled>✕ Clear</button>
    <div class="sep"></div>
    <div class="zoom-grp">
      <button class="zbtn" id="btn-zoom-out" title="Zoom out (Ctrl+−)">−</button>
      <select id="zoom-sel">
        <option value="0.25">25%</option>
        <option value="0.5">50%</option>
        <option value="0.75">75%</option>
        <option value="1">100%</option>
        <option value="1.25">125%</option>
        <option value="1.5">150%</option>
        <option value="2">200%</option>
        <option value="3">300%</option>
        <option value="fit">Fit width</option>
      </select>
      <button class="zbtn" id="btn-zoom-in" title="Zoom in (Ctrl+=)">+</button>
    </div>
    <div class="sep"></div>
    <span class="coords"    id="coords">Drag on any page to select a region</span>
    <span class="page-info" id="page-info"></span>
  </div>

  <div class="page-area" id="page-area">
    <div class="pages-inner" id="pages-inner"></div>
  </div>

  <div class="ext-bar" id="ext-bar" hidden>
    <div class="ext-bar-hdr">
      <h4>Extracted images</h4>
      <button class="btn-clear-all" id="btn-clear-all">Delete all</button>
    </div>
    <div class="ext-list" id="ext-list"></div>
  </div>
</div>

<script>
'use strict';

pdfjsLib.GlobalWorkerOptions.workerSrc =
  'https://unpkg.com/pdfjs-dist@3.11.174/build/pdf.worker.min.js';

// ── State ─────────────────────────────────────────────────────
let pdfDoc      = null;
let pageCount   = 0, extCount = 0, curPage = 0;
let renderScale = 1.0;   // CSS px per PDF pt
let pdfPtW      = 612, pdfPtH = 792;
const ZOOM_STOPS = [0.25, 0.5, 0.75, 1.0, 1.25, 1.5, 2.0, 3.0];
const rendered   = new Map();   // page_num → scale rendered at
const renderTasks = new Map();  // page_num → active render task

// Selection
let sel = null, dragging = false, dragStart = {x:0, y:0};
let dragPage = -1, activeSelCanvas = null, activeSelCtx = null;

// DOM
const pagesEl    = document.getElementById('page-area');
const pagesInner = document.getElementById('pages-inner');
const thumbsEl   = document.getElementById('thumbs');
const coordsEl   = document.getElementById('coords');
const pageInfo   = document.getElementById('page-info');
const btnExt     = document.getElementById('btn-extract');
const btnClr     = document.getElementById('btn-clear');
const extBar     = document.getElementById('ext-bar');
const extList    = document.getElementById('ext-list');
const zoomSel    = document.getElementById('zoom-sel');

// ── Boot ─────────────────────────────────────────────────────
(async () => {
  const info = await fetch('/api/info').then(r => r.json());
  pageCount = info.count;
  document.getElementById('fname').textContent  = info.filename;
  document.getElementById('pcount').textContent = info.count + ' pages';
  document.title = info.filename + ' — PDF Viewer';

  pdfDoc = await pdfjsLib.getDocument('/api/pdf').promise;

  const pg1 = await pdfDoc.getPage(1);
  const vp1 = pg1.getViewport({ scale: 1.0 });
  pdfPtW = vp1.width;
  pdfPtH = vp1.height;

  buildPages();
  buildThumbs();
  updatePageInfo(0);
  requestAnimationFrame(fitZoom);
})();

// ── Build page wrappers ────────────────────────────────────────
function buildPages() {
  pagesInner.innerHTML = '';
  for (let i = 0; i < pageCount; i++) {
    const wrap = document.createElement('div');
    wrap.className = 'page-wrap loading';
    wrap.dataset.page = i;
    // placeholder size before render
    wrap.style.width  = Math.round(pdfPtW * renderScale) + 'px';
    wrap.style.height = Math.round(pdfPtH * renderScale) + 'px';

    const pdfCanvas = document.createElement('canvas');
    pdfCanvas.className = 'pdf-canvas';

    const selCanvas = document.createElement('canvas');
    selCanvas.className = 'sel-canvas';

    wrap.append(pdfCanvas, selCanvas);
    pagesInner.appendChild(wrap);
    loadObserver.observe(wrap);
    pageObserver.observe(wrap);
  }
}

// ── Render a page via PDF.js ───────────────────────────────────
async function renderPage(wrap) {
  const n = +wrap.dataset.page;
  if (renderTasks.has(n)) {
    try { renderTasks.get(n).cancel(); } catch (_) {}
  }

  const page     = await pdfDoc.getPage(n + 1);
  const viewport = page.getViewport({ scale: renderScale });
  const dpr      = window.devicePixelRatio || 1;
  const cssW     = Math.round(viewport.width);
  const cssH     = Math.round(viewport.height);

  wrap.style.width  = cssW + 'px';
  wrap.style.height = cssH + 'px';

  const pdfCanvas = wrap.querySelector('.pdf-canvas');
  pdfCanvas.width  = Math.round(cssW * dpr);
  pdfCanvas.height = Math.round(cssH * dpr);
  pdfCanvas.style.width  = cssW + 'px';
  pdfCanvas.style.height = cssH + 'px';

  const selCanvas = wrap.querySelector('.sel-canvas');
  selCanvas.width  = cssW;
  selCanvas.height = cssH;
  selCanvas.style.width  = cssW + 'px';
  selCanvas.style.height = cssH + 'px';
  if (sel && sel.page === n) redrawSel();

  const ctx = pdfCanvas.getContext('2d');
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);

  const task = page.render({ canvasContext: ctx, viewport });
  renderTasks.set(n, task);
  try {
    await task.promise;
    rendered.set(n, renderScale);
    wrap.classList.remove('loading');
  } catch (e) {
    if (e.name !== 'RenderingCancelledException') console.error(e);
  }
  renderTasks.delete(n);
}

// ── Lazy load: render when page enters viewport ───────────────
const loadObserver = new IntersectionObserver(entries => {
  entries.forEach(entry => {
    if (!entry.isIntersecting) return;
    const n = +entry.target.dataset.page;
    if (!rendered.has(n)) renderPage(entry.target);
  });
}, { rootMargin: '600px' });

// ── Track current page ────────────────────────────────────────
const visible = new Set();
const pageObserver = new IntersectionObserver(entries => {
  entries.forEach(entry => {
    const n = +entry.target.dataset.page;
    if (entry.isIntersecting) visible.add(n); else visible.delete(n);
  });
  if (visible.size) updatePageInfo(Math.min(...visible));
}, { threshold: 0.01, root: pagesEl });

function updatePageInfo(n) {
  curPage = n;
  pageInfo.textContent = `Page ${n + 1} / ${pageCount}`;
  document.querySelectorAll('.thumb').forEach(c =>
    c.classList.toggle('active', +c.dataset.page === n)
  );
  thumbsEl.querySelector('.thumb.active')?.scrollIntoView({ block: 'nearest' });
}

// ── Thumbnails via PDF.js ─────────────────────────────────────
const THUMB_SCALE = 0.25;

function buildThumbs() {
  thumbsEl.innerHTML = '';
  const tw = Math.round(pdfPtW * THUMB_SCALE);
  const th = Math.round(pdfPtH * THUMB_SCALE);
  for (let i = 0; i < pageCount; i++) {
    const card = document.createElement('div');
    card.className = 'thumb' + (i === 0 ? ' active' : '');
    card.dataset.page = i;
    card.style.aspectRatio = `${tw} / ${th}`;

    const canvas = document.createElement('canvas');
    canvas.width  = tw;
    canvas.height = th;

    const pg = document.createElement('div');
    pg.className = 'pg'; pg.textContent = i + 1;

    card.append(canvas, pg);
    card.addEventListener('click', () =>
      pagesInner.children[i]?.scrollIntoView({ behavior: 'smooth', block: 'start' })
    );
    thumbsEl.appendChild(card);
    thumbObserver.observe(card);
  }
}

const thumbObserver = new IntersectionObserver(entries => {
  entries.forEach(async entry => {
    if (!entry.isIntersecting) return;
    const card   = entry.target;
    const canvas = card.querySelector('canvas');
    if (canvas.dataset.done) return;
    canvas.dataset.done = '1';
    thumbObserver.unobserve(card);
    const n        = +card.dataset.page;
    const page     = await pdfDoc.getPage(n + 1);
    const viewport = page.getViewport({ scale: THUMB_SCALE });
    const dpr      = window.devicePixelRatio || 1;
    canvas.width  = Math.round(viewport.width  * dpr);
    canvas.height = Math.round(viewport.height * dpr);
    canvas.style.width  = Math.round(viewport.width)  + 'px';
    canvas.style.height = Math.round(viewport.height) + 'px';
    const ctx = canvas.getContext('2d');
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    await page.render({ canvasContext: ctx, viewport }).promise;
  });
}, { root: thumbsEl, rootMargin: '300px' });

// ── Zoom ──────────────────────────────────────────────────────
function fitZoom() {
  renderScale = Math.max(0.25, Math.min(4.0,
    (pagesEl.clientWidth - 48) / pdfPtW));
  zoomSel.value = 'fit';
  applyZoom();
}

function setZoom(scale) {
  renderScale = Math.max(0.1, Math.min(4.0, scale));
  const val = String(renderScale);
  zoomSel.value = [...zoomSel.options].some(o => o.value === val) ? val : 'fit';
  applyZoom();
}

function applyZoom() {
  clearSel();
  // Cancel in-progress renders
  for (const [, task] of renderTasks) try { task.cancel(); } catch (_) {}
  renderTasks.clear();
  rendered.clear();

  // Resize all placeholders
  const cssW = Math.round(pdfPtW * renderScale);
  const cssH = Math.round(pdfPtH * renderScale);
  document.querySelectorAll('.page-wrap').forEach(wrap => {
    wrap.classList.add('loading');
    wrap.style.width  = cssW + 'px';
    wrap.style.height = cssH + 'px';
  });

  // Re-render pages currently near the viewport
  document.querySelectorAll('.page-wrap').forEach(wrap => {
    const r = wrap.getBoundingClientRect();
    if (r.top < window.innerHeight + 800 && r.bottom > -800)
      renderPage(wrap);
  });
}

document.getElementById('btn-zoom-in').addEventListener('click', () => {
  const next = ZOOM_STOPS.find(z => z > renderScale + 0.01);
  setZoom(next ?? 3.0);
});
document.getElementById('btn-zoom-out').addEventListener('click', () => {
  const prev = [...ZOOM_STOPS].reverse().find(z => z < renderScale - 0.01);
  setZoom(prev ?? 0.25);
});
zoomSel.addEventListener('change', e => {
  if (e.target.value === 'fit') fitZoom(); else setZoom(+e.target.value);
});

pagesEl.addEventListener('wheel', e => {
  if (!e.ctrlKey && !e.metaKey) return;
  e.preventDefault();
  if (e.deltaY < 0) {
    const next = ZOOM_STOPS.find(z => z > renderScale + 0.01);
    setZoom(next ?? 3.0);
  } else {
    const prev = [...ZOOM_STOPS].reverse().find(z => z < renderScale - 0.01);
    setZoom(prev ?? 0.25);
  }
}, { passive: false });

window.addEventListener('resize', () => { if (zoomSel.value === 'fit') fitZoom(); });

// ── Selection drag ────────────────────────────────────────────
function relPt(e, canvas) {
  const r = canvas.getBoundingClientRect();
  return {
    x: Math.max(0, Math.min(e.clientX - r.left, r.width)),
    y: Math.max(0, Math.min(e.clientY - r.top,  r.height)),
  };
}

pagesEl.addEventListener('mousedown', e => {
  const wrap = e.target.closest('.page-wrap');
  if (!wrap || wrap.classList.contains('loading')) return;
  e.preventDefault();
  clearSel();
  dragPage        = +wrap.dataset.page;
  activeSelCanvas = wrap.querySelector('.sel-canvas');
  activeSelCtx    = activeSelCanvas.getContext('2d');
  dragStart       = relPt(e, activeSelCanvas);
  dragging        = true;
});

window.addEventListener('mousemove', e => {
  if (!dragging) return;
  const p = relPt(e, activeSelCanvas);
  sel = { page: dragPage, x0: dragStart.x, y0: dragStart.y, x1: p.x, y1: p.y };
  redrawSel();
  coordsEl.textContent =
    `${Math.round(Math.abs(sel.x1-sel.x0)/renderScale)} × ` +
    `${Math.round(Math.abs(sel.y1-sel.y0)/renderScale)} pts`;
});

window.addEventListener('mouseup', () => {
  if (!dragging) return;
  dragging = false;
  if (!sel) return;
  if (Math.abs(sel.x1-sel.x0) < 8 || Math.abs(sel.y1-sel.y0) < 8) {
    clearSel(); return;
  }
  btnExt.disabled = false; btnClr.disabled = false;
  const pW = Math.round(Math.abs(sel.x1-sel.x0)/renderScale);
  const pH = Math.round(Math.abs(sel.y1-sel.y0)/renderScale);
  coordsEl.textContent = `Selected p${sel.page+1}: ${pW} × ${pH} pts — press Extract`;
});

function redrawSel() {
  if (!activeSelCtx || !sel) return;
  const { x0, y0, x1, y1 } = sel;
  const x = Math.min(x0,x1), y = Math.min(y0,y1);
  const w = Math.abs(x1-x0), h = Math.abs(y1-y0);
  const cv = activeSelCanvas;
  activeSelCtx.clearRect(0, 0, cv.width, cv.height);
  activeSelCtx.fillStyle = 'rgba(0,0,0,0.45)';
  activeSelCtx.fillRect(0, 0, cv.width, cv.height);
  activeSelCtx.clearRect(x, y, w, h);
  activeSelCtx.strokeStyle = '#38bdf8';
  activeSelCtx.lineWidth   = 2;
  activeSelCtx.strokeRect(x, y, w, h);
  activeSelCtx.fillStyle = '#38bdf8';
  for (const [cx,cy] of [[x,y],[x+w,y],[x,y+h],[x+w,y+h]]) {
    activeSelCtx.beginPath();
    activeSelCtx.arc(cx, cy, 5, 0, Math.PI*2);
    activeSelCtx.fill();
  }
}

function clearSel() {
  if (activeSelCtx && activeSelCanvas)
    activeSelCtx.clearRect(0, 0, activeSelCanvas.width, activeSelCanvas.height);
  sel = null; dragging = false;
  activeSelCanvas = activeSelCtx = null; dragPage = -1;
  btnExt.disabled = true; btnClr.disabled = true;
  coordsEl.textContent = 'Drag on any page to select a region';
}

btnClr.addEventListener('click', clearSel);

// ── Extract ───────────────────────────────────────────────────
btnExt.addEventListener('click', doExtract);

async function doExtract() {
  if (!sel) return;
  btnExt.textContent = 'Extracting…'; btnExt.disabled = true;
  const resp = await fetch('/api/extract', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      page: sel.page, scale: renderScale,
      x0: sel.x0, y0: sel.y0, x1: sel.x1, y1: sel.y1,
    }),
  });
  addExtCard(await resp.json());
  btnExt.textContent = '⬇ Extract';
  clearSel();
}

function addExtCard(data) {
  extBar.hidden = false; extCount++;
  const card = document.createElement('div');
  card.className = 'ext-card'; card.dataset.file = data.filename;

  const imgWrap = document.createElement('div');
  imgWrap.className = 'ext-card-img';
  const eimg = document.createElement('img');
  eimg.src = data.url; eimg.alt = data.filename;
  imgWrap.appendChild(eimg);

  const foot = document.createElement('div');
  foot.className = 'ext-card-foot';
  const lbl = document.createElement('span');
  lbl.className = 'lbl'; lbl.textContent = `p${data.page} · #${extCount}`; lbl.title = data.filename;

  const dlBtn = document.createElement('button');
  dlBtn.className = 'btn-dl'; dlBtn.textContent = '↓'; dlBtn.title = 'Download';
  dlBtn.addEventListener('click', () =>
    Object.assign(document.createElement('a'),
      { href: data.url, download: data.filename }).click()
  );

  const delBtn = document.createElement('button');
  delBtn.className = 'btn-del'; delBtn.textContent = '×'; delBtn.title = 'Delete';
  delBtn.addEventListener('click', async () => {
    await fetch(data.url, { method: 'DELETE' });
    card.remove();
    if (!extList.children.length) extBar.hidden = true;
  });

  foot.append(lbl, dlBtn, delBtn);
  card.append(imgWrap, foot);
  extList.prepend(card);
}

document.getElementById('btn-clear-all').addEventListener('click', async () => {
  const cards = [...extList.querySelectorAll('.ext-card')];
  await Promise.all(cards.map(c => fetch(`/api/img/${c.dataset.file}`, { method: 'DELETE' })));
  extList.innerHTML = ''; extBar.hidden = true; extCount = 0;
});

// ── Keys ─────────────────────────────────────────────────────
window.addEventListener('keydown', e => {
  if (e.key === 'Escape') clearSel();
  if (e.key === 'Enter' && !btnExt.disabled) doExtract();
  if (e.ctrlKey || e.metaKey) {
    if (e.key === '=' || e.key === '+') {
      e.preventDefault();
      const next = ZOOM_STOPS.find(z => z > renderScale + 0.01);
      setZoom(next ?? 3.0);
    } else if (e.key === '-') {
      e.preventDefault();
      const prev = [...ZOOM_STOPS].reverse().find(z => z < renderScale - 0.01);
      setZoom(prev ?? 0.25);
    } else if (e.key === '0') {
      e.preventDefault(); fitZoom();
    }
  }
});
</script>
</body>
</html>"""


def free_port(start=8765, end=8775):
    for p in range(start, end):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", p)); return p
            except OSError:
                continue
    raise RuntimeError("No free port available")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 pdfViewer.py <file.pdf>")
        sys.exit(1)

    PDF_PATH = os.path.abspath(sys.argv[1])
    if not os.path.isfile(PDF_PATH):
        sys.exit(f"File not found: {PDF_PATH}")

    doc = fitz.open(PDF_PATH)
    EXTRACT_DIR = tempfile.mkdtemp(prefix="pdf_extracts_")

    port   = free_port()
    import socketserver
    server = type("ThreadedServer",
                  (socketserver.ThreadingMixIn, HTTPServer), {})(
                  ("127.0.0.1", port), Handler)
    url = f"http://127.0.0.1:{port}"

    print(f"PDF      : {os.path.basename(PDF_PATH)}  ({len(doc)} pages)")
    print(f"URL      : {url}")
    print(f"Extracts : {EXTRACT_DIR}")
    print("Stop     : Ctrl+C")

    threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")
