import os
import shutil
import json
import re
import asyncio
import time
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Cookie, Header, Depends, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse
import uvicorn
import fitz # PyMuPDF
from pptx import Presentation
from pptx.util import Pt

# Import functions from existing files
from main import extract_template, emu_to_pt
from convert import convert
from alt_text_excel import parse_alttext_excel

app = FastAPI(title="PPT Builder API")

# Enable CORS for frontend development
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOAD_DIR = os.path.abspath("uploads")
TEMPLATES_DIR = os.path.join(UPLOAD_DIR, "templates")
os.makedirs(TEMPLATES_DIR, exist_ok=True)

# In-memory store for session-specific state
session_states = {}

def get_session_id(session_id: str = Cookie(default="default"), x_session_id: str = Header(default=None)):
    """Retrieve session ID from cookie or header fallback."""
    if x_session_id:
        return x_session_id
    return session_id

def get_session_upload_dir(session_id: str) -> str:
    """Gets and prepares the upload directory specific to a session."""
    clean_id = re.sub(r'[^a-zA-Z0-9_-]', '', session_id)
    if not clean_id:
        clean_id = "default"
    path = os.path.join(UPLOAD_DIR, clean_id)
    os.makedirs(path, exist_ok=True)
    os.makedirs(os.path.join(path, "pdf_extracts"), exist_ok=True)
    os.makedirs(os.path.join(path, "rendered_slides"), exist_ok=True)
    return path

def get_session_state(session_id: str) -> dict:
    """Retrieve or initialize state for a session."""
    if session_id not in session_states:
        session_states[session_id] = {
            "template_pptx": None,
            "template_style_json": None,
            "content_pptx": None,
            "styled_pptx": None,
            "pdf_path": None,
            "captions": []
        }
    return session_states[session_id]

def reset_session_state(session_id: str):
    """Clear session state keys."""
    session_states[session_id] = {
        "template_pptx": None,
        "template_style_json": None,
        "content_pptx": None,
        "styled_pptx": None,
        "pdf_path": None,
        "captions": []
    }

def cleanup_session_dir(session_id: str):
    """Deletes the entire user-specific session folder and resets state."""
    try:
        session_upload_dir = get_session_upload_dir(session_id)
        if os.path.exists(session_upload_dir):
            shutil.rmtree(session_upload_dir, ignore_errors=True)
        reset_session_state(session_id)
    except Exception as e:
        print(f"Error during background session cleanup for {session_id}: {e}")

def cleanup_old_sessions():
    """Deletes session folders in UPLOAD_DIR that are older than 24 hours."""
    now = time.time()
    cutoff = now - 24 * 3600 # 24 hours ago
    
    if not os.path.exists(UPLOAD_DIR):
        return
        
    for item in os.listdir(UPLOAD_DIR):
        item_path = os.path.join(UPLOAD_DIR, item)
        # Skip the templates directory
        if item == "templates":
            continue
            
        if os.path.isdir(item_path):
            try:
                mtime = os.path.getmtime(item_path)
                if mtime < cutoff:
                    shutil.rmtree(item_path, ignore_errors=True)
                    if item in session_states:
                        del session_states[item]
            except Exception as e:
                print(f"Failed to check/cleanup old session folder {item}: {e}")

async def periodic_cleanup_loop():
    while True:
        try:
            cleanup_old_sessions()
        except Exception as e:
            print(f"Error in periodic cleanup task: {e}")
        # Run every hour
        await asyncio.sleep(3600)

@app.on_event("startup")
async def startup_event():
    asyncio.create_task(periodic_cleanup_loop())



def _span_style(span):
    """Return (bold, italic) booleans from a PyMuPDF span dict."""
    flags = span.get("flags", 0)
    font_name = span.get("font", "").lower()
    bold = bool(flags & 16) or "bold" in font_name
    italic = bool(flags & 2) or "italic" in font_name or "oblique" in font_name
    return bold, italic


def _block_runs(block):
    """Extract list of {text, bold, italic} runs from a PyMuPDF dict block."""
    runs = []
    for line in block.get("lines", []):
        for span in line.get("spans", []):
            text = span.get("text", "")
            if text:
                bold, italic = _span_style(span)
                runs.append({"text": text, "bold": bold, "italic": italic})
    return runs


def _block_plain(block):
    """Return plain text string from a PyMuPDF dict block."""
    return " ".join(
        span.get("text", "")
        for line in block.get("lines", [])
        for span in line.get("spans", [])
    ).strip()


def extract_pdf_captions(pdf_path):
    captions = []
    pattern = re.compile(r"^\s*(Figure|Fig\.|Table)\s+(\d+[-.\d]*)\b", re.IGNORECASE)
    credit_pattern = re.compile(
        r"^\s*(©|Copyright\b|Courtesy of\b|Source:\b|Source\b|Reproduced from\b|"
        r"Reproduced with permission\b|Data from\b|Courtesy\b|Permission\b|Used with permission\b)",
        re.IGNORECASE,
    )

    if not os.path.exists(pdf_path):
        return captions

    try:
        doc = fitz.open(pdf_path)
        for page_idx in range(len(doc)):
            page = doc[page_idx]
            # Use dict mode so we get per-span font flags
            page_data = page.get_text("dict")
            blocks = [b for b in page_data.get("blocks", []) if b.get("type") == 0]

            for b_idx, block in enumerate(blocks):
                plain = _block_plain(block)
                if not plain:
                    continue

                lines = plain.split("\n") if "\n" in plain else plain.splitlines()
                # Recompute from actual line structure
                lines = []
                for line in block.get("lines", []):
                    line_text = " ".join(s.get("text", "") for s in line.get("spans", [])).strip()
                    if line_text:
                        lines.append(line_text)

                for line_idx, line in enumerate(lines):
                    match = pattern.match(line.strip())
                    if match:
                        full_caption = " ".join(lines[line_idx:])
                        clean_text = " ".join(full_caption.split())

                        # Runs for this block (all spans = caption text)
                        cap_runs = _block_runs(block)

                        # Credit: look at the next text block
                        credit_text = ""
                        credit_runs = []
                        for nb in blocks[b_idx + 1: b_idx + 3]:
                            nb_plain = _block_plain(nb)
                            if credit_pattern.match(nb_plain):
                                credit_text = " ".join(nb_plain.split())
                                credit_runs = _block_runs(nb)
                            break

                        cap_type = match.group(1).capitalize()
                        if cap_type.startswith("Fig"):
                            cap_type = "Figure"

                        captions.append({
                            "id": f"cap_{page_idx}_{b_idx}_{line_idx}",
                            "page": page_idx + 1,
                            "label": f"{cap_type} {match.group(2)}",
                            "text": clean_text,
                            "runs": cap_runs,
                            "credit": credit_text,
                            "creditRuns": credit_runs,
                        })
                        break
        doc.close()
    except Exception as e:
        print("Error extracting PDF captions:", e)
    return captions

@app.post("/api/reset")
async def reset_session(session_id: str = Depends(get_session_id)):
    """Wipe all uploaded session files (keep templates) and reset server state."""
    try:
        session_upload_dir = get_session_upload_dir(session_id)
        # Remove all files and subdirs in session_upload_dir except the templates folder
        for item in os.listdir(session_upload_dir):
            item_path = os.path.join(session_upload_dir, item)
            if item == "templates":
                continue  # keep saved templates intact
            if os.path.isdir(item_path):
                shutil.rmtree(item_path, ignore_errors=True)
            else:
                try:
                    os.remove(item_path)
                except Exception:
                    pass
        # Recreate necessary subdirs
        os.makedirs(os.path.join(session_upload_dir, "pdf_extracts"), exist_ok=True)
        os.makedirs(os.path.join(session_upload_dir, "rendered_slides"), exist_ok=True)
        # Reset server state
        reset_session_state(session_id)
        return {"ok": True, "message": "Session reset. All uploaded files cleared."}
    except Exception as e:
        return {"ok": False, "detail": str(e)}

def render_ppt_to_pngs(pptx_path, upload_dir):
    pdf_path = os.path.join(upload_dir, "styled_output.pdf")
    if os.path.exists(pdf_path):
        try:
            os.remove(pdf_path)
        except Exception:
            pass
            
    applescript_code = f'''
    set pptxPosix to "{pptx_path}"
    set pdfPosix to "{pdf_path}"
    
    tell application "Microsoft PowerPoint"
        open posix file pptxPosix
        set activePres to active presentation
        save activePres in posix file pdfPosix as save as PDF
        close activePres saving no
    end tell
    '''
    
    import subprocess
    try:
        subprocess.run(["osascript", "-e", applescript_code], check=True, capture_output=True, text=True)
    except subprocess.CalledProcessError as e:
        print("AppleScript export to PDF failed:", e.stderr)
        raise e

    import fitz
    doc = fitz.open(pdf_path)
    output_dir = os.path.join(upload_dir, "rendered_slides")
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir, exist_ok=True)
    
    for i in range(len(doc)):
        page = doc[i]
        pix = page.get_pixmap(dpi=150)
        png_path = os.path.join(output_dir, f"slide_{i}.png")
        pix.save(png_path)

@app.post("/api/upload-template")
async def upload_template(file: UploadFile = File(...), session_id: str = Depends(get_session_id)):
    try:
        filename = file.filename
        path = os.path.join(TEMPLATES_DIR, filename)
        with open(path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Auto-repair the template on upload so downstream extraction /
        # conversion sees a clean file (adds missing <p:ph> idx/type,
        # creationId GUIDs, etc.).
        from template_repair import repair_template
        repair_report = repair_template(path)

        state = get_session_state(session_id)
        state["template_pptx"] = path

        # Extract styles using main.py logic
        styles = extract_template(path)
        style_json_filename = os.path.splitext(filename)[0] + "_styles.json"
        style_json_path = os.path.join(TEMPLATES_DIR, style_json_filename)
        with open(style_json_path, "w") as f:
            json.dump(styles, f, indent=2)

        state["template_style_json"] = style_json_path
        return {
            "ok": True,
            "styles": styles,
            "filename": filename,
            "repair": {
                "count": repair_report.count,
                "fixes": repair_report.fixes,
            },
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/templates")
async def get_templates(session_id: str = Depends(get_session_id)):
    try:
        templates = []
        if os.path.exists(TEMPLATES_DIR):
            for f in os.listdir(TEMPLATES_DIR):
                if f.endswith(".pptx"):
                    templates.append({
                        "name": os.path.splitext(f)[0],
                        "filename": f
                    })
        return {"ok": True, "templates": templates}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/select-template")
async def select_template(data: dict, session_id: str = Depends(get_session_id)):
    filename = data.get("filename")
    if not filename:
        raise HTTPException(status_code=400, detail="Missing filename")
    
    path = os.path.join(TEMPLATES_DIR, filename)
    style_json_filename = os.path.splitext(filename)[0] + "_styles.json"
    style_json_path = os.path.join(TEMPLATES_DIR, style_json_filename)
    
    if not os.path.exists(path) or not os.path.exists(style_json_path):
        raise HTTPException(status_code=404, detail="Template or styles not found")
        
    state = get_session_state(session_id)
    state["template_pptx"] = path
    state["template_style_json"] = style_json_path
    
    with open(style_json_path, "r") as f:
        styles = json.load(f)
        
    return {"ok": True, "filename": filename, "styles": styles}


@app.post("/api/upload-ppt")
async def upload_ppt(file: UploadFile = File(...), session_id: str = Depends(get_session_id)):
    try:
        session_upload_dir = get_session_upload_dir(session_id)
        path = os.path.join(session_upload_dir, "uploaded_content.pptx")
        with open(path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        state = get_session_state(session_id)
        state["content_pptx"] = path
        
        from main import extract_template
        slides_info = extract_template(path)
        
        return {"ok": True, "filename": file.filename, "slidesInfo": slides_info}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

def _figure_dest_name(fig: dict) -> str | None:
    """Normalize a figure name into the on-disk name insert_figure_placeholders looks up.

    "Figure 2.1" / "Figure2.1" → "figure 2.1.png". Names that don't parse as a
    figure/table reference fall back to a lowercased slug.
    """
    name = fig.get("name")
    filename = fig.get("filename")
    if not (name and filename):
        return None
    m = re.match(r"(figure|table)\s*([\d.-]+)", name, re.IGNORECASE)
    if m:
        return f"{m.group(1).lower()} {m.group(2)}.png"
    return f"{name.lower()}.png"


def _prepare_pdf_extracts(session_upload_dir: str, figures: list[dict]) -> None:
    """(Re)create pdf_extracts/ and copy figure PNGs in with normalized names."""
    extracts_dir = os.path.join(session_upload_dir, "pdf_extracts")
    if os.path.exists(extracts_dir):
        shutil.rmtree(extracts_dir)
    os.makedirs(extracts_dir, exist_ok=True)
    for fig in figures:
        dest_name = _figure_dest_name(fig)
        if not dest_name:
            continue
        src_path = os.path.join(session_upload_dir, fig["filename"])
        if os.path.exists(src_path):
            shutil.copy(src_path, os.path.join(extracts_dir, dest_name))


def _resolve_auto_inserted_list(used_figs: list[dict], figures: list[dict]) -> list[dict]:
    """Map insert_figure_placeholders' `used` records back to caller filenames."""
    name_to_filename: dict[str, str] = {}
    for fig in figures:
        dest_name = _figure_dest_name(fig)
        if dest_name:
            name_to_filename.setdefault(dest_name, fig["filename"])
    out = []
    for uf in used_figs:
        orig = name_to_filename.get(uf["dest_name"])
        if orig:
            out.append({
                "filename": orig,
                "slideIndex": uf["slideIndex"],
                "shapeIndex": uf["shapeIndex"],
            })
    return out


def _extract_and_attach_image_urls(
    session_id: str, session_upload_dir: str, output_path: str, slides_info: dict
) -> None:
    """Persist each picture's blob to disk and inject its URL into slides_info."""
    prs = Presentation(output_path)
    urls: dict[str, str] = {}
    for s_idx, slide in enumerate(prs.slides):
        for sh_idx, shape in enumerate(slide.shapes):
            if not (hasattr(shape, "shape_type") and int(shape.shape_type) == 13):
                continue
            try:
                img = shape.image
                img_filename = f"slide_{s_idx}_shape_{sh_idx}.{img.ext}"
                with open(os.path.join(session_upload_dir, img_filename), "wb") as f:
                    f.write(img.blob)
                urls[f"{s_idx}_{sh_idx}"] = f"/api/media/{session_id}/{img_filename}"
            except Exception:
                pass
    for s_idx, slide_data in enumerate(slides_info.get("slides", [])):
        for sh_idx, shape_data in enumerate(slide_data.get("shapes", [])):
            key = f"{s_idx}_{sh_idx}"
            if key in urls:
                shape_data["imageUrl"] = urls[key]


@app.post("/api/process-ppt")
async def process_ppt(payload: dict = None, session_id: str = Depends(get_session_id)):
    print("PROCESS_PPT payload:", payload)
    state = get_session_state(session_id)
    if not state["content_pptx"] or not state["template_style_json"]:
        raise HTTPException(status_code=400, detail="Missing uploaded PPTX or Template style JSON")

    mode = (payload or {}).get("mode", "restyle")
    session_upload_dir = get_session_upload_dir(session_id)
    output_path = os.path.join(session_upload_dir, "styled_output.pptx")
    figures = (payload or {}).get("figures", []) or []
    include_figure_captions = (payload or {}).get("include_figure_captions", True)
    include_table_captions = (payload or {}).get("include_table_captions", True)

    if mode == "master_swap":
        # Rebuild the input deck inside a fresh copy of the template so the
        # output inherits the template's master + layouts + theme, then run
        # the same figure-insertion pass restyle uses. When `figures` is
        # empty the figure pass is a no-op — master_swap alone continues to
        # produce the same output it always did.
        from master_swap import apply_template
        from convert import insert_figure_placeholders

        try:
            apply_template(
                state["content_pptx"], state["template_pptx"], output_path
            )
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"master_swap failed: {e}")

        used_figs: list[dict] = []
        if figures:
            try:
                _prepare_pdf_extracts(session_upload_dir, figures)
                prs = Presentation(output_path)
                used_figs = insert_figure_placeholders(
                    prs, session_upload_dir,
                    figures_metadata=figures,
                    template=state["template_style_json"],
                    include_figure_captions=include_figure_captions,
                    include_table_captions=include_table_captions,
                )
                prs.save(output_path)
            except Exception as e:
                raise HTTPException(status_code=500, detail=f"figure insertion failed: {e}")

        state["styled_pptx"] = output_path
        auto_inserted_list = _resolve_auto_inserted_list(used_figs, figures)
        state["auto_inserted_list"] = auto_inserted_list

        slides_info = extract_template(output_path)
        _extract_and_attach_image_urls(session_id, session_upload_dir, output_path, slides_info)
        return {"ok": True, "slidesInfo": slides_info, "autoInserted": auto_inserted_list}

    try:
        _prepare_pdf_extracts(session_upload_dir, figures)
        # Run conversion style formatting and automatic figure insertion
        used_figs = convert(
            state["content_pptx"], state["template_style_json"], output_path,
            apply_geometry=True, figures_metadata=figures,
            include_figure_captions=include_figure_captions,
            include_table_captions=include_table_captions,
        )
        state["styled_pptx"] = output_path
        auto_inserted_list = _resolve_auto_inserted_list(used_figs, figures)
        state["auto_inserted_list"] = auto_inserted_list

        slides_info = extract_template(output_path)
        _extract_and_attach_image_urls(session_id, session_upload_dir, output_path, slides_info)
        return {"ok": True, "slidesInfo": slides_info, "autoInserted": auto_inserted_list}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/upload-pdf")
async def upload_pdf(file: UploadFile = File(...), session_id: str = Depends(get_session_id)):
    try:
        session_upload_dir = get_session_upload_dir(session_id)
        path = os.path.join(session_upload_dir, "uploaded_document.pdf")
        with open(path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        state = get_session_state(session_id)
        state["pdf_path"] = path
        
        # Open PDF to get details
        doc = fitz.open(path)
        page_count = len(doc)
        filename = file.filename
        doc.close()
        
        # Extract captions
        captions = extract_pdf_captions(path)
        state["captions"] = captions
        
        return {"ok": True, "filename": filename, "pageCount": page_count, "captions": captions}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/parse-alttext-excel")
async def parse_alttext_excel_endpoint(file: UploadFile = File(...), session_id: str = Depends(get_session_id)):
    """Parse a client-supplied alt-text Excel and return normalized entries.

    Fully generic: the parser locates columns by header name, so any
    customer's workbook layout works as long as it exposes Chapter Number,
    Element Number, and Alt Text - Short columns. No customer branching.
    """
    session_upload_dir = get_session_upload_dir(session_id)
    path = os.path.join(session_upload_dir, "alt_text.xlsx")
    try:
        with open(path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"failed to save upload: {e}")

    try:
        entries = parse_alttext_excel(path)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"failed to parse workbook: {e}")

    state = get_session_state(session_id)
    state["alt_text_entries"] = entries
    return {"ok": True, "entries": entries, "count": len(entries)}


@app.get("/api/pdf/captions")
async def get_pdf_captions(session_id: str = Depends(get_session_id)):
    state = get_session_state(session_id)
    if not state["pdf_path"]:
        raise HTTPException(status_code=404, detail="No PDF uploaded")
    if not state.get("captions"):
        state["captions"] = extract_pdf_captions(state["pdf_path"])
    return {"ok": True, "captions": state["captions"]}

@app.get("/api/pdf/info")
async def get_pdf_info(session_id: str = Depends(get_session_id)):
    state = get_session_state(session_id)
    if not state["pdf_path"]:
        raise HTTPException(status_code=404, detail="No PDF uploaded")
    try:
        doc = fitz.open(state["pdf_path"])
        count = len(doc)
        filename = os.path.basename(state["pdf_path"])
        doc.close()
        return {"filename": filename, "count": count}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/pdf/file")
async def get_pdf_file(session_id: str = Depends(get_session_id)):
    state = get_session_state(session_id)
    if not state["pdf_path"]:
        raise HTTPException(status_code=404, detail="No PDF uploaded")
    return FileResponse(state["pdf_path"], media_type="application/pdf")

@app.post("/api/extract")
async def extract_crop(
    page: int = Form(...),
    scale: float = Form(...),
    x0: float = Form(...),
    y0: float = Form(...),
    x1: float = Form(...),
    y1: float = Form(...),
    session_id: str = Depends(get_session_id)
):
    state = get_session_state(session_id)
    if not state["pdf_path"]:
        raise HTTPException(status_code=404, detail="No PDF uploaded")
    try:
        session_upload_dir = get_session_upload_dir(session_id)
        doc = fitz.open(state["pdf_path"])
        pdf_page = doc[page]
        
        # Bounding box
        clip = fitz.Rect(
            min(x0, x1) / scale,
            min(y0, y1) / scale,
            max(x0, x1) / scale,
            max(y0, y1) / scale
        )
        
        # Render crop
        pix = pdf_page.get_pixmap(
            matrix=fitz.Matrix(3.0, 3.0),
            clip=clip,
            alpha=False
        )
        png_data = pix.tobytes("png")
        doc.close()
        
        # Save crop image into session_upload_dir
        existing = [f for f in os.listdir(session_upload_dir) if f.startswith("crop_") and f.endswith(".png")]
        idx = len(existing) + 1
        filename = f"crop_p{page+1}_{idx:03d}.png"
        filepath = os.path.join(session_upload_dir, filename)
        
        with open(filepath, "wb") as f:
            f.write(png_data)
            
        return {"filename": filename, "url": f"/api/media/{session_id}/{filename}", "page": page + 1}
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/add-image")
async def add_image(
    slide_index: int = Form(...),
    image_name: str = Form(...),
    x_pt: float = Form(...),
    y_pt: float = Form(...),
    w_pt: float = Form(...),
    h_pt: float = Form(...),
    caption: str = Form(None),
    caption_runs: str = Form(None),
    alt_text: str = Form(None),
    session_id: str = Depends(get_session_id)
):
    state = get_session_state(session_id)
    target_pptx = state["styled_pptx"] or state["content_pptx"]
    if not target_pptx:
        raise HTTPException(status_code=400, detail="No PPTX presentation available to modify")
        
    session_upload_dir = get_session_upload_dir(session_id)
    # Look in pdf_extracts first, then fall back to session_upload_dir root
    extracts_dir = os.path.join(session_upload_dir, "pdf_extracts")
    image_path = os.path.join(extracts_dir, image_name)
    if not os.path.exists(image_path):
        image_path = os.path.join(session_upload_dir, image_name)
    if not os.path.exists(image_path):
        raise HTTPException(status_code=404, detail=f"Image {image_name} not found")
        
    try:
        prs = Presentation(target_pptx)
        if slide_index < 0 or slide_index >= len(prs.slides):
            raise HTTPException(status_code=400, detail="Invalid slide index")
            
        slide = prs.slides[slide_index]
        
        # Convert points to EMUs (1 pt = 12700 EMUs)
        EMU_PER_PT = 12700
        left = Pt(x_pt)
        top = Pt(y_pt)
        width = Pt(w_pt)
        height = Pt(h_pt)
        
        # Insert image
        pic = slide.shapes.add_picture(image_path, left, top, width, height)

        # Stamp WCAG-required alt text on the inserted picture. python-pptx
        # defaults `descr` to the source image filename on add_picture, which
        # accessibility tooling flags as a false positive; always overwrite.
        # Same idiom used by the batch convert() figure-insertion path so
        # both entry points behave identically.
        try:
            cNvPr = pic._element.nvPicPr.cNvPr
            trimmed = (alt_text or "").strip()
            if trimmed:
                cNvPr.set('descr', trimmed)
            else:
                cNvPr.attrib.pop('descr', None)
        except Exception:
            pass

        if caption:
            cap_top = top + height + Pt(10)
            cap_height = Pt(35)
            if cap_top + cap_height > prs.slide_height:
                cap_top = prs.slide_height - cap_height - Pt(10)
            txBox = slide.shapes.add_textbox(left, cap_top, width, cap_height)
            tf = txBox.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            runs_data = None
            if caption_runs:
                try:
                    runs_data = json.loads(caption_runs)
                except Exception:
                    pass
            if runs_data:
                for run_data in runs_data:
                    run = p.add_run()
                    run.text = run_data.get("text", "")
                    run.font.size = Pt(10)
                    run.font.bold = run_data.get("bold", False)
                    run.font.italic = run_data.get("italic", True)
            else:
                run = p.add_run()
                run.text = caption
                run.font.size = Pt(10)
                run.font.italic = True
            
        prs.save(target_pptx)
        
        # Re-extract slide info to reflect updates
        slides_info = extract_template(target_pptx)
        
        # Rescan and expose image paths
        extracted_images = {}
        for s_idx, s in enumerate(prs.slides):
            for sh_idx, shape in enumerate(s.shapes):
                if hasattr(shape, "shape_type") and int(shape.shape_type) == 13: # PICTURE
                    try:
                        img = shape.image
                        ext = img.ext
                        img_filename = f"slide_{s_idx}_shape_{sh_idx}.{ext}"
                        img_path = os.path.join(session_upload_dir, img_filename)
                        with open(img_path, "wb") as f:
                            f.write(img.blob)
                        extracted_images[f"{s_idx}_{sh_idx}"] = f"/api/media/{session_id}/{img_filename}"
                    except Exception:
                        pass
                        
        for s_idx, slide_data in enumerate(slides_info.get("slides", [])):
            for sh_idx, shape_data in enumerate(slide_data.get("shapes", [])):
                key = f"{s_idx}_{sh_idx}"
                if key in extracted_images:
                    shape_data["imageUrl"] = extracted_images[key]
                    
        return {"ok": True, "slidesInfo": slides_info}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/download")
async def download_pptx(session_id: str = Depends(get_session_id)):
    state = get_session_state(session_id)
    target_pptx = state["styled_pptx"] or state["content_pptx"]
    if not target_pptx or not os.path.exists(target_pptx):
        raise HTTPException(status_code=404, detail="No output PPTX available to download")
    return FileResponse(
        target_pptx,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename="styled_presentation.pptx"
    )


@app.get("/api/download-excel")
async def download_excel(customerName: str = "", projectName: str = "", session_id: str = Depends(get_session_id)):
    state = get_session_state(session_id)
    input_path = state.get("content_pptx")
    session_upload_dir = get_session_upload_dir(session_id)
    output_path = state.get("styled_pptx") or os.path.join(session_upload_dir, "styled_output.pptx")
    
    if not input_path or not os.path.exists(input_path):
        raise HTTPException(status_code=400, detail="Missing source presentation file.")
    if not os.path.exists(output_path):
        output_path = input_path
        
    excel_path = os.path.join(session_upload_dir, "compilation_report.xlsx")
    extracts_dir = os.path.join(session_upload_dir, "pdf_extracts")
    
    try:
        from excel_report import create_excel_report
        create_excel_report(
            input_pptx=input_path,
            output_pptx=output_path,
            extracts_dir=extracts_dir,
            customer_name=customerName,
            project_name=projectName,
            output_excel_path=excel_path
        )
        
        return FileResponse(
            excel_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="compilation_report.xlsx"
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to generate Excel report: {str(e)}")



@app.get("/api/report-data")
async def get_report_data(session_id: str = Depends(get_session_id)):
    state = get_session_state(session_id)
    session_upload_dir = get_session_upload_dir(session_id)
    output_path = os.path.join(session_upload_dir, "styled_output.pptx")
    if not state["content_pptx"] or not os.path.exists(output_path):
        return {"ok": False, "changes": []}
    try:
        from report import collect_changes
        changes = collect_changes(state["content_pptx"], output_path)
        return {"ok": True, "changes": changes}
    except Exception as e:
        return {"ok": False, "detail": str(e), "changes": []}

@app.get("/api/figure-diagnostics")
async def get_figure_diagnostics(session_id: str = Depends(get_session_id)):
    """Return missing (requested but not cropped) and unplaced (cropped but no placeholder) figures."""
    state = get_session_state(session_id)
    input_path = state.get("content_pptx")
    if not input_path or not os.path.exists(input_path):
        return {"ok": False, "missing": [], "unplaced": []}
    try:
        from report import collect_figure_diagnostics
        session_upload_dir = get_session_upload_dir(session_id)
        extracts_dir = os.path.join(session_upload_dir, "pdf_extracts")
        missing, unplaced = collect_figure_diagnostics(input_path, extracts_dir)
        return {"ok": True, "missing": missing, "unplaced": unplaced}
    except Exception as e:
        return {"ok": False, "detail": str(e), "missing": [], "unplaced": []}

@app.get("/api/accessibility-report")
async def get_accessibility_report(session_id: str = Depends(get_session_id)):
    state = get_session_state(session_id)
    session_upload_dir = get_session_upload_dir(session_id)
    output_path = state.get("styled_pptx") or os.path.join(session_upload_dir, "styled_output.pptx")
    if not os.path.exists(output_path):
        return {"ok": False, "issues": []}
    try:
        from accessibility import check_ppt_accessibility
        issues = check_ppt_accessibility(output_path)
        return {"ok": True, "issues": issues}
    except Exception as e:
        return {"ok": False, "detail": str(e), "issues": []}

@app.get("/api/customers")
async def get_customers():
    try:
        import json
        with open("customers.json", "r") as f:
            data = json.load(f)
        return {"ok": True, "customers": data}
    except Exception as e:
        return {"ok": False, "detail": str(e), "customers": []}

# Static media serving for images
app.mount("/api/media", StaticFiles(directory=UPLOAD_DIR), name="media")

# Serve React frontend static files if they exist
frontend_dist_path = os.path.abspath("frontend/dist")

@app.get("/")
async def read_index():
    index_path = os.path.join(frontend_dist_path, "index.html")
    if os.path.exists(index_path):
        return FileResponse(index_path)
    return {"message": "PPTBuilder API is running. Build the frontend to view the UI."}

if os.path.exists(frontend_dist_path):
    app.mount("/", StaticFiles(directory=frontend_dist_path, html=True), name="frontend")

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8001)
