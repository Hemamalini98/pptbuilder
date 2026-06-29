import os
import shutil
import json
import re
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
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

# State to keep track of currently active paths
state = {
    "template_pptx": None,
    "template_style_json": None,
    "content_pptx": None,
    "styled_pptx": None,
    "pdf_path": None
}

def reset_state():
    """Clear all session state keys."""
    for key in state:
        state[key] = None

@app.post("/api/reset")
async def reset_session():
    """Wipe all uploaded session files (keep templates) and reset server state."""
    try:
        # Remove all files and subdirs in UPLOAD_DIR except the templates folder
        for item in os.listdir(UPLOAD_DIR):
            item_path = os.path.join(UPLOAD_DIR, item)
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
        os.makedirs(os.path.join(UPLOAD_DIR, "pdf_extracts"), exist_ok=True)
        os.makedirs(os.path.join(UPLOAD_DIR, "rendered_slides"), exist_ok=True)
        # Reset server state
        reset_state()
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
async def upload_template(file: UploadFile = File(...)):
    try:
        filename = file.filename
        path = os.path.join(TEMPLATES_DIR, filename)
        with open(path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        state["template_pptx"] = path
        
        # Extract styles using main.py logic
        styles = extract_template(path)
        style_json_filename = os.path.splitext(filename)[0] + "_styles.json"
        style_json_path = os.path.join(TEMPLATES_DIR, style_json_filename)
        with open(style_json_path, "w") as f:
            json.dump(styles, f, indent=2)
            
        state["template_style_json"] = style_json_path
        return {"ok": True, "styles": styles, "filename": filename}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/templates")
async def get_templates():
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
async def select_template(data: dict):
    filename = data.get("filename")
    if not filename:
        raise HTTPException(status_code=400, detail="Missing filename")
    
    path = os.path.join(TEMPLATES_DIR, filename)
    style_json_filename = os.path.splitext(filename)[0] + "_styles.json"
    style_json_path = os.path.join(TEMPLATES_DIR, style_json_filename)
    
    if not os.path.exists(path) or not os.path.exists(style_json_path):
        raise HTTPException(status_code=404, detail="Template or styles not found")
        
    state["template_pptx"] = path
    state["template_style_json"] = style_json_path
    
    with open(style_json_path, "r") as f:
        styles = json.load(f)
        
    return {"ok": True, "filename": filename, "styles": styles}


@app.post("/api/upload-ppt")
async def upload_ppt(file: UploadFile = File(...)):
    try:
        path = os.path.join(UPLOAD_DIR, "uploaded_content.pptx")
        with open(path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        state["content_pptx"] = path
        return {"ok": True, "filename": file.filename}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/process-ppt")
async def process_ppt(payload: dict = None):
    print("PROCESS_PPT payload:", payload)
    if not state["content_pptx"] or not state["template_style_json"]:
        raise HTTPException(status_code=400, detail="Missing uploaded PPTX or Template style JSON")
    
    try:
        # Recreate pdf_extracts folder
        extracts_dir = os.path.join(UPLOAD_DIR, "pdf_extracts")
        if os.path.exists(extracts_dir):
            shutil.rmtree(extracts_dir)
        os.makedirs(extracts_dir, exist_ok=True)

        # Copy crops with proper names
        figures = (payload or {}).get("figures", [])
        for fig in figures:
            name = fig.get("name")
            filename = fig.get("filename")
            if name and filename:
                src_path = os.path.join(UPLOAD_DIR, filename)
                if os.path.exists(src_path):
                    # Match name (e.g. "Figure 1.1" or "Figure1.1") and write it with a space
                    m = re.match(r"(figure|table)\s*([\d.]+)", name, re.IGNORECASE)
                    if m:
                        dest_name = f"{m.group(1).lower()} {m.group(2)}.png"
                    else:
                        dest_name = f"{name.lower()}.png"
                    shutil.copy(src_path, os.path.join(extracts_dir, dest_name))

        output_path = os.path.join(UPLOAD_DIR, "styled_output.pptx")
        # Run conversion style formatting and automatic figure insertion
        used_figs = convert(state["content_pptx"], state["template_style_json"], output_path, apply_geometry=True)
        state["styled_pptx"] = output_path

        # Resolve auto-inserted figures back to original filenames
        auto_inserted_list = []
        for uf in used_figs:
            dest_name = uf["dest_name"]
            orig_filename = None
            for fig in figures:
                name = fig.get("name")
                filename = fig.get("filename")
                if name and filename:
                    m = re.match(r"(figure|table)\s*([\d.]+)", name, re.IGNORECASE)
                    if m:
                        candidate = f"{m.group(1).lower()} {m.group(2)}.png"
                    else:
                        candidate = f"{name.lower()}.png"
                    if candidate == dest_name:
                        orig_filename = filename
                        break
            if orig_filename:
                auto_inserted_list.append({
                    "filename": orig_filename,
                    "slideIndex": uf["slideIndex"],
                    "shapeIndex": uf["shapeIndex"]
                })
        
        # Save auto inserted list into local variable for returning in the response
        state["auto_inserted_list"] = auto_inserted_list

        # Generate style difference and figure placement diagnostics report
        try:
            from report import generate_report
            report_path = os.path.join(UPLOAD_DIR, "change_report.html")
            generate_report(state["content_pptx"], output_path, report_path)
            print("Successfully generated style change report at:", report_path)
        except Exception as e:
            print("Failed to generate style change report:", e)
        

        
        # Extract slides structure of the output file
        slides_info = extract_template(output_path)
        
        # Extract image components if any so we can display them
        # Go through shapes to extract image blobs
        prs = Presentation(output_path)
        extracted_images = {}
        for s_idx, slide in enumerate(prs.slides):
            for sh_idx, shape in enumerate(slide.shapes):
                if hasattr(shape, "shape_type") and int(shape.shape_type) == 13: # PICTURE
                    try:
                        img = shape.image
                        ext = img.ext
                        img_filename = f"slide_{s_idx}_shape_{sh_idx}.{ext}"
                        img_path = os.path.join(UPLOAD_DIR, img_filename)
                        with open(img_path, "wb") as f:
                            f.write(img.blob)
                        extracted_images[f"{s_idx}_{sh_idx}"] = f"/api/media/{img_filename}"
                    except Exception:
                        pass
        
        # Inject image URLs into the shapes json
        for s_idx, slide_data in enumerate(slides_info.get("slides", [])):
            for sh_idx, shape_data in enumerate(slide_data.get("shapes", [])):
                key = f"{s_idx}_{sh_idx}"
                if key in extracted_images:
                    shape_data["imageUrl"] = extracted_images[key]

        return {
            "ok": True, 
            "slidesInfo": slides_info, 
            "autoInserted": state.get("auto_inserted_list", [])
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/upload-pdf")
async def upload_pdf(file: UploadFile = File(...)):
    try:
        path = os.path.join(UPLOAD_DIR, "uploaded_document.pdf")
        with open(path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        state["pdf_path"] = path
        
        # Open PDF to get details
        doc = fitz.open(path)
        page_count = len(doc)
        filename = file.filename
        doc.close()
        
        return {"ok": True, "filename": filename, "pageCount": page_count}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/pdf/info")
async def get_pdf_info():
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
async def get_pdf_file():
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
    y1: float = Form(...)
):
    if not state["pdf_path"]:
        raise HTTPException(status_code=404, detail="No PDF uploaded")
    try:
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
        
        # Save crop image into UPLOAD_DIR
        existing = [f for f in os.listdir(UPLOAD_DIR) if f.startswith("crop_") and f.endswith(".png")]
        idx = len(existing) + 1
        filename = f"crop_p{page+1}_{idx:03d}.png"
        filepath = os.path.join(UPLOAD_DIR, filename)
        
        with open(filepath, "wb") as f:
            f.write(png_data)
            
        return {"filename": filename, "url": f"/api/media/{filename}", "page": page + 1}
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
    h_pt: float = Form(...)
):
    target_pptx = state["styled_pptx"] or state["content_pptx"]
    if not target_pptx:
        raise HTTPException(status_code=400, detail="No PPTX presentation available to modify")
        
    # Look in pdf_extracts first, then fall back to UPLOAD_DIR root
    extracts_dir = os.path.join(UPLOAD_DIR, "pdf_extracts")
    image_path = os.path.join(extracts_dir, image_name)
    if not os.path.exists(image_path):
        image_path = os.path.join(UPLOAD_DIR, image_name)
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
        slide.shapes.add_picture(image_path, left, top, width, height)
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
                        img_path = os.path.join(UPLOAD_DIR, img_filename)
                        with open(img_path, "wb") as f:
                            f.write(img.blob)
                        extracted_images[f"{s_idx}_{sh_idx}"] = f"/api/media/{img_filename}"
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
async def download_pptx():
    target_pptx = state["styled_pptx"] or state["content_pptx"]
    if not target_pptx or not os.path.exists(target_pptx):
        raise HTTPException(status_code=404, detail="No output PPTX available to download")
    return FileResponse(
        target_pptx,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename="styled_presentation.pptx"
    )

@app.get("/api/report")
async def get_change_report():
    report_path = os.path.join(UPLOAD_DIR, "change_report.html")
    if not os.path.exists(report_path):
        raise HTTPException(status_code=404, detail="No style change report generated yet.")
    return FileResponse(report_path, media_type="text/html")

@app.get("/api/report-data")
async def get_report_data():
    output_path = os.path.join(UPLOAD_DIR, "styled_output.pptx")
    if not state["content_pptx"] or not os.path.exists(output_path):
        return {"ok": False, "changes": []}
    try:
        from report import collect_changes
        changes = collect_changes(state["content_pptx"], output_path)
        return {"ok": True, "changes": changes}
    except Exception as e:
        return {"ok": False, "detail": str(e), "changes": []}

@app.get("/api/figure-diagnostics")
async def get_figure_diagnostics():
    """Return missing (requested but not cropped) and unplaced (cropped but no placeholder) figures."""
    input_path = state.get("content_pptx")
    if not input_path or not os.path.exists(input_path):
        return {"ok": False, "missing": [], "unplaced": []}
    try:
        from report import collect_figure_diagnostics
        extracts_dir = os.path.join(UPLOAD_DIR, "pdf_extracts")
        missing, unplaced = collect_figure_diagnostics(input_path, extracts_dir)
        return {"ok": True, "missing": missing, "unplaced": unplaced}
    except Exception as e:
        return {"ok": False, "detail": str(e), "missing": [], "unplaced": []}

@app.get("/api/accessibility-report")
async def get_accessibility_report():
    output_path = state.get("styled_pptx") or os.path.join(UPLOAD_DIR, "styled_output.pptx")
    if not os.path.exists(output_path):
        return {"ok": False, "issues": []}
    try:
        from accessibility import check_ppt_accessibility
        issues = check_ppt_accessibility(output_path)
        return {"ok": True, "issues": issues}
    except Exception as e:
        return {"ok": False, "detail": str(e), "issues": []}

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
