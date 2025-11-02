
import os
import uuid
from typing import List

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

from utils import pdf_ops_v3 as pdf_ops

BASE = os.path.dirname(__file__)
UPLOAD_DIR = os.path.join(BASE, "uploads")
OUT_DIR = os.path.join(BASE, "output")
FRONTEND_DIR = os.path.abspath(os.path.join(BASE, "..", "frontend"))

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUT_DIR, exist_ok=True)

app = FastAPI(title="UnlimitedPDFSuite v3")

# Serve the frontend
app.mount("/assets", StaticFiles(directory=os.path.join(FRONTEND_DIR, "assets")), name="assets")

@app.get("/", response_class=HTMLResponse)
def index():
    with open(os.path.join(FRONTEND_DIR, "index.html"), "r", encoding="utf-8") as f:
        return HTMLResponse(f.read())

# Serve generated files

@app.get("/output/{path:path}")
def serve_output(path: str):
    # Normalize and ensure the requested path stays within OUT_DIR
    requested = os.path.normpath(os.path.join(OUT_DIR, path))
    if not requested.startswith(os.path.abspath(OUT_DIR) + os.sep) and requested != os.path.abspath(OUT_DIR):
        raise HTTPException(status_code=400, detail="Invalid path")
    if not os.path.exists(requested):
        raise HTTPException(status_code=404, detail="File not found")
    filename = os.path.basename(requested)
    mime = "application/octet-stream"
    if filename.lower().endswith(".pdf"):
        mime = "application/pdf"
    elif filename.lower().endswith(".txt"):
        mime = "text/plain"
    return FileResponse(requested, media_type=mime, filename=filename)


@app.get("/images/{dirname}/{filename}")
def serve_image(dirname: str, filename: str):
    file_path = os.path.join(OUT_DIR, dirname, filename)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File not found")
    if filename.lower().endswith(".png"):
        mime = "image/png"
    elif filename.lower().endswith((".jpg",".jpeg")):
        mime = "image/jpeg"
    else:
        mime = "application/octet-stream"
    return FileResponse(file_path, media_type=mime, filename=filename)

def _save_upload(u: UploadFile) -> str:
    ext = os.path.splitext(u.filename or "")[1].lower()
    name = f"{uuid.uuid4().hex}{ext or ''}"
    path = os.path.join(UPLOAD_DIR, name)
    data = u.file.read()
    with open(path, "wb") as f:
        f.write(data)
    return path

def _outpath(name: str) -> str:
    return os.path.join(OUT_DIR, name)

# ---- API Endpoints ----

@app.post("/api/convert/to-pdf")
def api_convert_to_pdf(files: List[UploadFile] = File(...)):
    in_files = [_save_upload(u) for u in files]
    out_path = _outpath(f"converted_{uuid.uuid4().hex}.pdf")
    doc_files, img_files = [], []
    for p in in_files:
        ext = os.path.splitext(p)[1].lower()
        if ext in [".jpg",".jpeg",".png",".gif",".tiff",".bmp",".webp"]:
            img_files.append(p)
        else:
            doc_files.append(p)
    temp_pdfs = []
    try:
        for p in doc_files:
            ext = os.path.splitext(p)[1].lower()
            tmp = _outpath(f"tmp_{uuid.uuid4().hex}.pdf")
            if ext == ".txt":
                pdf_ops.txt_to_pdf(p, tmp)
            else:
                pdf_ops.office_to_pdf(p, tmp)
            temp_pdfs.append(tmp)
        if img_files:
            img_pdf = _outpath(f"tmp_img_{uuid.uuid4().hex}.pdf")
            pdf_ops.images_to_pdf(img_files, img_pdf)
            temp_pdfs.append(img_pdf)
        if not temp_pdfs:
            raise HTTPException(status_code=400, detail="No convertible files provided.")
        pdf_ops.merge_pdfs(temp_pdfs, out_path)
        return {"file": f"/output/{os.path.basename(out_path)}"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Conversion failed: {e}")
    finally:
        for t in temp_pdfs:
            try: os.remove(t)
            except: pass

@app.post("/api/convert/from-pdf")
def api_convert_from_pdf(file: UploadFile = File(...),
                         to: str = Form(...),
                         dpi: int = Form(200),
                         fmt: str = Form("png")):
    in_path = _save_upload(file)
    try:
        if to == "images":
            out_dir_name = f"images_{uuid.uuid4().hex}"
            out_dir = os.path.join(OUT_DIR, out_dir_name)
            files = pdf_ops.pdf_to_images(in_path, out_dir, dpi=dpi, fmt=fmt)
            served = [f"/images/{out_dir_name}/{os.path.basename(f)}" for f in files]
            return {"files": served}
        elif to == "text":
            out_path = _outpath(f"text_{uuid.uuid4().hex}.txt")
            pdf_ops.extract_text(in_path, out_path)
            return {"file": f"/output/{os.path.basename(out_path)}"}
        else:
            raise HTTPException(status_code=400, detail="Unsupported 'to' value")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Conversion failed: {e}")

@app.post("/api/ocr")
def api_ocr(file: UploadFile = File(...), lang: str = Form("eng")):
    in_path = _save_upload(file)
    out_path = _outpath(f"ocr_{uuid.uuid4().hex}.pdf")
    try:
        pdf_ops.ocr_pdf(in_path, out_path, lang=lang)
        return {"file": f"/output/{os.path.basename(out_path)}"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"OCR failed: {e}")

@app.post("/api/compress")
def api_compress(file: UploadFile = File(...), dpi: int = Form(150), grayscale: bool = Form(False)):
    in_path = _save_upload(file)
    out_path = _outpath(f"compressed_{uuid.uuid4().hex}.pdf")
    try:
        pdf_ops.compress_pdf(in_path, out_path, dpi=dpi, grayscale=grayscale)
        return {"file": f"/output/{os.path.basename(out_path)}"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Compression failed: {e}")

@app.post("/api/merge")
def api_merge(files: List[UploadFile] = File(...)):
    in_paths = [_save_upload(u) for u in files]
    out_path = _outpath(f"merged_{uuid.uuid4().hex}.pdf")
    try:
        pdf_ops.merge_pdfs(in_paths, out_path)
        return {"file": f"/output/{os.path.basename(out_path)}"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Merge failed: {e}")

@app.post("/api/merge/alternating")
def api_merge_alternating(file_a: UploadFile = File(...), file_b: UploadFile = File(...)):
    a = _save_upload(file_a)
    b = _save_upload(file_b)
    out_path = _outpath(f"altmerge_{uuid.uuid4().hex}.pdf")
    try:
        pdf_ops.merge_alternating(a, b, out_path)
        return {"file": f"/output/{os.path.basename(out_path)}"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Alternating merge failed: {e}")

@app.post("/api/split")
def api_split(file: UploadFile = File(...), ranges: str = Form(...)):
    in_path = _save_upload(file)
    out_dir = os.path.join(OUT_DIR, f"split_{uuid.uuid4().hex}")
    try:
        files = pdf_ops.split_pdf(in_path, out_dir, ranges)
        served = [f\"/output/{os.path.relpath(f, OUT_DIR)}\" for f in files]
        return {"files": served}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Split failed: {e}")

@app.post("/api/rotate")
def api_rotate(file: UploadFile = File(...), rotation: int = Form(90)):
    in_path = _save_upload(file)
    out_path = _outpath(f"rotated_{uuid.uuid4().hex}.pdf")
    try:
        pdf_ops.rotate_pdf(in_path, out_path, rotation)
        return {"file": f"/output/{os.path.basename(out_path)}"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Rotate failed: {e}")

@app.post("/api/reorder")
def api_reorder(file: UploadFile = File(...), order: str = Form(...)):
    in_path = _save_upload(file)
    out_path = _outpath(f"reordered_{uuid.uuid4().hex}.pdf")
    try:
        order_list = [int(i.strip()) for i in order.split(",") if i.strip()]
        pdf_ops.reorder_pdf(in_path, out_path, order_list)
        return {"file": f"/output/{os.path.basename(out_path)}"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Reorder failed: {e}")

@app.post("/api/delete-pages")
def api_delete_pages(file: UploadFile = File(...), indices: str = Form(...)):
    in_path = _save_upload(file)
    out_path = _outpath(f"deleted_{uuid.uuid4().hex}.pdf")
    try:
        idx = [int(i.strip()) for i in indices.split(",") if i.strip()]
        pdf_ops.delete_pages(in_path, out_path, idx)
        return {"file": f"/output/{os.path.basename(out_path)}"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Delete pages failed: {e}")

@app.post("/api/layout/nup")
def api_nup(file: UploadFile = File(...), cols: int = Form(2), rows: int = Form(2), margin_mm: float = Form(5.0)):
    in_path = _save_upload(file)
    out_path = _outpath(f"nup_{uuid.uuid4().hex}.pdf")
    try:
        pdf_ops.n_up_pdf(in_path, out_path, cols=cols, rows=rows, margin_mm=margin_mm)
        return {"file": f"/output/{os.path.basename(out_path)}"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"N-up layout failed: {e}")

@app.post("/api/layout/booklet")
def api_booklet(file: UploadFile = File(...), margin_mm: float = Form(5.0)):
    in_path = _save_upload(file)
    out_path = _outpath(f"booklet_{uuid.uuid4().hex}.pdf")
    try:
        pdf_ops.booklet_impose(in_path, out_path, margin_mm=margin_mm)
        return {"file": f"/output/{os.path.basename(out_path)}"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Booklet layout failed: {e}")

@app.post("/api/password/set")
def api_set_password(file: UploadFile = File(...), user_pass: str = Form(""), owner_pass: str = Form(""),
                     allow_printing: bool = Form(True), allow_copy: bool = Form(True), allow_modify: bool = Form(True)):
    in_path = _save_upload(file)
    out_path = _outpath(f"protected_{uuid.uuid4().hex}.pdf")
    try:
        pdf_ops.set_password(in_path, out_path, user_pass, owner_pass or None, allow_printing, allow_copy, allow_modify)
        return {"file": f"/output/{os.path.basename(out_path)}"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Password protection failed: {e}")

@app.post("/api/password/unlock")
def api_unlock_password(file: UploadFile = File(...), password: str = Form(...)):
    in_path = _save_upload(file)
    out_path = _outpath(f"unlocked_{uuid.uuid4().hex}.pdf")
    try:
        pdf_ops.unlock_pdf(in_path, out_path, password)
        return {"file": f"/output/{os.path.basename(out_path)}"}
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Unlock failed: {e}")

@app.post("/api/header-footer")
def api_header_footer(file: UploadFile = File(...), header: str = Form(""), footer: str = Form(""), font_size: int = Form(10)):
    in_path = _save_upload(file)
    out_path = _outpath(f"hf_{uuid.uuid4().hex}.pdf")
    try:
        pdf_ops.add_header_footer(in_path, out_path, header, footer, font_size)
        return {"file": f"/output/{os.path.basename(out_path)}"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Header/Footer failed: {e}")

@app.post("/api/bookmarks/from-filenames")
def api_bookmarks_from_filenames(files: List[UploadFile] = File(...)):
    in_paths = [_save_upload(u) for u in files]
    out_path = _outpath(f"bookmarked_merged_{uuid.uuid4().hex}.pdf")
    try:
        pdf_ops.bookmarks_from_filenames(in_paths, out_path)
        return {"file": f"/output/{os.path.basename(out_path)}"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Bookmarks/Merge failed: {e}")


@app.post("/api/convert/pdf-to-docx")
def api_pdf_to_docx(file: UploadFile = File(...),
                    ocr: bool = Form(False),
                    ocr_lang: str = Form(""),
                    mode: str = Form("auto")):
    """
    Convert a single PDF to DOCX using pdf2docx.
    Optional OCR pre-pass (ocrmypdf) to improve text extraction.
    """
    in_path = _save_upload(file)
    docx_name = f"docx_{uuid.uuid4().hex}.docx"
    out_path = _outpath(docx_name)
    try:
        pdf_ops.pdf_to_docx(in_path, out_path, mode=mode, ocr_lang=(ocr_lang if ocr else ""))
        return {"file": f"/output/{os.path.basename(out_path)}"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PDF→DOCX failed: {e}")
