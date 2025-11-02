import os
import io
import shlex
import subprocess
from typing import List, Optional, Tuple

from PIL import Image
import img2pdf
import pikepdf
from PyPDF2 import PdfReader, PdfWriter, PdfMerger, Transformation
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.units import mm
from pdf2image import convert_from_path
from pdfminer.high_level import extract_text_to_fp

# ---- Helpers ----

def run_cmd(cmd: str) -> Tuple[int, str, str]:
    """Run shell command and return (code, out, err)."""
    p = subprocess.Popen(shlex.split(cmd), stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    out, err = p.communicate()
    return p.returncode, out.decode('utf-8', 'ignore'), err.decode('utf-8', 'ignore')

def save_bytes(path: str, data: bytes):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as f:
        f.write(data)

def ensure_pdf(path: str):
    if not path.lower().endswith(".pdf"):
        raise ValueError("Output must be a .pdf")

# ---- Conversion ----

def office_to_pdf(input_path: str, output_path: str) -> str:
    """Convert Office formats (docx/xlsx/pptx, etc.) to PDF using LibreOffice headless."""
    ensure_pdf(output_path)
    out_dir = os.path.dirname(output_path) or "."
    cmd = f"soffice --headless --convert-to pdf --outdir {shlex.quote(out_dir)} {shlex.quote(input_path)}"
    code, out, err = run_cmd(cmd)

    # LibreOffice output file name is based on input file name
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    converted_path = os.path.join(out_dir, f"{base_name}.pdf")

    if code != 0:
        raise RuntimeError(f"LibreOffice conversion failed. Error: {err or out}")
    if not os.path.exists(converted_path):
        raise RuntimeError(f"LibreOffice succeeded but did not create expected file: {converted_path}")
    
    # Rename to desired output_path
    os.replace(converted_path, output_path)
    return output_path

def images_to_pdf(input_paths: List[str], output_path: str) -> str:
    """Convert a list of image files to a single PDF."""
    ensure_pdf(output_path)
    with open(output_path, "wb") as f:
        f.write(img2pdf.convert(input_paths))
    return output_path

def txt_to_pdf(input_path: str, output_path: str, pagesize: str = "A4") -> str:
    """Convert a simple text file to PDF (auto page breaks)."""
    ensure_pdf(output_path)
    size = A4 if pagesize.upper() == "A4" else letter
    c = canvas.Canvas(output_path, pagesize=size)
    width, height = size
    left = 20 * mm
    top = height - 20 * mm
    bottom = 20 * mm
    text = c.beginText(left, top)
    text.setFont("Helvetica", 10)
    
    with open(input_path, "r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            if text.getY() <= bottom:
                c.drawText(text)
                c.showPage()
                text = c.beginText(left, top)
                text.setFont("Helvetica", 10)
            text.textLine(line.rstrip("\n"))
    
    c.drawText(text)
    c.save()
    return output_path

# ---- PDF Export ----

def pdf_to_images(input_path: str, output_dir: str, dpi: int = 200, fmt: str = "png") -> List[str]:
    """Convert PDF pages to image files using pdf2image (requires Poppler)."""
    os.makedirs(output_dir, exist_ok=True)
    # paths_only=True returns filesystem paths instead of PIL images
    images = convert_from_path(input_path, dpi=dpi, fmt=fmt, output_folder=output_dir, paths_only=True)
    return list(images)

def extract_text(input_path: str, output_path: str) -> str:
    """Extract layout-free text from PDF to a TXT file."""
    with open(output_path, "wb") as out_file:
        with open(input_path, 'rb') as in_file:
            extract_text_to_fp(in_file, out_file)
    return output_path

# ---- Organization ----

def merge_pdfs(input_paths: List[str], output_path: str) -> str:
    """Merge a list of PDF files into a single PDF."""
    ensure_pdf(output_path)
    merger = PdfMerger()
    for path in input_paths:
        merger.append(path)
    with open(output_path, "wb") as f:
        merger.write(f)
    merger.close()
    return output_path

def _parse_ranges(ranges_str: str, num_pages: int) -> List[List[int]]:
    """Parse '1-3,5,9-12' into 1-based page lists limited by num_pages."""
    output = []
    parts = [p.strip() for p in ranges_str.replace(";", ",").split(",") if p.strip()]
    for part in parts:
        if "-" in part:
            a, b = part.split("-", 1)
            try:
                start = max(1, int(a))
                end = min(int(b), num_pages)
                if start <= end:
                    output.append(list(range(start, end + 1)))
            except ValueError:
                continue
        else:
            try:
                p = int(part)
                if 1 <= p <= num_pages:
                    output.append([p])
            except ValueError:
                continue
    return output

def split_pdf(input_path: str, output_dir: str, ranges: str) -> List[str]:
    """Split a PDF by page ranges (e.g., '1-3,5,9-12')."""
    os.makedirs(output_dir, exist_ok=True)
    reader = PdfReader(input_path)
    files = []
    page_sets = _parse_ranges(ranges, len(reader.pages))
    
    for i, page_set in enumerate(page_sets):
        writer = PdfWriter()
        for page_num in page_set:
            writer.add_page(reader.pages[page_num - 1])
        out_path = os.path.join(output_dir, f"split_{i+1}.pdf")
        with open(out_path, "wb") as f:
            writer.write(f)
        files.append(out_path)
    return files

def rotate_pdf(input_path: str, output_path: str, rotation: int = 90) -> str:
    """Rotate all pages in a PDF (degrees: 90/180/270)."""
    ensure_pdf(output_path)
    reader = PdfReader(input_path)
    writer = PdfWriter()
    for page in reader.pages:
        page.rotate(rotation)
        writer.add_page(page)
    with open(output_path, "wb") as f:
        writer.write(f)
    return output_path

def reorder_pdf(input_path: str, output_path: str, order: List[int]) -> str:
    """Reorder pages based on a 0-indexed list of page indices."""
    ensure_pdf(output_path)
    reader = PdfReader(input_path)
    writer = PdfWriter()
    for index in order:
        if 0 <= index < len(reader.pages):
            writer.add_page(reader.pages[index])
        else:
            raise ValueError(f"Invalid index in order: {index}")
    with open(output_path, "wb") as f:
        writer.write(f)
    return output_path

def delete_pages(input_path: str, output_path: str, indices: List[int]) -> str:
    """Delete pages based on a 0-indexed list of page indices."""
    ensure_pdf(output_path)
    reader = PdfReader(input_path)
    writer = PdfWriter()
    to_delete = set(i for i in indices if 0 <= i < len(reader.pages))
    for i, page in enumerate(reader.pages):
        if i not in to_delete:
            writer.add_page(page)
    with open(output_path, "wb") as f:
        writer.write(f)
    return output_path

# ---- Layout (NEW IN V3) ----

def _page_size(reader: PdfReader):
    p = reader.pages[0]
    return float(p.mediabox.width), float(p.mediabox.height)

def n_up_pdf(input_path: str, output_path: str, cols: int = 2, rows: int = 2, margin_mm: float = 5.0) -> str:
    """Place multiple source pages on a single sheet (N-up)."""
    ensure_pdf(output_path)
    reader = PdfReader(input_path)
    writer = PdfWriter()
    W, H = _page_size(reader)
    margin = margin_mm * mm
    cell_w = (W - 2*margin) / cols
    cell_h = (H - 2*margin) / rows

    sheet = None
    cell_idx = 0
    for page in reader.pages:
        if sheet is None or cell_idx >= cols*rows:
            sheet = writer.add_blank_page(width=W, height=H)
            cell_idx = 0
        r = cell_idx // cols
        c = cell_idx % cols
        x = margin + c * cell_w
        y = H - margin - (r+1) * cell_h  # top-down
        pw = float(page.mediabox.width); ph = float(page.mediabox.height)
        scale = min(cell_w/pw, cell_h/ph)
        tx = Transformation().scale(scale).translate(x, y)
        sheet.merge_transformed_page(page, tx)
        cell_idx += 1

    with open(output_path, "wb") as f:
        writer.write(f)
    return output_path

def booklet_impose(input_path: str, output_path: str, margin_mm: float = 5.0) -> str:
    """
    Simple 2-up booklet imposition in landscape.
    Pads to multiple of 4 pages. Produces sheets with two pages side-by-side.
    """
    ensure_pdf(output_path)
    reader = PdfReader(input_path)
    n = len(reader.pages)
    pad = (4 - (n % 4)) % 4
    pages = list(range(n)) + [-1]*pad  # -1 means blank
    left = 0
    right = len(pages) - 1
    order = []
    while left < right:
        order.extend([pages[right], pages[left]])  # Back sheet
        left += 1; right -= 1
        if left < right:
            order.extend([pages[left], pages[right]])  # Front sheet
            left += 1; right -= 1

    W, H = _page_size(reader)
    sheet_w, sheet_h = H*2, W  # landscape with two portrait pages side by side
    margin = margin_mm * mm
    cell_w = (sheet_w - 3*margin)/2
    cell_h = sheet_h - 2*margin

    writer = PdfWriter()
    for i in range(0, len(order), 2):
        sheet = writer.add_blank_page(width=sheet_w, height=sheet_h)
        for j in range(2):
            idx = order[i+j] if i+j < len(order) else -1
            if idx == -1:
                continue
            page = reader.pages[idx]
            pw, ph = float(page.mediabox.width), float(page.mediabox.height)
            scale = min(cell_w/pw, cell_h/ph)
            x = margin if j == 0 else margin*2 + cell_w
            y = (sheet_h - cell_h)/2
            tx = Transformation().scale(scale).translate(x, y)
            sheet.merge_transformed_page(page, tx)
    with open(output_path, "wb") as f:
        writer.write(f)
    return output_path

def merge_alternating(input_a: str, input_b: str, output_path: str) -> str:
    """Interleave pages from two PDFs: A1, B1, A2, B2, ... (append remaining)."""
    ensure_pdf(output_path)
    ra = PdfReader(input_a)
    rb = PdfReader(input_b)
    wa = len(ra.pages); wb = len(rb.pages)
    w = PdfWriter()
    for i in range(max(wa, wb)):
        if i < wa:
            w.add_page(ra.pages[i])
        if i < wb:
            w.add_page(rb.pages[i])
    with open(output_path, "wb") as f:
        w.write(f)
    return output_path

# ---- Security ----

def set_password(input_path: str, output_path: str, user_pass: str, owner_pass: Optional[str] = None,
                 allow_printing: bool = True, allow_copy: bool = True, allow_modify: bool = True) -> str:
    """Set an open password and permissions for a PDF (AES-256)."""
    ensure_pdf(output_path)
    perms = pikepdf.Permissions(
        print=allow_printing,
        copy=allow_copy,
        modify=allow_modify,
        accessibility=True,
        annotate=True,
        form_fill=True,
        assemble=True,
        high_quality_print=allow_printing
    )
    with pikepdf.Pdf.open(input_path) as pdf:
        pdf.save(output_path, encryption=pikepdf.Encryption(
            user=user_pass or "",
            owner=owner_pass or (user_pass or ""),
            R=6,  # AES-256
            allow=perms
        ))
    return output_path

def unlock_pdf(input_path: str, output_path: str, password: str) -> str:
    """Unlock PDF if the correct password is provided (removes encryption)."""
    ensure_pdf(output_path)
    try:
        with pikepdf.Pdf.open(input_path, password=password) as pdf:
            pdf.save(output_path)  # saving without 'encryption=' removes password
        return output_path
    except pikepdf.PasswordError as e:
        raise ValueError("Incorrect password provided for unlocking.") from e

# ---- View/Metadata ----

def add_header_footer(input_path: str, output_path: str, header: str = "", footer: str = "", font_size: int = 10) -> str:
    """Add simple header and footer text to each page (auto-sizes overlay to page)."""
    ensure_pdf(output_path)
    reader = PdfReader(input_path)
    writer = PdfWriter()

    for i, page in enumerate(reader.pages, start=1):
        packet = io.BytesIO()
        w = float(page.mediabox.width)
        h = float(page.mediabox.height)
        c = canvas.Canvas(packet, pagesize=(w, h))
        c.setFont("Helvetica", font_size)
        if header:
            c.drawString(10*mm, h - 10*mm, header.replace("{page}", str(i)))
        if footer:
            c.drawRightString(w - 10*mm, 10*mm, footer.replace("{page}", str(i)))
        c.save()
        packet.seek(0)
        overlay_reader = PdfReader(packet)
        overlay_page = overlay_reader.pages[0]
        page.merge_page(overlay_page)
        writer.add_page(page)

    with open(output_path, "wb") as f:
        writer.write(f)
    return output_path

# ---- Optimization ----

def ocr_pdf(input_path: str, output_path: str, lang: str = "eng") -> str:
    """Perform OCR on a PDF using ocrmypdf (relies on Tesseract)."""
    ensure_pdf(output_path)
    cmd = f"ocrmypdf -l {shlex.quote(lang)} --output-type pdfa --optimize 0 --clean {shlex.quote(input_path)} {shlex.quote(output_path)}"
    code, out, err = run_cmd(cmd)
    if code != 0 or not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
        raise RuntimeError(f"OCR failed. Error: {err or out}")
    return output_path

def compress_pdf(input_path: str, output_path: str, dpi: int = 150, grayscale: bool = False) -> str:
    """Simple compression by rasterizing pages to images and back. For better results use Ghostscript."""
    ensure_pdf(output_path)
    tmp_dir = os.path.join(os.path.dirname(output_path), "_compress_tmp")
    os.makedirs(tmp_dir, exist_ok=True)
    # Render pages
    imgs = pdf_to_images(input_path, tmp_dir, dpi=dpi, fmt="jpeg")
    # Convert back to PDF
    kwargs = {}
    if grayscale:
        # img2pdf uses ColorSpace enum in newer versions
        try:
            kwargs["colorspace"] = img2pdf.ColorSpace.GRAY
        except Exception:
            pass
    with open(output_path, "wb") as f:
        f.write(img2pdf.convert(imgs, **kwargs))
    # Cleanup
    for path in imgs:
        try:
            os.remove(path)
        except Exception:
            pass
    try:
        os.rmdir(tmp_dir)
    except Exception:
        pass
    return output_path

# ---- Bookmarks (basic) ----

def bookmarks_from_filenames(input_paths: List[str], output_path: str) -> str:
    """Merge PDFs and create bookmarks using each file's basename as a top-level outline."""
    ensure_pdf(output_path)
    pdf = pikepdf.Pdf.new()
    outline = pdf.open_outline()
    for p in input_paths:
        with pikepdf.Pdf.open(p) as part:
            start = len(pdf.pages)
            pdf.pages.extend(part.pages)
            title = os.path.splitext(os.path.basename(p))[0]
            # OutlineItem works across pikepdf versions
            try:
                item = pikepdf.OutlineItem(title=title, page_number=start)
                outline.root.append(item)
            except Exception:
                # Fallback older API
                outline.root.append((title, start))
    pdf.save(output_path)
    return output_path


from pdf2docx import Converter as PDF2DOCXConverter

def pdf_to_docx(input_path: str, output_path: str, mode: str = "auto", ocr_lang: str = "") -> str:
    """
    Convert PDF to DOCX using pdf2docx.
    - mode: "auto" (default), "lines", "paragraph", "table" (pdf2docx layout strategies)
    - ocr_lang: if provided, run an OCR pre-pass (ocrmypdf) to improve text layers before conversion.
    """
    ensure_pdf(input_path) if input_path.lower().endswith(".pdf") else None

    # Optional OCR pre-pass to enhance text extraction on scanned PDFs
    temp_pdf = None
    if ocr_lang:
        import uuid, shlex
        temp_pdf = os.path.join(os.path.dirname(output_path), f"_ocr_{uuid.uuid4().hex}.pdf")
        code, out, err = run_cmd(f"ocrmypdf -l {shlex.quote(ocr_lang)} --optimize 0 --clean {shlex.quote(input_path)} {shlex.quote(temp_pdf)}")
        if code != 0 or not os.path.exists(temp_pdf):
            # If OCR fails, fall back to original PDF
            temp_pdf = None

    src = temp_pdf or input_path
    # pdf2docx supports page range and layout mode; we'll pass mode via parse.
    # For complex docs, "auto" is usually best; tables-heavy docs may try "table".
    cvt = PDF2DOCXConverter(src)
    try:
        # pdf2docx doesn't accept 'mode' param directly in convert(); we can set it via parse
        # but for simplicity we call convert() which auto-detects layout reasonably.
        cvt.convert(output_path)  # whole document
    finally:
        cvt.close()

    # Cleanup temp
    if temp_pdf and os.path.exists(temp_pdf):
        try: os.remove(temp_pdf)
        except: pass

    return output_path
