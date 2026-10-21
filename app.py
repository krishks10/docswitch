"""
DocSwitch - a small PDF / PPT converter toolkit built with Python + Gradio.

Tools:
  1. PDF  -> PPTX   (image slides, or editable-text slides)
  2. PPTX -> PDF    (LibreOffice headless)
  3. Images -> PDF  (handy for scanned lab reports / handwritten assignments)
  4. Merge PDFs
  5. Extract / split pages from a PDF
  6. PDF -> Text
"""

import io
import os
import shutil
import subprocess
import tempfile
import uuid

import gradio as gr
import pymupdf
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Emu, Pt
from pypdf import PdfReader, PdfWriter

MAX_PAGES = 120          # keep free hosting safe
MAX_MB = 50              # per file
OUT_ROOT = tempfile.mkdtemp(prefix="docswitch_")


# ----------------------------------------------------------------- helpers
def _new_out(name: str) -> str:
    """Return a unique output path inside our temp root."""
    d = os.path.join(OUT_ROOT, uuid.uuid4().hex[:8])
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, name)


def _path(f) -> str:
    return f if isinstance(f, str) else f.name


def _check_size(path: str):
    mb = os.path.getsize(path) / (1024 * 1024)
    if mb > MAX_MB:
        raise gr.Error(f"'{os.path.basename(path)}' is {mb:.1f} MB. Limit is {MAX_MB} MB.")


def _base(path: str) -> str:
    return os.path.splitext(os.path.basename(path))[0]


def _parse_ranges(spec: str, total: int):
    """'1-3, 5, 8-' -> [0,1,2,4,7,...] (0-based, order kept, no duplicates)."""
    pages, seen = [], set()
    for part in spec.replace(" ", "").split(","):
        if not part:
            continue
        try:
            if "-" in part:
                a, b = part.split("-", 1)
                start = int(a) if a else 1
                end = int(b) if b else total
            else:
                start = end = int(part)
        except ValueError:
            raise gr.Error(f"Could not understand '{part}'. Use something like 1-3, 5, 8-")
        if start < 1 or end > total or start > end:
            raise gr.Error(f"Range '{part}' is outside 1-{total}.")
        for p in range(start - 1, end):
            if p not in seen:
                seen.add(p)
                pages.append(p)
    if not pages:
        raise gr.Error("No pages selected.")
    return pages


# ------------------------------------------------------------- PDF -> PPTX
def pdf_to_pptx(file, mode, quality, progress=gr.Progress()):
    if file is None:
        raise gr.Error("Upload a PDF first.")
    src = _path(file)
    _check_size(src)

    doc = pymupdf.open(src)
    if doc.page_count > MAX_PAGES:
        raise gr.Error(f"PDF has {doc.page_count} pages. Limit is {MAX_PAGES}.")

    first = doc[0].rect
    # Slide size = PDF page size (so nothing gets stretched)
    slide_w = Emu(int(first.width * 12700))    # 1 pt = 12700 EMU
    slide_h = Emu(int(first.height * 12700))
    prs = Presentation()
    prs.slide_width, prs.slide_height = slide_w, slide_h
    blank = prs.slide_layouts[6]

    zoom = {"Fast (150 dpi)": 150 / 72, "Standard (200 dpi)": 200 / 72, "Sharp (300 dpi)": 300 / 72}[quality]

    for i, page in enumerate(doc):
        progress((i + 1) / doc.page_count, desc=f"Page {i + 1}/{doc.page_count}")
        slide = prs.slides.add_slide(blank)
        sx = slide_w / page.rect.width
        sy = slide_h / page.rect.height

        if mode == "Image slides (looks exactly like the PDF)":
            pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
            buf = io.BytesIO(pix.tobytes("jpeg", jpg_quality=88))
            slide.shapes.add_picture(buf, 0, 0, width=slide_w, height=slide_h)
        else:
            # Editable text: background = page images only, text = real text boxes
            blocks = page.get_text("dict")["blocks"]
            for b in blocks:
                if b["type"] == 1:  # image block
                    try:
                        x0, y0, x1, y1 = b["bbox"]
                        slide.shapes.add_picture(
                            io.BytesIO(b["image"]),
                            Emu(int(x0 * sx)), Emu(int(y0 * sy)),
                            Emu(int((x1 - x0) * sx)), Emu(int((y1 - y0) * sy)),
                        )
                    except Exception:
                        pass  # unsupported image format, skip it
            for b in blocks:
                if b["type"] != 0:
                    continue
                x0, y0, x1, y1 = b["bbox"]
                tb = slide.shapes.add_textbox(
                    Emu(int(x0 * sx)), Emu(int(y0 * sy)),
                    Emu(int((x1 - x0) * sx * 1.05)), Emu(int((y1 - y0) * sy)),
                )
                tf = tb.text_frame
                tf.word_wrap = True
                tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
                first_para = True
                for line in b["lines"]:
                    para = tf.paragraphs[0] if first_para else tf.add_paragraph()
                    first_para = False
                    for sp in line["spans"]:
                        if not sp["text"].strip() and not sp["text"]:
                            continue
                        run = para.add_run()
                        run.text = sp["text"]
                        run.font.size = Pt(max(sp["size"], 4))
                        run.font.bold = bool(sp["flags"] & 16)
                        run.font.italic = bool(sp["flags"] & 2)
                        c = sp["color"]
                        run.font.color.rgb = RGBColor((c >> 16) & 255, (c >> 8) & 255, c & 255)

    out = _new_out(_base(src) + ".pptx")
    prs.save(out)
    doc.close()
    return out


# ------------------------------------------------------------- PPTX -> PDF
def pptx_to_pdf(file, progress=gr.Progress()):
    if file is None:
        raise gr.Error("Upload a .pptx / .ppt file first.")
    src = _path(file)
    _check_size(src)
    if not shutil.which("soffice"):
        raise gr.Error("LibreOffice is not installed on this server.")

    progress(0.2, desc="Converting with LibreOffice...")
    work = tempfile.mkdtemp(prefix="lo_")
    profile = tempfile.mkdtemp(prefix="lo_profile_")
    ext = os.path.splitext(src)[1] or ".pptx"
    tmp_in = os.path.join(work, "input" + ext)
    shutil.copy(src, tmp_in)
    try:
        subprocess.run(
            [
                "soffice", f"-env:UserInstallation=file://{profile}",
                "--headless", "--convert-to", "pdf", "--outdir", work, tmp_in,
            ],
            check=True, capture_output=True, timeout=180,
        )
    except subprocess.TimeoutExpired:
        raise gr.Error("Conversion took too long. Try a smaller file.")
    except subprocess.CalledProcessError as e:
        raise gr.Error("LibreOffice failed: " + e.stderr.decode(errors="ignore")[:300])
    finally:
        shutil.rmtree(profile, ignore_errors=True)

    produced = os.path.join(work, "input.pdf")
    if not os.path.exists(produced):
        raise gr.Error("Conversion produced no output. Is the file a valid PowerPoint?")
    out = _new_out(_base(src) + ".pdf")
    shutil.move(produced, out)
    shutil.rmtree(work, ignore_errors=True)
    return out


# ------------------------------------------------------------ Images -> PDF
def images_to_pdf(files, page_size, progress=gr.Progress()):
    if not files:
        raise gr.Error("Upload at least one image.")
    if len(files) > MAX_PAGES:
        raise gr.Error(f"Too many images. Limit is {MAX_PAGES}.")

    A4 = (595, 842)  # points
    doc = pymupdf.open()
    for i, f in enumerate(files):
        progress((i + 1) / len(files), desc=f"Image {i + 1}/{len(files)}")
        p = _path(f)
        _check_size(p)
        with Image.open(p) as im:
            im = im.convert("RGB")
            w, h = im.size
            buf = io.BytesIO()
            im.save(buf, "JPEG", quality=90)
        if page_size == "Fit to image":
            page = doc.new_page(width=w * 0.75, height=h * 0.75)
            rect = page.rect
        else:
            landscape = w > h
            pw, ph = (A4[1], A4[0]) if landscape else A4
            page = doc.new_page(width=pw, height=ph)
            margin = 20
            scale = min((pw - 2 * margin) / w, (ph - 2 * margin) / h)
            iw, ih = w * scale, h * scale
            x0, y0 = (pw - iw) / 2, (ph - ih) / 2
            rect = pymupdf.Rect(x0, y0, x0 + iw, y0 + ih)
        page.insert_image(rect, stream=buf.getvalue())

    out = _new_out("images.pdf")
    doc.save(out, deflate=True)
    doc.close()
    return out


# ---------------------------------------------------------------- Merge PDFs
def merge_pdfs(files):
    if not files or len(files) < 2:
        raise gr.Error("Upload at least two PDFs to merge.")
    writer = PdfWriter()
    total = 0
    for f in files:
        p = _path(f)
        _check_size(p)
        reader = PdfReader(p)
        total += len(reader.pages)
        if total > MAX_PAGES * 3:
            raise gr.Error("Combined page count too large.")
        for page in reader.pages:
            writer.add_page(page)
    out = _new_out("merged.pdf")
    with open(out, "wb") as fh:
        writer.write(fh)
    return out


# ------------------------------------------------------ Extract / split pages
def extract_pages(file, spec):
    if file is None:
        raise gr.Error("Upload a PDF first.")
    src = _path(file)
    _check_size(src)
    reader = PdfReader(src)
    pages = _parse_ranges(spec, len(reader.pages))
    writer = PdfWriter()
    for p in pages:
        writer.add_page(reader.pages[p])
    out = _new_out(_base(src) + "_selected.pdf")
    with open(out, "wb") as fh:
        writer.write(fh)
    return out


# ---------------------------------------------------------------- PDF -> Text
def pdf_to_text(file):
    if file is None:
        raise gr.Error("Upload a PDF first.")
    src = _path(file)
    _check_size(src)
    doc = pymupdf.open(src)
    if doc.page_count > MAX_PAGES * 3:
        raise gr.Error("PDF is too long.")
    parts = []
    for i, page in enumerate(doc, 1):
        parts.append(f"--- Page {i} ---\n{page.get_text().strip()}\n")
    text = "\n".join(parts)
    doc.close()
    if len(text.replace("-", "").replace("Page", "").strip()) < 20:
        text += "\n(No selectable text found. This looks like a scanned PDF and would need OCR.)"
    out = _new_out(_base(src) + ".txt")
    with open(out, "w", encoding="utf-8") as fh:
        fh.write(text)
    return text, out


# ----------------------------------------------------------------------- UI
INTRO = """
# DocSwitch
Free PDF and PowerPoint tools. Files are processed on the server and deleted when the session ends.
Made with Python, PyMuPDF, python-pptx and Gradio.
"""

with gr.Blocks(title="DocSwitch - PDF & PPT converter") as demo:
    gr.Markdown(INTRO)

    with gr.Tab("PDF → PPT"):
        gr.Markdown("Turn a PDF into a PowerPoint. **Image slides** keep the exact look. **Editable text** rebuilds text boxes so you can edit them (layout may shift a little).")
        a_in = gr.File(label="PDF file", file_types=[".pdf"])
        with gr.Row():
            a_mode = gr.Radio(
                ["Image slides (looks exactly like the PDF)", "Editable text (best for text-heavy PDFs)"],
                value="Image slides (looks exactly like the PDF)", label="Mode",
            )
            a_q = gr.Dropdown(["Fast (150 dpi)", "Standard (200 dpi)", "Sharp (300 dpi)"],
                              value="Standard (200 dpi)", label="Image quality")
        a_btn = gr.Button("Convert to PPTX", variant="primary")
        a_out = gr.File(label="Download PPTX")
        a_btn.click(pdf_to_pptx, [a_in, a_mode, a_q], a_out)

    with gr.Tab("PPT → PDF"):
        gr.Markdown("Convert .pptx / .ppt slides to a PDF using LibreOffice.")
        b_in = gr.File(label="PowerPoint file", file_types=[".pptx", ".ppt"])
        b_btn = gr.Button("Convert to PDF", variant="primary")
        b_out = gr.File(label="Download PDF")
        b_btn.click(pptx_to_pdf, b_in, b_out)

    with gr.Tab("Images → PDF"):
        gr.Markdown("Photos of handwritten notes, lab reports or scanned pages into one PDF. Order = upload order.")
        c_in = gr.File(label="Images (JPG / PNG)", file_count="multiple", file_types=["image"])
        c_size = gr.Radio(["A4 page", "Fit to image"], value="A4 page", label="Page size")
        c_btn = gr.Button("Create PDF", variant="primary")
        c_out = gr.File(label="Download PDF")
        c_btn.click(images_to_pdf, [c_in, c_size], c_out)

    with gr.Tab("Merge PDFs"):
        gr.Markdown("Combine several PDFs into one, in the order you upload them.")
        d_in = gr.File(label="PDF files", file_count="multiple", file_types=[".pdf"])
        d_btn = gr.Button("Merge", variant="primary")
        d_out = gr.File(label="Download merged PDF")
        d_btn.click(merge_pdfs, d_in, d_out)

    with gr.Tab("Extract pages"):
        gr.Markdown("Keep only the pages you need. Example: `1-3, 5, 8-` (8 to the end).")
        e_in = gr.File(label="PDF file", file_types=[".pdf"])
        e_spec = gr.Textbox(label="Pages", placeholder="1-3, 5, 8-")
        e_btn = gr.Button("Extract", variant="primary")
        e_out = gr.File(label="Download PDF")
        e_btn.click(extract_pages, [e_in, e_spec], e_out)

    with gr.Tab("PDF → Text"):
        gr.Markdown("Pull out selectable text. Scanned PDFs need OCR, which this tool does not do.")
        f_in = gr.File(label="PDF file", file_types=[".pdf"])
        f_btn = gr.Button("Extract text", variant="primary")
        f_text = gr.Textbox(label="Extracted text", lines=12, buttons=["copy"])
        f_out = gr.File(label="Download .txt")
        f_btn.click(pdf_to_text, f_in, [f_text, f_out])

demo.queue(max_size=20)

if __name__ == "__main__":
    demo.launch(server_name="0.0.0.0", server_port=int(os.environ.get("PORT", 7860)))
