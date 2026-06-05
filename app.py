from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field
from typing import List, Optional
from pathlib import Path
from copy import deepcopy
from io import BytesIO
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.shared import Inches, Pt
from docx.oxml.ns import qn
from PIL import Image, ImageOps
import base64
import uuid
import re
import requests
import urllib.parse

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "template.docx"
OUT_DIR = BASE_DIR / "outputs"
OUT_DIR.mkdir(exist_ok=True)

app = FastAPI(title="Receipt DOCX Generator", version="1.0.6")


class ReceiptItem(BaseModel):
    use_date: str = Field(..., description="YYYY-MM-DD")
    store: str = Field(..., description="영수증 상호명")
    amount: str = Field(..., description="예: 52,000원")
    receipt_image_base64: Optional[str] = Field(None, description="영수증 원본 이미지 base64 또는 data URL")
    receipt_image_url: Optional[str] = Field(None, description="영수증 원본 이미지 URL")


class ReceiptRequest(BaseModel):
    receipts: List[ReceiptItem]
    card_name: str = "중앙청년지원센터 법인카드(0000)"
    user_name: str = "중앙청년지원센터 매니저"
    purpose: str = ""
    output_filename: Optional[str] = "법인카드_영수증_증빙자료.docx"


def clear_cell(cell):
    tc = cell._tc
    for child in list(tc):
        if child.tag != qn("w:tcPr"):
            tc.remove(child)


def style_run(run, size_pt=13, bold=False):
    run.font.name = "맑은 고딕"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "맑은 고딕")
    run.font.size = Pt(size_pt)
    run.font.bold = bold


def write_center(cell, text, size_pt=13, bold=False):
    clear_cell(cell)
    p = cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text or "")
    style_run(run, size_pt=size_pt, bold=bold)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def write_amount_cell(cell, amount):
    clear_cell(cell)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    p1 = cell.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p1.add_run(amount or "")
    style_run(r1, size_pt=13)

    p2 = cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run("* 비목: 인증연구 > 업무추진비 > 회의비")
    style_run(r2, size_pt=11)


def image_from_base64(image_base64: str) -> Image.Image:
    data = image_base64.strip()
    if "," in data and data.lower().startswith("data:image"):
        data = data.split(",", 1)[1]
    data = re.sub(r"\s+", "", data)
    missing_padding = len(data) % 4
    if missing_padding:
        data += "=" * (4 - missing_padding)
    raw = base64.b64decode(data, validate=False)
    img = Image.open(BytesIO(raw))
    img = ImageOps.exif_transpose(img)
    img.load()
    return img


def image_from_url(url: str) -> Image.Image:
    resp = requests.get(url, timeout=20)
    resp.raise_for_status()
    img = Image.open(BytesIO(resp.content))
    img = ImageOps.exif_transpose(img)
    img.load()
    return img


def save_receipt_image(receipt: ReceiptItem, output_path: Path):
    img = None

    if receipt.receipt_image_base64 and len(receipt.receipt_image_base64.strip()) > 100:
        try:
            img = image_from_base64(receipt.receipt_image_base64)
        except Exception:
            img = None

    if img is None and receipt.receipt_image_url:
        try:
            img = image_from_url(receipt.receipt_image_url)
        except Exception:
            img = None

    if img is None:
        return None

    if img.mode not in ("RGB", "RGBA"):
        img = img.convert("RGB")
    img.save(output_path, format="PNG")
    return output_path


def insert_receipt_area(doc, img_path: Optional[Path]):
    table = doc.tables[0]
    receipt_cell = table.cell(7, 0)
    clear_cell(receipt_cell)
    receipt_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    p = receipt_cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if img_path is None:
        run = p.add_run("")
        style_run(run, size_pt=13)
        return

    with Image.open(img_path) as im:
        w_px, h_px = im.size

    max_w_in = 5.7
    max_h_in = 3.35
    aspect = w_px / h_px if h_px else 1
    width_in = min(max_w_in, max_h_in * aspect)
    height_in = width_in / aspect

    if height_in > max_h_in:
        height_in = max_h_in
        width_in = height_in * aspect

    run = p.add_run()
    run.add_picture(str(img_path), width=Inches(width_in))


def fill_template(receipt: ReceiptItem, card_name: str, user_name: str, purpose: str, img_path: Optional[Path]) -> Document:
    doc = Document(str(TEMPLATE_PATH))
    table = doc.tables[0]

    write_center(table.cell(1, 1), card_name, size_pt=13)
    write_center(table.cell(1, 3), receipt.use_date, size_pt=13)
    write_center(table.cell(2, 1), user_name, size_pt=13)
    write_center(table.cell(3, 1), purpose or "", size_pt=13)
    write_center(table.cell(4, 1), receipt.store, size_pt=13)
    write_amount_cell(table.cell(5, 1), receipt.amount)

    for row in table.rows:
        for cell in row.cells:
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    insert_receipt_area(doc, img_path)
    return doc


def append_doc_body(target_doc: Document, source_doc: Document):
    target_doc.add_page_break()
    body = target_doc.element.body
    for child in source_doc.element.body:
        if child.tag.endswith("sectPr"):
            continue
        body.append(deepcopy(child))


@app.get("/")
def root():
    return {"status": "ok", "service": "Receipt DOCX Generator", "version": "1.0.6"}


@app.get("/health")
def health():
    return {"status": "ok"}


@app.get("/download/{folder_id}/{filename}")
def download_file(folder_id: str, filename: str):
    safe_folder = Path(folder_id).name
    safe_filename = Path(filename).name
    file_path = OUT_DIR / safe_folder / safe_filename

    if not file_path.exists():
        raise HTTPException(status_code=404, detail="파일을 찾을 수 없습니다.")

    return FileResponse(
        file_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=safe_filename,
    )


@app.post("/create-receipt-docx")
def create_receipt_docx(req: ReceiptRequest, request: Request):
    if not TEMPLATE_PATH.exists():
        raise HTTPException(status_code=500, detail="template.docx 파일이 서버에 없습니다.")
    if not req.receipts:
        raise HTTPException(status_code=400, detail="영수증 데이터가 없습니다.")

    folder_id = str(uuid.uuid4())
    tmp_dir = OUT_DIR / folder_id
    tmp_dir.mkdir(parents=True, exist_ok=True)

    pages = []
    for i, receipt in enumerate(req.receipts, start=1):
        img_path = tmp_dir / f"receipt_{i}.png"
        saved_img = save_receipt_image(receipt, img_path)
        pages.append(fill_template(receipt, req.card_name, req.user_name, req.purpose, saved_img))

    final_doc = pages[0]
    for page in pages[1:]:
        append_doc_body(final_doc, page)

    filename = req.output_filename or "법인카드_영수증_증빙자료.docx"
    if not filename.endswith(".docx"):
        filename += ".docx"

    out_path = tmp_dir / filename
    final_doc.save(out_path)

    encoded_filename = urllib.parse.quote(filename)
    base_url = str(request.base_url).rstrip("/")
    download_url = f"{base_url}/download/{folder_id}/{encoded_filename}"

    return {
        "status": "success",
        "message": "DOCX 파일이 생성되었습니다. download_url에서 파일을 내려받을 수 있습니다.",
        "download_url": download_url,
        "filename": filename,
    }
