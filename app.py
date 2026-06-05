from fastapi import FastAPI, HTTPException
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
from PIL import Image, ImageOps, UnidentifiedImageError
import base64
import uuid
import re

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "template.docx"
OUT_DIR = BASE_DIR / "outputs"
OUT_DIR.mkdir(exist_ok=True)

app = FastAPI(title="Receipt DOCX Generator", version="1.0.4")


class ReceiptItem(BaseModel):
    use_date: str = Field(..., description="YYYY-MM-DD")
    store: str = Field(..., description="영수증 상호명")
    amount: str = Field(..., description="예: 52,000원")
    receipt_image_base64: str = Field(..., description="영수증 원본 이미지 base64 또는 data URL")


class ReceiptRequest(BaseModel):
    receipts: List[ReceiptItem]
    card_name: str = "중앙청년지원센터 법인카드(0000)"
    user_name: str = "중앙청년지원센터 매니저"
    purpose: str = ""
    output_filename: Optional[str] = "법인카드_영수증_증빙자료.docx"


def clear_cell(cell):
    """셀 속성(tcPr)은 유지하고 내용만 삭제한다."""
    tc = cell._tc
    for child in list(tc):
        if child.tag != qn("w:tcPr"):
            tc.remove(child)


def style_run(run, size_pt=13, bold=False):
    run.font.name = "맑은 고딕"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "맑은 고딕")
    run.font.size = Pt(size_pt)
    run.font.bold = bold


def set_center(cell):
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for p in cell.paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER


def write_center(cell, text, size_pt=13, bold=False):
    clear_cell(cell)
    p = cell.add_paragraph()
    try:
        p.style = "바탕글"
    except Exception:
        pass
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    style_run(run, size_pt=size_pt, bold=bold)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def write_amount_cell(cell, amount):
    clear_cell(cell)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    p1 = cell.add_paragraph()
    try:
        p1.style = "바탕글"
    except Exception:
        pass
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p1.add_run(amount)
    style_run(r1, size_pt=13)

    p2 = cell.add_paragraph()
    try:
        p2.style = "바탕글"
    except Exception:
        pass
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run("* 비목: 인증연구 > 업무추진비 > 회의비")
    style_run(r2, size_pt=11)


def normalize_base64_image(image_base64: str, output_path: Path, index: int):
    if not image_base64 or len(image_base64.strip()) < 100:
        raise HTTPException(status_code=400, detail=f"{index}번째 영수증 이미지 데이터가 너무 짧습니다.")

    data = image_base64.strip()
    if "," in data and data.lower().startswith("data:image"):
        data = data.split(",", 1)[1]
    data = re.sub(r"\s+", "", data)
    missing_padding = len(data) % 4
    if missing_padding:
        data += "=" * (4 - missing_padding)

    try:
        raw = base64.b64decode(data, validate=False)
    except Exception:
        raise HTTPException(status_code=400, detail=f"{index}번째 영수증 이미지 base64를 디코딩하지 못했습니다.")

    try:
        img = Image.open(BytesIO(raw))
        img = ImageOps.exif_transpose(img)
        img.load()
    except (UnidentifiedImageError, OSError, ValueError):
        raise HTTPException(status_code=400, detail=f"{index}번째 영수증 이미지가 올바른 이미지 파일이 아닙니다.")

    if img.mode not in ("RGB", "RGBA"):
        img = img.convert("RGB")
    img.save(output_path, format="PNG")
    return output_path


def insert_receipt_image(doc, img_path: Path):
    # 템플릿 기준 8번째 행(인덱스 7)의 빈 영수증 칸에 삽입
    table = doc.tables[0]
    receipt_cell = table.cell(7, 0)
    clear_cell(receipt_cell)
    receipt_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    # 이미지 비율을 유지하면서 한 페이지/영수증 칸 안에 들어오도록 제한
    with Image.open(img_path) as im:
        w_px, h_px = im.size

    max_w_in = 5.7
    max_h_in = 3.45
    aspect = w_px / h_px if h_px else 1
    width_in = min(max_w_in, max_h_in * aspect)
    height_in = width_in / aspect

    if height_in > max_h_in:
        height_in = max_h_in
        width_in = height_in * aspect

    p = receipt_cell.add_paragraph()
    try:
        p.style = "바탕글"
    except Exception:
        pass
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(str(img_path), width=Inches(width_in))


def fill_template(receipt: ReceiptItem, card_name: str, user_name: str, purpose: str, img_path: Path) -> Document:
    doc = Document(str(TEMPLATE_PATH))
    table = doc.tables[0]

    # 원본 표 셀 위치 기준으로 직접 입력
    write_center(table.cell(1, 1), card_name, size_pt=13)
    write_center(table.cell(1, 3), receipt.use_date, size_pt=13)
    write_center(table.cell(2, 1), user_name, size_pt=13)
    write_center(table.cell(3, 1), purpose or "", size_pt=13)
    write_center(table.cell(4, 1), receipt.store, size_pt=13)
    write_amount_cell(table.cell(5, 1), receipt.amount)

    # 모든 셀은 세로 가운데, 라벨/제목 포함 기존 정렬 유지 보정
    for row in table.rows:
        for cell in row.cells:
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for p in cell.paragraphs:
                if p.alignment is None:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    insert_receipt_image(doc, img_path)
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
    return {"status": "ok", "service": "Receipt DOCX Generator", "version": "1.0.4"}


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/create-receipt-docx")
def create_receipt_docx(req: ReceiptRequest):
    if not TEMPLATE_PATH.exists():
        raise HTTPException(status_code=500, detail="template.docx 파일이 서버에 없습니다.")
    if not req.receipts:
        raise HTTPException(status_code=400, detail="영수증 데이터가 없습니다.")

    tmp_dir = OUT_DIR / str(uuid.uuid4())
    tmp_dir.mkdir(parents=True, exist_ok=True)

    pages = []
    for i, receipt in enumerate(req.receipts, start=1):
        img_path = tmp_dir / f"receipt_{i}.png"
        normalize_base64_image(receipt.receipt_image_base64, img_path, i)
        pages.append(fill_template(receipt, req.card_name, req.user_name, req.purpose, img_path))

    final_doc = pages[0]
    for page in pages[1:]:
        append_doc_body(final_doc, page)

    filename = req.output_filename or "법인카드_영수증_증빙자료.docx"
    if not filename.endswith(".docx"):
        filename += ".docx"
    out_path = tmp_dir / filename
    final_doc.save(out_path)

    return FileResponse(
        out_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=filename,
    )
