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
from PIL import Image, ImageOps, UnidentifiedImageError
import base64
import uuid
import re

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "template.docx"
OUT_DIR = BASE_DIR / "outputs"
OUT_DIR.mkdir(exist_ok=True)

app = FastAPI(title="Receipt DOCX Generator", version="1.0.2")


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


def clear_paragraph_keep_style(paragraph):
    """문단/첫 run 서식은 유지하고 텍스트만 비운다."""
    if not paragraph.runs:
        paragraph.add_run("")
    for run in paragraph.runs:
        run.text = ""


def set_first_run_text_keep_style(paragraph, text: str, font_size_pt: Optional[float] = None, bold: Optional[bool] = None):
    """기존 문단/첫 run 스타일을 그대로 사용하면서 텍스트만 입력한다."""
    if not paragraph.runs:
        paragraph.add_run("")
    paragraph.runs[0].text = text or ""
    if font_size_pt is not None:
        paragraph.runs[0].font.size = Pt(font_size_pt)
    if bold is not None:
        paragraph.runs[0].bold = bold
    # 나머지 run 비움
    for run in paragraph.runs[1:]:
        run.text = ""


def set_cell_value(cell, text: str, font_size_pt: float = 13.0, bold: Optional[bool] = None, center: bool = True):
    """셀 내부의 기존 서식을 최대한 유지하며 값만 교체한다."""
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    if not cell.paragraphs:
        p = cell.add_paragraph()
    else:
        p = cell.paragraphs[0]
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_first_run_text_keep_style(p, text, font_size_pt=font_size_pt, bold=bold)
    # 나머지 문단은 비워서 원본 셀 구조는 유지하되 표시 텍스트만 제거
    for extra_p in cell.paragraphs[1:]:
        extra_p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else extra_p.alignment
        clear_paragraph_keep_style(extra_p)


def set_amount_cell(cell, amount: str):
    """사용금액 셀은 1번째 문단 금액만 교체하고, 2번째 문단 비목은 원본 그대로 유지한다."""
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    if not cell.paragraphs:
        p = cell.add_paragraph()
    else:
        p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_first_run_text_keep_style(p, amount, font_size_pt=13.0, bold=None)
    # 비목 문단은 건드리지 않음. 단, 원본처럼 가운데 정렬 유지
    if len(cell.paragraphs) > 1:
        cell.paragraphs[1].alignment = WD_ALIGN_PARAGRAPH.CENTER


def set_table_cells(doc: Document, receipt: ReceiptItem, card_name: str, user_name: str, purpose: str):
    if not doc.tables:
        raise HTTPException(status_code=500, detail="템플릿에 표가 없습니다.")
    table = doc.tables[0]

    # 원본 표 구조 기준 행/열 입력. 병합 셀은 첫 번째 셀에만 넣어도 같은 셀에 반영됨.
    set_cell_value(table.rows[1].cells[1], card_name, font_size_pt=13.0)
    set_cell_value(table.rows[1].cells[3], receipt.use_date, font_size_pt=13.0)
    set_cell_value(table.rows[2].cells[1], user_name, font_size_pt=13.0)
    set_cell_value(table.rows[3].cells[1], purpose or "", font_size_pt=13.0)
    set_cell_value(table.rows[4].cells[1], receipt.store, font_size_pt=13.0)
    set_amount_cell(table.rows[5].cells[1], receipt.amount)

    # 레이블/헤더 셀도 원본처럼 셀 가운데 정렬 유지
    for row in table.rows:
        for cell in row.cells:
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for p in cell.paragraphs:
                if p.alignment is None:
                    # 템플릿 값 셀은 모두 가운데 기준으로 보이게 함
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER


def receipt_insert_cell(doc: Document):
    table = doc.tables[0]
    # 원본 양식에서 마지막 행이 영수증 이미지 영역임
    return table.rows[-1].cells[0]


def clear_receipt_area(cell):
    for p in cell.paragraphs:
        clear_paragraph_keep_style(p)
    if not cell.paragraphs:
        cell.add_paragraph()


def image_display_width_inches(img_path: Path, max_width=4.8, max_height=4.55):
    """영수증 칸 안에 들어오도록 비율 유지한 표시 폭을 계산한다."""
    with Image.open(img_path) as img:
        w, h = img.size
    if w <= 0 or h <= 0:
        return max_width
    aspect = h / w
    width_by_height = max_height / aspect
    return max(1.0, min(max_width, width_by_height))


def add_receipt_image(doc: Document, img_path: Path):
    cell = receipt_insert_cell(doc)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    clear_receipt_area(cell)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    width_in = image_display_width_inches(img_path, max_width=4.8, max_height=4.55)
    run.add_picture(str(img_path), width=Inches(width_in))


def build_one_page(receipt: ReceiptItem, card_name: str, user_name: str, purpose: str, img_path: Path) -> Document:
    doc = Document(str(TEMPLATE_PATH))
    set_table_cells(doc, receipt, card_name, user_name, purpose)
    add_receipt_image(doc, img_path)
    return doc


def append_template_page(target_doc: Document, source_doc: Document):
    target_doc.add_page_break()
    body = target_doc.element.body
    for child in source_doc.element.body:
        if child.tag.endswith("sectPr"):
            continue
        body.append(deepcopy(child))


@app.get("/")
def root():
    return {"status": "ok", "service": "Receipt DOCX Generator"}


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

    page_docs = []
    for i, receipt in enumerate(req.receipts, start=1):
        img_path = tmp_dir / f"receipt_{i}.png"
        normalize_base64_image(receipt.receipt_image_base64, img_path, i)
        page_docs.append(build_one_page(receipt, req.card_name, req.user_name, req.purpose, img_path))

    final_doc = page_docs[0]
    for doc in page_docs[1:]:
        append_template_page(final_doc, doc)

    safe_name = req.output_filename or "법인카드_영수증_증빙자료.docx"
    if not safe_name.endswith(".docx"):
        safe_name += ".docx"
    out_path = tmp_dir / safe_name
    final_doc.save(out_path)

    return FileResponse(
        out_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=safe_name,
    )
