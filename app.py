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
from docx.shared import Inches
from PIL import Image, ImageOps, UnidentifiedImageError
import base64
import uuid
import re
import shutil

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


def set_cell_alignment(cell, paragraph_align=WD_ALIGN_PARAGRAPH.CENTER):
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for paragraph in cell.paragraphs:
        paragraph.alignment = paragraph_align


def copy_run_style(src_run, dst_run):
    if not src_run:
        return
    dst_run.bold = src_run.bold
    dst_run.italic = src_run.italic
    dst_run.underline = src_run.underline
    dst_run.font.name = src_run.font.name
    dst_run.font.size = src_run.font.size
    dst_run.font.bold = src_run.font.bold
    dst_run.font.italic = src_run.font.italic
    dst_run.font.underline = src_run.font.underline
    if src_run.font.color and src_run.font.color.rgb:
        dst_run.font.color.rgb = src_run.font.color.rgb


def set_paragraph_text_keep_style(paragraph, text: str, align=WD_ALIGN_PARAGRAPH.CENTER):
    """기존 첫 run의 서식을 최대한 유지해서 문단 텍스트만 교체한다."""
    src_run = paragraph.runs[0] if paragraph.runs else None
    for run in paragraph.runs:
        run.text = ""
    dst_run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
    copy_run_style(src_run, dst_run)
    dst_run.text = text or ""
    paragraph.alignment = align


def clear_cell_keep_first_paragraph(cell):
    # 첫 문단만 남기고 나머지 문단은 비움. 표/셀 구조는 건드리지 않음.
    for i, p in enumerate(cell.paragraphs):
        if i == 0:
            set_paragraph_text_keep_style(p, "")
        else:
            for r in p.runs:
                r.text = ""


def normalize_amount(value: str) -> str:
    if not value:
        return "금액 확인 필요"
    value = value.strip()
    return value if value.endswith("원") else f"{value}원"


def normalize_base64_image(image_base64: str, output_path: Path, index: int):
    if not image_base64 or len(image_base64.strip()) < 100:
        raise HTTPException(
            status_code=400,
            detail=f"{index}번째 영수증 이미지 데이터가 너무 짧습니다. 실제 이미지 base64가 전달되지 않았습니다."
        )

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

    if len(raw) < 200:
        raise HTTPException(status_code=400, detail=f"{index}번째 영수증 이미지 파일이 비어 있거나 손상되었습니다.")

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


def add_receipt_image_in_cell(cell, img_path: Path):
    """영수증 영역 셀 안에 이미지가 다음 페이지로 밀리지 않도록 최대 폭/높이를 제한한다."""
    clear_cell_keep_first_paragraph(cell)
    set_cell_alignment(cell)

    # 영수증 칸 안에 들어오도록 보수적으로 제한
    max_w_in = 4.25
    max_h_in = 3.70

    with Image.open(img_path) as im:
        w_px, h_px = im.size
    ratio = min(max_w_in / w_px, max_h_in / h_px)  # inches-per-pixel scale 개념
    final_w = w_px * ratio
    final_h = h_px * ratio

    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(str(img_path), width=Inches(final_w), height=Inches(final_h))


def fill_template(doc: Document, receipt: ReceiptItem, card_name: str, user_name: str, purpose: str, img_path: Path):
    if not doc.tables:
        raise HTTPException(status_code=500, detail="템플릿에서 표를 찾지 못했습니다.")
    t = doc.tables[0]

    # 원본 템플릿 셀 위치 기준 입력. 셀/행/열 구조는 절대 변경하지 않음.
    set_paragraph_text_keep_style(t.cell(1, 1).paragraphs[0], card_name, WD_ALIGN_PARAGRAPH.CENTER)
    set_paragraph_text_keep_style(t.cell(1, 3).paragraphs[0], receipt.use_date, WD_ALIGN_PARAGRAPH.CENTER)
    set_paragraph_text_keep_style(t.cell(2, 1).paragraphs[0], user_name, WD_ALIGN_PARAGRAPH.CENTER)
    set_paragraph_text_keep_style(t.cell(3, 1).paragraphs[0], purpose or "", WD_ALIGN_PARAGRAPH.CENTER)

    # 사용처: 참석자 예시 문구 제거 후 상호명만 가운데 정렬
    cell_store = t.cell(4, 1)
    clear_cell_keep_first_paragraph(cell_store)
    set_paragraph_text_keep_style(cell_store.paragraphs[0], receipt.store, WD_ALIGN_PARAGRAPH.CENTER)

    # 사용금액: 첫 문단만 금액으로 교체, 두 번째 비목 문단은 원본 유지
    cell_amount = t.cell(5, 1)
    if cell_amount.paragraphs:
        set_paragraph_text_keep_style(cell_amount.paragraphs[0], normalize_amount(receipt.amount), WD_ALIGN_PARAGRAPH.CENTER)
    # 비목 문단도 원본처럼 가운데 정렬 유지
    for p in cell_amount.paragraphs[1:]:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 전체 입력 셀 가로/세로 가운데 정렬 유지
    for r, c in [(1,1), (1,3), (2,1), (3,1), (4,1), (5,1)]:
        set_cell_alignment(t.cell(r, c))

    # 영수증 이미지는 마지막 빈 행 셀 안에 삽입
    receipt_cell = t.cell(7, 0)
    add_receipt_image_in_cell(receipt_cell, img_path)


def append_template_body(target_doc: Document, source_doc: Document):
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

    final_doc = None
    for i, receipt in enumerate(req.receipts, start=1):
        img_path = tmp_dir / f"receipt_{i}.png"
        normalize_base64_image(receipt.receipt_image_base64, img_path, i)
        page_doc = Document(str(TEMPLATE_PATH))
        fill_template(page_doc, receipt, req.card_name, req.user_name, req.purpose, img_path)
        if final_doc is None:
            final_doc = page_doc
        else:
            append_template_body(final_doc, page_doc)

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
