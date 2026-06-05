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

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "template.docx"
OUT_DIR = BASE_DIR / "outputs"
OUT_DIR.mkdir(exist_ok=True)

app = FastAPI(title="Receipt DOCX Generator", version="1.0.1")


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


def set_cell_center(cell):
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for paragraph in cell.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER


def replace_text_keep_style(paragraph, old: str, new: str):
    full_text = "".join(run.text for run in paragraph.runs)
    if old not in full_text:
        return False

    new_text = full_text.replace(old, new)
    if paragraph.runs:
        # Keep the first run style, clear the rest.
        paragraph.runs[0].text = new_text
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(new_text)

    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return True


def replace_everywhere(doc: Document, replacements: dict):
    for p in doc.paragraphs:
        for old, new in replacements.items():
            replace_text_keep_style(p, old, new)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                set_cell_center(cell)
                for p in cell.paragraphs:
                    for old, new in replacements.items():
                        replace_text_keep_style(p, old, new)


def find_receipt_cell(doc: Document):
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if "영 수 증" in cell.text or "영수증" in cell.text:
                    return cell
    return None


def append_template_page(target_doc: Document, source_doc: Document):
    target_doc.add_page_break()
    body = target_doc.element.body
    for child in source_doc.element.body:
        if child.tag.endswith("sectPr"):
            continue
        body.append(deepcopy(child))


def normalize_base64_image(image_base64: str, output_path: Path, index: int):
    """
    GPT Actions에서 넘어온 이미지 base64가 data:image/png;base64,... 형태이거나
    줄바꿈/공백이 섞여 있어도 처리한다.
    깨진 이미지면 500이 아니라 400으로 명확히 반환한다.
    """
    if not image_base64 or len(image_base64.strip()) < 100:
        raise HTTPException(
            status_code=400,
            detail=f"{index}번째 영수증 이미지 데이터가 너무 짧습니다. 실제 이미지 base64가 전달되지 않았습니다."
        )

    data = image_base64.strip()

    # data URL prefix 제거
    if "," in data and data.lower().startswith("data:image"):
        data = data.split(",", 1)[1]

    # 공백/줄바꿈 제거
    data = re.sub(r"\s+", "", data)

    # base64 padding 보정
    missing_padding = len(data) % 4
    if missing_padding:
        data += "=" * (4 - missing_padding)

    try:
        raw = base64.b64decode(data, validate=False)
    except Exception:
        raise HTTPException(
            status_code=400,
            detail=f"{index}번째 영수증 이미지 base64를 디코딩하지 못했습니다."
        )

    if len(raw) < 200:
        raise HTTPException(
            status_code=400,
            detail=f"{index}번째 영수증 이미지 파일이 비어 있거나 손상되었습니다."
        )

    try:
        img = Image.open(BytesIO(raw))
        img = ImageOps.exif_transpose(img)
        img.load()
    except (UnidentifiedImageError, OSError, ValueError):
        raise HTTPException(
            status_code=400,
            detail=f"{index}번째 영수증 이미지가 올바른 이미지 파일이 아닙니다. base64가 잘렸을 가능성이 큽니다."
        )

    # python-docx가 안정적으로 읽도록 PNG로 재저장
    if img.mode not in ("RGB", "RGBA"):
        img = img.convert("RGB")

    img.save(output_path, format="PNG")
    return output_path


def build_one_page(receipt: ReceiptItem, card_name: str, user_name: str, purpose: str, img_path: Path) -> Document:
    doc = Document(str(TEMPLATE_PATH))
    replacements = {
        "★★★팀 법인카드(0000)": card_name,
        "2025-09-04": receipt.use_date,
        "★★★팀 ★★★ 매니저": user_name,
        "◎◎◎◎ 관련 연구회의 회의비": purpose,
        "**** 식당": receipt.store,
        "70,000원": receipt.amount,
    }
    replace_everywhere(doc, replacements)

    receipt_cell = find_receipt_cell(doc)
    if receipt_cell is None:
        raise HTTPException(status_code=500, detail="템플릿에서 영수증 삽입 위치를 찾지 못했습니다.")

    p = receipt_cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(str(img_path), width=Inches(4.8))
    return doc


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
