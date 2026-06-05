from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field
from typing import List, Optional
from pathlib import Path
from copy import deepcopy
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.shared import Inches
import base64, uuid, os

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "template.docx"
OUT_DIR = BASE_DIR / "outputs"
OUT_DIR.mkdir(exist_ok=True)

app = FastAPI(title="Receipt DOCX Generator", version="1.0.0")

class ReceiptItem(BaseModel):
    use_date: str = Field(..., description="YYYY-MM-DD")
    store: str = Field(..., description="영수증 상호명")
    amount: str = Field(..., description="예: 52,000원")
    receipt_image_base64: str = Field(..., description="영수증 원본 이미지 base64")

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
    """Replace text while keeping the first run's style as much as possible."""
    full_text = "".join(run.text for run in paragraph.runs)
    if old not in full_text:
        return False
    new_text = full_text.replace(old, new)
    if paragraph.runs:
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
    """Find the cell/paragraph after the receipt label."""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if "영 수 증" in cell.text or "영수증" in cell.text:
                    return cell
    return None


def append_template_page(target_doc: Document, source_doc: Document):
    # Page break before appending the next template body.
    target_doc.add_page_break()
    body = target_doc.element.body
    for child in source_doc.element.body:
        if child.tag.endswith('sectPr'):
            continue
        body.append(deepcopy(child))


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

    # Append image under receipt label, preserving ratio.
    p = receipt_cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(str(img_path), width=Inches(4.8))
    return doc


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
        try:
            raw = base64.b64decode(receipt.receipt_image_base64.split(',')[-1])
        except Exception:
            raise HTTPException(status_code=400, detail=f"{i}번째 영수증 이미지 base64를 읽지 못했습니다.")
        img_path = tmp_dir / f"receipt_{i}.png"
        img_path.write_bytes(raw)
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

@app.get("/health")
def health():
    return {"status": "ok"}
