import os
import io
import json
import httpx
import re
import unicodedata
from datetime import datetime
from fastapi import FastAPI, HTTPException, Request, UploadFile, File, Form
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse
from pydantic import BaseModel
from typing import Optional, List, Any
from google import genai
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

from sqlalchemy import create_engine, Column, BigInteger, Integer, String, Text, DateTime, Boolean
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

# Инициализация приложения
app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")

FEEDBACK_WEBHOOK_URL = os.environ.get("GOOGLE_SHEET_WEBHOOK_URL", "")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
DATABASE_URL = os.environ.get("DATABASE_URL", "")
BOT_TOKEN = os.environ.get("BOT_TOKEN", "")

ADMIN_IDS = [8804083078]

if DATABASE_URL and DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

engine = None
SessionLocal = None
Base = declarative_base()

class UserDB(Base):
    __tablename__ = "users"

    telegram_id = Column(BigInteger, primary_key=True, index=True)
    first_name = Column(String, default="")
    username = Column(String, default="")
    free_checks_left = Column(Integer, default=999)
    is_pro = Column(Boolean, default=True)
    created_at = Column(DateTime, default=datetime.utcnow)

class AnalysisLogDB(Base):
    __tablename__ = "analysis_logs"

    id = Column(Integer, primary_key=True, index=True)
    telegram_id = Column(BigInteger, nullable=True)
    doc_type = Column(String, default="")
    jurisdiction = Column(String, default="")
    risk_level = Column(String, default="")
    created_at = Column(DateTime, default=datetime.utcnow)

class PaymentLogDB(Base):
    __tablename__ = "payment_logs"

    id = Column(Integer, primary_key=True, index=True)
    telegram_id = Column(BigInteger, nullable=False)
    charge_id = Column(String, unique=True, index=True, nullable=False)
    payload = Column(String, default="")
    created_at = Column(DateTime, default=datetime.utcnow)

if DATABASE_URL:
    try:
        engine = create_engine(DATABASE_URL)
        SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
        Base.metadata.create_all(bind=engine)
    except Exception as e:
        print(f"Ошибка БД: {e}")

ai_client = None
if GEMINI_API_KEY:
    ai_client = genai.Client(api_key=GEMINI_API_KEY)

class FeedbackSchema(BaseModel):
    type: str = ""
    rating: int = 0
    sentiment: str = ""
    feedback: str = ""
    doc_type: str = ""
    jurisdiction: str = ""

class AnalyzeSchema(BaseModel):
    text: str
    doc_type: str = "Общий"
    jurisdiction: str = "Россия (РФ)"
    telegram_id: Optional[int] = None
    first_name: Optional[str] = ""
    username: Optional[str] = ""

class InvoiceRequest(BaseModel):
    telegram_id: int
    package_type: str = "5_checks"

class RiskItem(BaseModel):
    category: Optional[str] = "yellow"
    title: Optional[str] = "Замечание"
    quote: Optional[str] = "—"
    explanation: Optional[str] = "—"
    recommendation: Optional[str] = "—"

class AnalysisResponse(BaseModel):
    overall_risk_level: Optional[str] = "Средний"
    summary: Optional[str] = ""
    risks: Optional[List[RiskItem]] = []

class ProtocolRequest(BaseModel):
    doc_type: str = "Общий"
    jurisdiction: str = "Россия (РФ)"
    telegram_id: Optional[int] = None
    risks: Optional[List[Any]] = []

class ContractGenerationRequest(BaseModel):
    doc_type: str = "contract"
    jurisdiction: str = "Россия (РФ)"
    executor: str = "Самозанятый"
    client: str = "Заказчик"
    description: str = "Оказание услуг"
    payment: str = "Постоплата"
    deadline: str = "По согласованию"
    revisions: str = "2"
    telegram_id: Optional[int] = None

@app.get("/", response_class=HTMLResponse)
async def serve_root():
    if os.path.exists("static/index.html"):
        return FileResponse("static/index.html")
    return HTMLResponse("<h3>Файл static/index.html не найден</h3>", status_code=404)

@app.get("/app", response_class=HTMLResponse)
async def serve_app():
    if os.path.exists("static/index.html"):
        return FileResponse("static/index.html")
    return HTMLResponse("<h3>Файл static/index.html не найден</h3>", status_code=404)

@app.get("/privacy", response_class=HTMLResponse)
async def serve_privacy():
    if os.path.exists("static/privacy.html"):
        return FileResponse("static/privacy.html")
    raise HTTPException(status_code=404, detail="Страница не найдена")

@app.get("/terms", response_class=HTMLResponse)
async def serve_terms():
    if os.path.exists("static/terms.html"):
        return FileResponse("static/terms.html")
    raise HTTPException(status_code=404, detail="Страница не найдена")

@app.get("/user-info/{telegram_id}")
async def get_user_info(telegram_id: int):
    return {"telegram_id": telegram_id, "free_checks_left": 999, "is_pro": True, "is_admin": True}

@app.get("/admin/stats/{telegram_id}")
async def get_admin_stats(telegram_id: int):
    if not SessionLocal:
        return {"total_users": 0, "total_analyses": 0, "total_payments": 0, "users_list": []}
    
    db = SessionLocal()
    try:
        total_users = db.query(UserDB).count()
        total_analyses = db.query(AnalysisLogDB).count()
        total_payments = db.query(PaymentLogDB).count()
        
        users = db.query(UserDB).order_by(UserDB.created_at.desc()).limit(50).all()
        users_list = []
        for u in users:
            users_list.append({
                "telegram_id": u.telegram_id,
                "first_name": u.first_name or "Без имени",
                "username": f"@{u.username}" if u.username else "—",
                "free_checks_left": 999,
                "is_pro": True,
                "created_at": u.created_at.strftime("%d.%m.%Y %H:%M") if u.created_at else "—"
            })

        return {
            "total_users": total_users,
            "total_analyses": total_analyses,
            "total_payments": total_payments,
            "users_list": users_list
        }
    finally:
        db.close()

@app.post("/create-invoice")
async def create_invoice(data: InvoiceRequest):
    if not BOT_TOKEN:
        raise HTTPException(status_code=500, detail="BOT_TOKEN не настроен на сервере.")
    
    packages = {
        "5_checks": {"title": "Пакет 5 проверок", "description": "5 дополнительных экспресс-анализов договоров и генерация протоколов", "amount": 50},
        "15_checks": {"title": "Пакет 15 проверок", "description": "15 дополнительных экспресс-анализов договоров и генерация протоколов", "amount": 120}
    }
    
    pkg = packages.get(data.package_type, packages["5_checks"])
    
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/createInvoiceLink"
    payload = {
        "title": pkg["title"],
        "description": pkg["description"],
        "payload": f"pack_{data.package_type}_{data.telegram_id}",
        "currency": "XTR",
        "prices": [{"label": pkg["title"], "amount": pkg["amount"]}]
    }
    
    async with httpx.AsyncClient() as client:
        response = await client.post(url, json=payload)
        res_data = response.json()
        if res_data.get("ok"):
            return {"invoice_link": res_data["result"]}
        else:
            raise HTTPException(status_code=500, detail=f"Ошибка создания счета Telegram: {res_data.get('description')}")

@app.post("/telegram-webhook")
async def telegram_webhook(request: Request):
    return {"status": "ok"}

def set_cell_background(cell, fill_hex: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def clean_text(val: Optional[str]) -> str:
    if not val:
        return ""
    text_str = unicodedata.normalize("NFKC", str(val))
    text_str = re.sub(r'[\xa0\u200b\u202f\u200e\u200f\ufeff]', ' ', text_str)
    text_str = re.sub(r'[^\w\s\d.,!?;:()"\'-–—«»/%\\]', '', text_str)
    return text_str.strip()

@app.post("/generate-protocol")
async def generate_protocol(data: ProtocolRequest):
    doc = docx.Document()

    for section in doc.sections:
        section.top_margin = Inches(0.59)
        section.bottom_margin = Inches(0.59)
        section.left_margin = Inches(0.59)
        section.right_margin = Inches(0.59)

    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(10)
    normal_style.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run("ПРОТОКОЛ РАЗНОГЛАСИЙ")
    title_run.bold = True
    title_run.font.size = Pt(15)
    title_run.font.color.rgb = RGBColor(0x11, 0x18, 0x27)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = sub_p.add_run(f"к проекту договора / статьи ({clean_text(data.doc_type)})\n")
    sub_run.font.size = Pt(11)
    sub_run.italic = True
    sub_run.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)

    law_p = doc.add_paragraph()
    law_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    law_run = law_p.add_run(f"Применимое право: законодательство ({clean_text(data.jurisdiction)})\n")
    law_run.font.size = Pt(10)
    law_run.font.color.rgb = RGBColor(0x6B, 0x72, 0x80)

    table = doc.add_table(rows=1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'

    col_widths = [Inches(1.2), Inches(2.3), Inches(2.5), Inches(2.0)]

    hdr_cells = table.rows[0].cells
    headers = ["Пункт", "Редакция Заказчика", "Предлагаемая редакция Исполнителя", "Обоснование"]

    for i, header_text in enumerate(headers):
        cell = hdr_cells[i]
        cell.width = col_widths[i]
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        set_cell_background(cell, "1F2937")
        set_cell_margins(cell, top=100, bottom=100, left=100, right=100)

        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(header_text)
        run.bold = True
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    if not data.risks:
        row_cells = table.add_row().cells
        for i in range(4):
            row_cells[i].width = col_widths[i]
        row_cells[0].merge(row_cells[3])
        row_cells[0].text = "Существенных рисков и замечаний к редакции не выявлено."
    else:
        for risk in data.risks:
            row_cells = table.add_row().cells

            for i in range(4):
                row_cells[i].width = col_widths[i]
                set_cell_margins(row_cells[i], top=80, bottom=80, left=80, right=80)

            if isinstance(risk, dict):
                r_title = risk.get('title', '—')
                r_quote = risk.get('quote', '—')
                r_rec = risk.get('recommendation', '—')
                r_exp = risk.get('explanation', 'Защита интересов исполнителя')
            else:
                r_title = getattr(risk, 'title', '—')
                r_quote = getattr(risk, 'quote', '—')
                r_rec = getattr(risk, 'recommendation', '—')
                r_exp = getattr(risk, 'explanation', 'Защита интересов исполнителя')

            row_cells[0].text = clean_text(r_title) or "—"
            if row_cells[0].paragraphs[0].runs:
                row_cells[0].paragraphs[0].runs[0].bold = True
                row_cells[0].paragraphs[0].runs[0].font.size = Pt(9.5)

            row_cells[1].text = clean_text(r_quote) or "—"
            if row_cells[1].paragraphs[0].runs:
                row_cells[1].paragraphs[0].runs[0].font.size = Pt(9.5)

            row_cells[2].text = clean_text(r_rec) or "—"
            if row_cells[2].paragraphs[0].runs:
                row_cells[2].paragraphs[0].runs[0].font.size = Pt(9.5)

            row_cells[3].text = clean_text(r_exp) or "Защита интересов исполнителя"
            if row_cells[3].paragraphs[0].runs:
                row_cells[3].paragraphs[0].runs[0].font.size = Pt(9.5)

    doc.add_paragraph("\n")
    doc.add_paragraph("Настоящий Протокол разногласий является неотъемлемой частью Договора.\n")

    sig_table = doc.add_table(rows=1, cols=2)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    cell_left = sig_table.rows[0].cells[0]
    cell_right = sig_table.rows[0].cells[1]

    cell_left.width = Inches(3.8)
    cell_right.width = Inches(3.8)

    cell_left.text = "ПОДРЯДЧИК / ИСПОЛНИТЕЛЬ:\n\n\n_______________________ / ____________ /\nМ.П."
    cell_right.text = "ЗАКАЗЧИК:\n\n\n_______________________ / ____________ /\nМ.П."

    # Возвращаем файл напрямую через память (StreamingResponse), чтобы избежать проблем с временными файлами на Render
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)

    headers = {
        'Content-Disposition': 'attachment; filename="Protocol_Raznogasiy.docx"'
    }
    return StreamingResponse(
        file_stream, 
        headers=headers, 
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

async def process_document_generation(data: ContractGenerationRequest):
    if not ai_client:
        raise HTTPException(status_code=500, detail="GEMINI_API_KEY не настроен.")

    doc_type = data.doc_type
    jurisdiction = clean_text(data.jurisdiction) or "Россия (РФ)"
    executor = clean_text(data.executor) or "Самозанятый"
    client = clean_text(data.client) or "Заказчик"
    description = clean_text(data.description) or "Оказание услуг"
    payment = clean_text(data.payment) or "Постоплата"
    deadline = clean_text(data.deadline) or "По согласованию"
    revisions = clean_text(data.revisions) or "2"

    if doc_type == 'contract':
        prompt = f"""
Ты — опытный юрист по IT-контрактам и защите интересов фрилансеров в юрисдикции: {jurisdiction}.
Сгенерируй юридически грамотный, структурированный Договор оказания услуг / подряда в формате Markdown с учетом применимого законодательства ({jurisdiction}).

Стороны:
- Исполнитель: {executor}
- Заказчик: {client}

Суть и условия:
- Предмет: {description}
- Порядок оплаты: {payment}
- Сроки выполнения: {deadline}

ОБЯЗАТЕЛЬНЫЕ ЗАЩИТНЫЕ УСЛОВИЯ (в пользу Исполнителя):
1. Порядок приемки: если Заказчик в течение 3 рабочих дней после получения акта не подписал его и не предоставил мотивированный отказ, работы считаются принятыми в полном объеме и подлежат оплате.
2. Ограничение правок: допускается строго {revisions} бесплатная итерация правок в рамках первоначального ТЗ. Доработки сверх ТЗ или сверх лимита оплачиваются отдельно по часовой ставке.
3. Интеллектуальные права: исключительные права на результаты работ переходят к Заказчику ТОЛЬКО после полной оплаты (100%) стоимости услуг.
4. Сдвиг сроков: сроки автоматически сдвигаются на время задержки предоставления доступов, материалов или ответов со стороны Заказчика.

Оформи документ профессионально, с номерами разделов, полями для реквизитов и подписей.
"""
    else:
        prompt = f"""
Ты — профессиональный юрист. Сгенерируй Публичную оферту для оказания услуг в юрисдикции: {jurisdiction} в формате Markdown.
Описание услуг: {description}.
Условия оплаты: {payment}.
Сроки: {deadline}.
Включи стандартные юридические разделы: акцепт оферты, предмет, права и обязанности сторон, финансовые условия, порядок сдачи-приемки, ограничение ответственности исполнителя и порядок расторжения с учетом законодательства ({jurisdiction}).
"""

    try:
        response = ai_client.models.generate_content(
            model='gemini-3.8-flash',
            contents=prompt
        )
        doc_text = response.text
        return {
            "status": "success", 
            "document": doc_text,
            "text": doc_text,
            "result": doc_text,
            "response": doc_text,
            "content": doc_text
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка генерации документа через Gemini: {str(e)}")

@app.post("/generate-document")
async def generate_document_main(data: ContractGenerationRequest):
    return await process_document_generation(data)

@app.post("/api/generate-document")
async def generate_document_api(data: ContractGenerationRequest):
    return await process_document_generation(data)

@app.post("/feedback")
async def receive_feedback(data: FeedbackSchema):
    return {"status": "ok", "message": "Feedback saved"}

async def process_analysis(data: AnalyzeSchema):
    if not ai_client:
        raise HTTPException(status_code=500, detail="GEMINI_API_KEY не настроен.")

    contract_content = clean_text(data.text)
    if not contract_content:
        raise HTTPException(status_code=400, detail="Текст договора пуст.")

    prompt = f"""
Ты — профессиональный юрист-аналитик, специализирующийся на защите интересов фрилансеров и исполнителей.
Проведи экспресс-анализ договора или статьи. Язык ответа — русский.

Контекст:
- Тип: {data.doc_type}
- Юрисдикция: {data.jurisdiction}

Текст:
---
{contract_content[:15000]}
---

Верни ответ СТРОГО в формате JSON, соответствующем схеме AnalysisResponse:
- `overall_risk_level`: строка ('Низкий', 'Средний' или 'Высокий')
- `summary`: краткое резюме (1-2 предложения)
- `risks`: массив объектов, где каждый объект содержит:
    - `category`: 'red' (высокий риск), 'yellow' (средний) или 'green' (низкий/безопасно)
    - `title`: название пункта или риска
    - `quote`: точная цитата из текста
    - `explanation`: чем это грозит исполнителю
    - `recommendation`: безопасная редакция или рекомендация
"""

    try:
        response = ai_client.models.generate_content(
            model='gemini-3.8-flash',
            contents=prompt,
            config={
                "response_mime_type": "application/json",
                "response_schema": AnalysisResponse,
            }
        )
        
        try:
            parsed_json = json.loads(response.text)
        except Exception:
            parsed_json = response.text

        return {
            "status": "ok", 
            "result": parsed_json,
            "checks_left": 999
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка Gemini API: {str(e)}")

@app.post("/analyze")
async def analyze_contract_main(data: AnalyzeSchema):
    return await process_analysis(data)

@app.post("/api/analyze")
async def analyze_contract_api(data: AnalyzeSchema):
    return await process_analysis(data)

@app.post("/api/analyze-file")
async def analyze_file_api(file: UploadFile = File(...), jurisdiction: str = Form("Россия (РФ)"), telegram_id: Optional[int] = Form(None)):
    content_bytes = await file.read()
    extracted_text = ""
    
    if file.filename.endswith(".docx"):
        try:
            doc = docx.Document(io.BytesIO(content_bytes))
            extracted_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Ошибка чтения .docx файла: {str(e)}")
    else:
        extracted_text = content_bytes.decode("utf-8", errors="ignore")

    if not extracted_text.strip():
        raise HTTPException(status_code=400, detail="Загруженный файл пуст или не содержит текста.")

    schema_data = AnalyzeSchema(text=extracted_text, jurisdiction=jurisdiction, telegram_id=telegram_id)
    return await process_analysis(schema_data)
