import os
import io
import json
import httpx
import re
import unicodedata
import tempfile
from datetime import datetime
from fastapi import FastAPI, HTTPException, Request
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, StreamingResponse, HTMLResponse
from pydantic import BaseModel
from typing import Optional, List
from google import genai

from sqlalchemy import create_engine, Column, BigInteger, Integer, String, Text, DateTime, Boolean
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

# Импорт наших генераторов
from contract_generator import generate_legal_document
from protocol_generator import generate_disagreements_protocol

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
    free_checks_left = Column(Integer, default=3)
    is_pro = Column(Boolean, default=False)
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
    category: str
    title: str
    quote: str
    explanation: str
    recommendation: str

class AnalysisResponse(BaseModel):
    overall_risk_level: str
    summary: str
    risks: List[RiskItem]

class ProtocolItem(BaseModel):
    title: str
    quote: str
    explanation: str
    recommendation: str

class ProtocolRequest(BaseModel):
    doc_type: str = "Общий"
    jurisdiction: str = "Россия (РФ)"
    telegram_id: Optional[int] = None
    risks: List[ProtocolItem]

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

class ConstructorFileRequest(BaseModel):
    document_text: str
    doc_type: str = "contract"
    jurisdiction: str = "Россия (РФ)"

@app.get("/")
async def serve_landing():
    return FileResponse("static/landing.html")

@app.get("/app")
async def serve_app():
    return FileResponse("static/index.html")

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
    is_admin = telegram_id in ADMIN_IDS or telegram_id == 0
    if is_admin:
        return {"telegram_id": telegram_id, "free_checks_left": 999, "is_pro": True, "is_admin": True}

    if not SessionLocal:
        return {"telegram_id": telegram_id, "free_checks_left": 3, "is_pro": False, "is_admin": False}
    
    db = SessionLocal()
    try:
        user = db.query(UserDB).filter(UserDB.telegram_id == telegram_id).first()
        if not user:
            user = UserDB(telegram_id=telegram_id, free_checks_left=3, is_pro=False)
            db.add(user)
            db.commit()
            db.refresh(user)
        return {
            "telegram_id": user.telegram_id,
            "free_checks_left": user.free_checks_left,
            "is_pro": user.is_pro,
            "is_admin": False
        }
    finally:
        db.close()

@app.get("/admin/stats/{telegram_id}")
async def get_admin_stats(telegram_id: int):
    if telegram_id not in ADMIN_IDS and telegram_id != 0:
        raise HTTPException(status_code=403, detail="Доступ запрещен")
    
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
                "free_checks_left": u.free_checks_left,
                "is_pro": u.is_pro,
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
    body = await request.json()
    
    if "message" in body and "text" in body["message"]:
        text = body["message"]["text"]
        chat_id = body["message"]["chat"]["id"]
        
        if text.startswith("/start"):
            if BOT_TOKEN:
                welcome_text = (
                    "👋 **Привет! Я Lexify — ваш ИИ-юрист и конструктор документов.**\n\n"
                    "Я помогу за 1 минуту выявить скрытые риски в договоре, сформировать Протокол разногласий в Word (.docx) или с нуля составить безопасный договор/оферту.\n\n"
                    "🎁 **Вам доступно 3 бесплатные проверки!**"
                )
                reply_markup = {
                    "inline_keyboard": [
                        [
                            {
                                "text": "🛠️ Открыть Lexify (Анализ и Конструктор)",
                                "web_app": {"url": "https://lexify-web.onrender.com/app"}
                            }
                        ]
                    ]
                }
                url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage"
                async with httpx.AsyncClient() as client:
                    await client.post(url, json={
                        "chat_id": chat_id,
                        "text": welcome_text,
                        "parse_mode": "Markdown",
                        "reply_markup": reply_markup
                    })
            return {"status": "ok"}

    if "message" in body and "successful_payment" in body["message"]:
        payment = body["message"]["successful_payment"]
        payload = payment.get("invoice_payload", "")
        user_id = body["message"]["from"]["id"]
        charge_id = payment.get("telegram_payment_charge_id", "")
        
        if payload.startswith("pack_"):
            parts = payload.split("_")
            if len(parts) >= 3:
                checks_to_add = int(parts[1]) if parts[1].isdigit() else 5
                
                if SessionLocal:
                    db = SessionLocal()
                    try:
                        existing_payment = db.query(PaymentLogDB).filter(PaymentLogDB.charge_id == charge_id).first()
                        if existing_payment:
                            return {"status": "ok"}
                        
                        user = db.query(UserDB).filter(UserDB.telegram_id == user_id).first()
                        if user:
                            user.free_checks_left += checks_to_add
                            new_payment_log = PaymentLogDB(telegram_id=user_id, charge_id=charge_id, payload=payload)
                            db.add(new_payment_log)
                            db.commit()
                    except Exception as e:
                        db.rollback()
                    finally:
                        db.close()
                        
    return {"status": "ok"}

def clean_text(val: Optional[str]) -> str:
    if not val:
        return ""
    text_str = unicodedata.normalize("NFKC", str(val))
    text_str = re.sub(r'[\xa0\u200b\u202f\u200e\u200f\ufeff]', ' ', text_str)
    text_str = re.sub(r'[^\w\s\d.,!?;:()"\'-–—«»/%\\]', '', text_str)
    return text_str.strip()

@app.post("/generate-protocol")
async def generate_protocol(data: ProtocolRequest):
    # Превращаем входящие риски в формат, ожидаемый protocol_generator
    formatted_risks = []
    for r in data.risks:
        formatted_risks.append({
            "level": "red",
            "clause_num": clean_text(r.title),
            "description": clean_text(r.quote),
            "suggested_fix": clean_text(r.recommendation),
            "reasoning": clean_text(r.explanation)
        })
    
    analysis_data = {"risks": formatted_risks}
    contract_title = f"проекту договора ({clean_text(data.doc_type)}), юрисдикция: {clean_text(data.jurisdiction)}"
    
    file_stream = generate_disagreements_protocol(analysis_data, contract_title)
    
    return StreamingResponse(
        file_stream,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": "attachment; filename=Protocol_Raznogasiy.docx"}
    )

@app.post("/api/generate-document")
async def api_generate_document(data: ContractGenerationRequest):
    document_text = generate_legal_document(data.dict())
    if document_text.startswith("Ошибка при генерации"):
        raise HTTPException(status_code=500, detail=document_text)
    return {"status": "success", "document": document_text}

@app.post("/feedback")
async def receive_feedback(data: FeedbackSchema):
    if not FEEDBACK_WEBHOOK_URL:
        return {"status": "ok", "message": "Logged locally"}

    try:
        async with httpx.AsyncClient() as client:
            response = await client.post(FEEDBACK_WEBHOOK_URL, json=data.dict(), timeout=10.0)
            if response.status_code == 200:
                return {"status": "ok", "message": "Feedback saved"}
            else:
                raise HTTPException(status_code=500, detail="Failed to post feedback")
    except Exception as e:
        return {"status": "error", "detail": str(e)}

@app.post("/analyze")
async def analyze_contract(data: AnalyzeSchema):
    if not ai_client:
        raise HTTPException(status_code=500, detail="GEMINI_API_KEY не настроен.")

    contract_content = clean_text(data.text)
    if not contract_content:
        raise HTTPException(status_code=400, detail="Текст договора пуст.")

    checks_left = 3
    if data.telegram_id and SessionLocal:
        db = SessionLocal()
        try:
            user = db.query(UserDB).filter(UserDB.telegram_id == int(data.telegram_id)).first()
            if not user:
                user = UserDB(
                    telegram_id=int(data.telegram_id),
                    first_name=data.first_name or "",
                    username=data.username or "",
                    free_checks_left=3
                )
                db.add(user)
                db.commit()
                db.refresh(user)
            else:
                if data.first_name and user.first_name != data.first_name:
                    user.first_name = data.first_name
                if data.username and user.username != data.username:
                    user.username = data.username
                db.commit()
            
            if int(data.telegram_id) in ADMIN_IDS or user.is_pro:
                checks_left = 999
            else:
                if user.free_checks_left <= 0:
                    raise HTTPException(status_code=403, detail="Закончились проверки.")
                user.free_checks_left -= 1
                db.commit()
                db.refresh(user)
                checks_left = user.free_checks_left
        except HTTPException as he:
            raise he
        except Exception as db_err:
            print(f"Ошибка БД: {db_err}")
        finally:
            db.close()

    prompt = f"""
Ты — профессиональный юрист-аналитик, специализирующийся на защите интересов фрилансеров и исполнителей.
Проведи экспресс-анализ договора. Язык — современный и понятный (без канцеляризмов).

Контекст:
- Тип договора: {data.doc_type}
- Юрисдикция: {data.jurisdiction}

Текст договора:
---
{contract_content[:15000]}
---

Верни JSON строго по схеме:
1. `overall_risk_level`: ('Низкий', 'Средний', 'Высокий').
2. `summary`: Вывод в 1-2 предложениях.
3. `risks`: Список рисков (3-6 штук). Поля:
    - `category`: ('red', 'yellow', 'green')
    - `title`: Название риска
    - `quote`: Обязательно начни с точного номера пункта или статьи договора (например: "Пункт 3.2. ...", "Статья 5.1. ..."). Опиши условие Заказчика простыми словами.
    - `explanation`: Чем грозит данное условие Исполнителю
    - `recommendation`: Предлагаемая безопасная редакция пункта для Исполнителя
"""

    try:
        response = ai_client.models.generate_content(
            model="gemini-2.5-flash",
            contents=prompt,
            config={
                "response_mime_type": "application/json",
                "response_schema": AnalysisResponse,
            }
        )
        
        if SessionLocal:
            try:
                db = SessionLocal()
                db_log = AnalysisLogDB(
                    telegram_id=int(data.telegram_id) if data.telegram_id else None,
                    doc_type=data.doc_type,
                    jurisdiction=data.jurisdiction,
                    risk_level="Анализ выполнен"
                )
                db.add(db_log)
                db.commit()
                db.close()
            except Exception as db_err:
                print(f"Ошибка логирования: {db_err}")

        return {
            "status": "ok", 
            "result": response.text,
            "checks_left": checks_left
        }
    except HTTPException as he:
        raise he
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка Gemini API: {str(e)}")
