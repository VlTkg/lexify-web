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

    if "
