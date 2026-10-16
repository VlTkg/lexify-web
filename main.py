import os
import io
import httpx
from fastapi import FastAPI, HTTPException, UploadFile, File, Form, Request
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import Optional, List
from google import genai
from pypdf import PdfReader
import docx

# Импорты для работы с SQLAlchemy и PostgreSQL
from sqlalchemy import create_engine, Column, BigInteger, Integer, String, Text, DateTime, Boolean
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
from datetime import datetime

app = FastAPI()

# Монтируем статические файлы
app.mount("/static", StaticFiles(directory="static"), name="static")

# Переменные окружения
FEEDBACK_WEBHOOK_URL = os.environ.get("GOOGLE_SHEET_WEBHOOK_URL", "")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
DATABASE_URL = os.environ.get("DATABASE_URL", "")
BOT_TOKEN = os.environ.get("BOT_TOKEN", "") # Токен вашего Telegram бота от BotFather (нужен для выставления инвойсов Stars)

# Настройка подключения к PostgreSQL через SQLAlchemy
if DATABASE_URL and DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

engine = None
SessionLocal = None
Base = declarative_base()

# Модель пользователя в базе данных
class UserDB(Base):
    __tablename__ = "users"

    telegram_id = Column(BigInteger, primary_key=True, index=True)
    first_name = Column(String, default="")
    username = Column(String, default="")
    free_checks_left = Column(Integer, default=3)
    is_pro = Column(Boolean, default=False)
    created_at = Column(DateTime, default=datetime.utcnow)

# Модель логов анализа
class AnalysisLogDB(Base):
    __tablename__ = "analysis_logs"

    id = Column(Integer, primary_key=True, index=True)
    telegram_id = Column(BigInteger, nullable=True)
    doc_type = Column(String, default="")
    jurisdiction = Column(String, default="")
    risk_level = Column(String, default="")
    created_at = Column(DateTime, default=datetime.utcnow)

if DATABASE_URL:
    try:
        engine = create_engine(DATABASE_URL)
        SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
        Base.metadata.create_all(bind=engine)
        print("База данных подключена, таблицы успешно инициализированы.")
    except Exception as e:
        print(f"Ошибка подключения к базе данных: {e}")

# Инициализация клиента Gemini
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
    package_type: str = "5_checks" # Варианты: '5_checks' или '15_checks'

# Модели для структурированного ответа анализа рисков
class RiskItem(BaseModel):
    category: str  # 'red', 'yellow', 'green'
    title: str     # Краткий заголовок риска
    quote: str     # Цитата или упоминание пункта договора
    explanation: str  # В чем заключается риск для фрилансера/исполнителя
    recommendation: str  # Готовый текст правки или рекомендация для общения с заказчиком

class AnalysisResponse(BaseModel):
    overall_risk_level: str  # 'Низкий', 'Средний', 'Высокий'
    summary: str  # Краткое общее резюме (1-2 предложения)
    risks: List[RiskItem]

@app.get("/")
async def serve_index():
    return FileResponse("static/index.html")

# Эндпоинт для получения информации о пользователе и остатке лимита
@app.get("/user-info/{telegram_id}")
async def get_user_info(telegram_id: int):
    if not SessionLocal:
        return {"telegram_id": telegram_id, "free_checks_left": 3, "is_pro": False}
    
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
            "is_pro": user.is_pro
        }
    finally:
        db.close()

# Эндпоинт создания инвойса оплаты через Telegram Stars
@app.post("/create-invoice")
async def create_invoice(data: InvoiceRequest):
    if not BOT_TOKEN:
        raise HTTPException(status_code=500, detail="BOT_TOKEN не настроен на сервере.")
    
    # Настройка пакетов (цены в Telegram Stars XTR)
    packages = {
        "5_checks": {"title": "Пакет 5 проверок", "description": "5 дополнительных экспресс-анализов договоров", "amount": 50},
        "15_checks": {"title": "Пакет 15 проверок", "description": "15 дополнительных экспресс-анализов договоров", "amount": 120}
    }
    
    pkg = packages.get(data.package_type, packages["5_checks"])
    
    # Запрос к Telegram Bot API для создания ссылки на оплату (createInvoiceLink)
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/createInvoiceLink"
    payload = {
        "title": pkg["title"],
        "description": pkg["description"],
        "payload": f"pack_{data.package_type}_{data.telegram_id}",
        "currency": "XTR", # Обялюта Telegram Stars
        "prices": [{"label": pkg["title"], "amount": pkg["amount"]}]
    }
    
    async with httpx.AsyncClient() as client:
        response = await client.post(url, json=payload)
        res_data = response.json()
        if res_data.get("ok"):
            return {"invoice_link": res_data["result"]}
        else:
            raise HTTPException(status_code=500, detail=f"Ошибка создания счета Telegram: {res_data.get('description')}")

# Webhook для получения обновлений от Telegram (обработка успешных платежей)
@app.post("/telegram-webhook")
async def telegram_webhook(request: Request):
    body = await request.json()
    
    # Проверяем успешную оплату (successful_payment)
    if "message" in body and "successful_payment" in body["message"]:
        payment = body["message"]["successful_payment"]
        payload = payment.get("invoice_payload", "")
        user_id = body["message"]["from"]["id"]
        
        # Разбираем payload (например: pack_5_checks_12345678)
        if payload.startswith("pack_"):
            parts = payload.split("_")
            if len(parts) >= 3:
                checks_to_add = int(parts[1]) if parts[1].isdigit() else 5
                
                if SessionLocal:
                    db = SessionLocal()
                    try:
                        user = db.query(UserDB).filter(UserDB.telegram_id == user_id).first()
                        if user:
                            user.free_checks_left += checks_to_add
                            db.commit()
                            print( пользователю {user_id} начислено {checks_to_add} проверок за Telegram Stars)
                    except Exception as e:
                        print(f"Ошибка начисления проверок по платежу: {e}")
                    finally:
                        db.close()
                        
    return {"status": "ok"}

@app.post("/feedback")
async def receive_feedback(data: FeedbackSchema):
    if not FEEDBACK_WEBHOOK_URL:
        print(f"Feedback received (no webhook configured): {data}")
        return {"status": "ok", "message": "Logged locally"}

    try:
        async with httpx.AsyncClient() as client:
            response = await client.post(
                FEEDBACK_WEBHOOK_URL,
                json=data.dict(),
                timeout=10.0
            )
            if response.status_code == 200:
                return {"status": "ok", "message": "Feedback saved to Google Sheet"}
            else:
                raise HTTPException(status_code=500, detail="Failed to post to Google Sheet")
    except Exception as e:
        print(f"Error sending feedback: {e}")
        return {"status": "error", "detail": str(e)}

@app.post("/analyze")
async def analyze_contract(data: AnalyzeSchema):
    if not ai_client:
        raise HTTPException(status_code=500, detail="GEMINI_API_KEY не настроен на сервере.")

    contract_content = data.text.strip()
    if not contract_content:
        raise HTTPException(status_code=400, detail="Текст договора пуст.")

    # Проверка и списание бесплатных лимитов с учетом BigInteger
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
            
            # Если не Pro и проверки кончились
            if not user.is_pro:
                if user.free_checks_left <= 0:
                    raise HTTPException(
                        status_code=403, 
                        detail="У вас закончились бесплатные проверки. Пополните баланс звездами Telegram!"
                    )
                user.free_checks_left -= 1
                db.commit()
                db.refresh(user)
            
            checks_left = user.free_checks_left
        except HTTPException as he:
            raise he
        except Exception as db_err:
            print(f"Ошибка работы с БД пользователей: {db_err}")
        finally:
            db.close()

    # Промпт для юриста-эксперта с фокусом на интересы исполнителя
    prompt = f"""
Ты — профессиональный юрист-аналитик, специализирующийся на защите интересов фрилансеров, подрядчиков и специалистов.
Проведи экспресс-анализ представленного договора и выяви потенциальные риски для Исполнителя.
Пиши современным, понятным и деловым языком. Избегай архаичных юридических канцеляризмов (не используй слова вроде «кабальный», заменяй их на «неприемлемые условия», «критические риски» или «односторонние невыгодные ограничения»).

Контекст анализа:
- Тип договора / Занятость: {data.doc_type}
- Юрисдикция: {data.jurisdiction}

Текст договора:
---
{contract_content[:15000]}
---

Сформируй разбор строго в формате JSON:
1. `overall_risk_level`: Оцени общий уровень риска (строго одно из значений: 'Низкий', 'Средний', 'Высокий').
2. `summary`: Краткий вывод по договору в 1-2 предложениях (без канцелярита).
3. `risks`: Список выявленных пунктов и рисков (3-6 пунктов). Для каждого укажи:
   - `category`: Цветовой маркер риска ('red' — критический, 'yellow' — средний/внимание, 'green' — безопасное/выгодное условие).
   - `title`: Четкое название риска (например, "Неограниченный штраф за задержку", "Передача прав до оплаты").
   - `quote`: Цитата или фрагмент пункта договора (если применимо).
   - `explanation`: Чем именно это условие грозит Исполнителю.
   - `recommendation`: Готовая формулировка правки, которую фрилансер может скопировать и выслать заказчику для защиты своих интересов.
"""

    try:
        response = ai_client.models.generate_content(
            model="gemini-3.8-flash",
            contents=prompt,
            config={
                "response_mime_type": "application/json",
                "response_schema": AnalysisResponse,
            }
        )
        
        # Сохранение лога анализа в базу данных
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
                print(f"Ошибка сохранения лога в БД: {db_err}")

        # Возвращаем результат анализа и остаток бесплатных проверок
        return {
            "status": "ok", 
            "result": response.text,
            "checks_left": checks_left
        }
    except HTTPException as he:
        raise he
    except Exception as e:
        print(f"Gemini API error: {e}")
        raise HTTPException(status_code=500, detail=f"Ошибка анализа через Gemini API: {str(e)}")
