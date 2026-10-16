import os
import io
import httpx
from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import Optional, List
from google import genai
from pypdf import PdfReader
import docx

# Импорты для работы с SQLAlchemy и PostgreSQL
from sqlalchemy import create_engine, Column, Integer, String, Text, DateTime
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

# Настройка подключения к PostgreSQL через SQLAlchemy
if DATABASE_URL and DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

engine = None
SessionLocal = None
Base = declarative_base()

# Описание модели таблицы для хранения истории анализа или обратной связи
class AnalysisLogDB(Base):
    __tablename__ = "analysis_logs"

    id = Column(Integer, primary_key=True, index=True)
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

    # Промпт для юриста-эксперта с фокусом на интересы исполнителя
    prompt = f"""
Ты — профессиональный юрист-аналитик, специализирующийся на защите интересов фрилансеров, подрядчиков и специалистов.
Проведи экспресс-анализ представленного договора и выяви потенциальные риски для Исполнителя.

Контекст анализа:
- Тип договора / Занятость: {data.doc_type}
- Юрисдикция: {data.jurisdiction}

Текст договора:
---
{contract_content[:15000]}
---

Сформируй разбор строго в формате JSON:
1. `overall_risk_level`: Оцени общий уровень риска (строго одно из значений: 'Низкий', 'Средний', 'Высокий').
2. `summary`: Краткий вывод по договору в 1-2 предложениях.
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
        
        # Сохранение лога анализа в базу данных PostgreSQL
        if SessionLocal:
            try:
                db = SessionLocal()
                db_log = AnalysisLogDB(
                    doc_type=data.doc_type,
                    jurisdiction=data.jurisdiction,
                    risk_level="Анализ выполнен"
                )
                db.add(db_log)
                db.commit()
                db.close()
            except Exception as db_err:
                print(f"Ошибка сохранения в БД: {db_err}")

        # Возвращаем распарсенный JSON напрямую
        return {"status": "ok", "result": response.text}
    except Exception as e:
        print(f"Gemini API error: {e}")
        raise HTTPException(status_code=500, detail=f"Ошибка анализа через Gemini API: {str(e)}")
