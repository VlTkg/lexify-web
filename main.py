import os
import io
import httpx
from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import Optional
from google import genai
from pypdf import PdfReader
import docx

app = FastAPI()

# Монтируем статические файлы
app.mount("/static", StaticFiles(directory="static"), name="static")

# Переменные окружения
FEEDBACK_WEBHOOK_URL = os.environ.get("GOOGLE_SHEET_WEBHOOK_URL", "")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")

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

def extract_text_from_file(file_bytes: bytes, filename: str) -> str:
    """Извлечение текста из файлов PDF, DOCX, TXT, ODT"""
    ext = filename.split(".")[-1].lower() if "." in filename else ""
    text = ""

    try:
        if ext == "pdf":
            reader = PdfReader(io.BytesIO(file_bytes))
            for page in reader.pages:
                extracted = page.extract_text()
                if extracted:
                    text += extracted + "\n"
        elif ext == "docx":
            doc = docx.Document(io.BytesIO(file_bytes))
            for para in doc.paragraphs:
                text += para.text + "\n"
        elif ext in ["txt", "odt"]:
            text = file_bytes.decode("utf-8", errors="ignore")
        else:
            text = file_bytes.decode("utf-8", errors="ignore")
    except Exception as e:
        print(f"File parsing error: {e}")

    return text.strip()

@app.post("/analyze")
async def analyze_contract(
    text: Optional[str] = Form(None),
    doc_type: str = Form("Общий"),
    jurisdiction: str = Form("Россия (РФ)"),
    file: Optional[UploadFile] = File(None)
):
    if not ai_client:
        raise HTTPException(status_code=500, detail="GEMINI_API_KEY не настроен на сервере.")

    contract_content = ""

    # Извлекаем текст из файла или берем из поля ввода
    if file:
        file_bytes = await file.read()
        contract_content = extract_text_from_file(file_bytes, file.filename)
    elif text:
        contract_content = text.strip()

    if not contract_content:
        raise HTTPException(status_code=400, detail="Текст договора или файл не предоставлен, либо файл пуст.")

    # Формируем промпт для юриста-эксперта
    prompt = f"""
Ты — профессиональный юрист-аналитик, специализирующийся на проверке договоров.
Проведи экспресс-анализ представленного договора и выяви потенциальные риски для пользователя.

Контекст анализа:
- Тип договора / Занятость: {doc_type}
- Юрисдикция: {jurisdiction}

Текст договора:
---
{contract_content[:15000]}
---

Сформируй четкий, структурированный ответ на русском языке по следующему плану:
1. **Общий уровень риска**: (Низкий / Средний / Высокий) + краткое обоснование в 1 предложении.
2. **Ключевые риски и «подводные камни»**: (3-5 главных рисков списком с пояснением, чем это грозит).
3. **Рекомендации по изменению**: (что именно стоит переписать или добавить в договор).

Пиши кратко, емко, доступным языком без лишней канцелярии.
"""

    try:
        response = ai_client.models.generate_content(
            model="gemini-2.5-flash",
            contents=prompt,
        )
        return {"status": "ok", "result": response.text}
    except Exception as e:
        print(f"Gemini API error: {e}")
        raise HTTPException(status_code=500, detail=f"Ошибка анализа через Gemini API: {str(e)}")
