import os
import re
import io
import json
import time
from typing import Dict, Any, Optional

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pydantic import BaseModel

import docx
from pypdf import PdfReader
from odf import text
from odf.opendocument import load as load_odt

from google import genai
from google.genai import types as genai_types

app = FastAPI(title="LegalGuard Web API", version="1.0.0")

# Подключение статики
if os.path.exists("static"):
    app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/")
async def serve_ui():
    if os.path.exists("static/index.html"):
        return FileResponse("static/index.html")
    return {"status": "API is running. UI file static/index.html not found."}

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
client = genai.Client(api_key=GEMINI_API_KEY) if GEMINI_API_KEY else None


def extract_text_from_file(file_bytes: bytes, filename: str) -> str:
    ext = filename.split('.')[-1].lower()
    extracted_text = ""

    try:
        if ext == "txt":
            extracted_text = file_bytes.decode("utf-8", errors="ignore")

        elif ext == "docx":
            doc = docx.Document(io.BytesIO(file_bytes))
            paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
            extracted_text = "\n".join(paragraphs)

        elif ext == "pdf":
            reader = PdfReader(io.BytesIO(file_bytes))
            pages_text = []
            for page in reader.pages:
                t = page.extract_text()
                if t:
                    pages_text.append(t)
            extracted_text = "\n".join(pages_text)

        elif ext == "odt":
            odt_doc = load_odt(io.BytesIO(file_bytes))
            paragraphs = []
            for el in odt_doc.getElementsByType(text.P):
                t = "".join([node.data for node in el.childNodes if node.nodeType == 3])
                if t.strip():
                    paragraphs.append(t)
            extracted_text = "\n".join(paragraphs)

        else:
            raise HTTPException(
                status_code=400, 
                detail="Неподдерживаемый формат. Используйте: .docx, .pdf, .odt, .txt"
            )

    except Exception as e:
        if isinstance(e, HTTPException):
            raise e
        raise HTTPException(status_code=422, detail=f"Ошибка обработки файла: {str(e)}")

    if not extracted_text.strip():
        raise HTTPException(status_code=400, detail="Файл пуст или не содержит текста.")

    return extracted_text


def anonymize_text(raw_text: str) -> str:
    text_data = raw_text
    text_data = re.sub(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', '[EMAIL]', text_data)
    text_data = re.sub(r'(\+7|8)[\s\-]?\(?\d{3}\)?[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2}', '[PHONE]', text_data)
    text_data = re.sub(r'\b\d{2}\s?\d{2}\s?\d{6}\b', '[PASSPORT]', text_data)
    text_data = re.sub(r'\b\d{10}\b|\b\d{12}\b', '[TAX_ID]', text_data)
    text_data = re.sub(r'\b\d{4}[\s\-]?\d{4}[\s\-]?\d{4}[\s\-]?\d{4}\b', '[CARD_NUMBER]', text_data)
    text_data = re.sub(r'\b40[8780]1810\d{12}\b', '[BANK_ACCOUNT]', text_data)
    return text_data


SYSTEM_PROMPT = """
Ты — LegalGuard, юридический аудитор и защитник прав фрилансера.
Твоя цель: найти скрытые риски в договоре и дать готовые формулировки для правок.

Оценивай риски по уровням: RED (критический), YELLOW (средний), GREEN (безопасно).

Отвечай СТРОГО в формате JSON без каких-либо вводных слов:
{
  "summary": {
    "overall_risk_level": "RED",
    "score": 45,
    "verdict": "Договор содержит кабальные условия по штрафам и передаче прав."
  },
  "risks": [
    {
      "id": "risk-1",
      "severity": "RED",
      "title": "Односторонняя ответственность и штрафы",
      "quote": "Исполнитель уплачивает штраф 100% за любой срыв срока",
      "explanation": "Размер штрафа несоразмерен возможному ущербу.",
      "suggested_edit": "В случае просрочки Исполнитель уплачивает неустойку в размере 0.1% от стоимости этапа за каждый день просрочки."
    }
  ],
  "negotiation_message": "Здравствуйте! Проанализировал договор и предлагаю скорректировать пару пунктов..."
}
"""

class TextAnalysisRequest(BaseModel):
    text: str
    jurisdiction: Optional[str] = "RU"
    category: Optional[str] = "general"


def run_gemini_analysis(contract_text: str) -> Dict[str, Any]:
    if not GEMINI_API_KEY:
        raise HTTPException(status_code=500, detail="Переменная GEMINI_API_KEY не задана на сервере.")
    
    prompt = f"Проанализируй текст договора:\n\n{contract_text}"
    
    model_name = "gemini-3.8-flash"
    max_retries = 3
    last_error = None

    for attempt in range(max_retries):
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=prompt,
                config=genai_types.GenerateContentConfig(
                    system_instruction=SYSTEM_PROMPT,
                    response_mime_type="application/json",
                    temperature=0.2,
                ),
            )
            return json.loads(response.text)
        except Exception as e:
            last_error = str(e)
            if "503" in last_error or "UNAVAILABLE" in last_error:
                time.sleep(2 * (attempt + 1))
            else:
                break

    raise HTTPException(status_code=500, detail=f"Ошибка AI API: {last_error}")


# ---------------- ИНТЕРФЕЙСЫ ЭНДПОИНТОВ ---------------- #

@app.post("/api/v1/analyze-text")
async def analyze_text_v1(request: TextAnalysisRequest):
    anonymized = anonymize_text(request.text)
    return run_gemini_analysis(anonymized)


@app.post("/api/v1/analyze-file")
async def analyze_file_v1(file: UploadFile = File(...)):
    content = await file.read()
    raw_text = extract_text_from_file(content, file.filename)
    anonymized = anonymize_text(raw_text)
    return run_gemini_analysis(anonymized)


@app.post("/api/analyze")
async def analyze_universal(
    file: Optional[UploadFile] = File(None),
    text: Optional[str] = Form(None),
    jurisdiction: Optional[str] = Form("RU"),
    category: Optional[str] = Form("general")
):
    raw_text = ""
    if file and file.filename:
        content = await file.read()
        raw_text = extract_text_from_file(content, file.filename)
    elif text and text.strip():
        raw_text = text.strip()
    else:
        raise HTTPException(status_code=400, detail="Передайте файл или текст договора.")

    anonymized = anonymize_text(raw_text)
    return run_gemini_analysis(anonymized)


@app.post("/api/feedback")
async def handle_feedback(data: Dict[str, Any]):
    return {"status": "ok", "message": "Feedback received"}
