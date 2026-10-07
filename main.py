import os
import re
import io
import json
from typing import Dict, Any

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pydantic import BaseModel

import docx
from pypdf import PdfReader
from odf import text
from odf.opendocument import load as load_odt

from google import genai
from google.genai import types

app = FastAPI(title="Lexify Web API", version="1.0.0")

app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/")
async def serve_ui():
    return FileResponse("static/index.html")

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
Ты — Lexify, юридический аудитор фрилансера (разработчика, дизайнера, копирайтера).
Твоя цель: найти скрытые риски в договоре и дать формулировки для правок.

Оценивай риски:
1. RED: передача IP до полной оплаты, штраф > 0.1%/день, отказ от договора без оплаты, постоплата > 30 дней.
2. YELLOW: неопределенный порядок приемки, безлимитные правки, требование быть на связи в фиксированное время.
3. GREEN: предоплата/этапы, четкие сроки приемки (3-5 дней), ограничение ответственности.

Отвечай СТРОГО в формате JSON:
{
  "summary": {
    "overall_risk_level": "RED | YELLOW | GREEN",
    "score": 0-100,
    "verdict": "Краткий вывод"
  },
  "risks": [
    {
      "id": "risk-1",
      "category": "penalties | ip | payment | scope_and_deadlines",
      "severity": "RED | YELLOW | GREEN",
      "title": "Заголовок",
      "quote": "Цитата из договора",
      "explanation": "В чем риск",
      "recommendation": "Как исправить",
      "suggested_edit": "Формулировка правки"
    }
  ],
  "negotiation_message": "Готовый вежливый текст сообщения заказчику со всеми правками."
}
"""

class TextAnalysisRequest(BaseModel):
    text: str

@app.post("/api/v1/analyze-text")
async def analyze_text(request: TextAnalysisRequest) -> Dict[str, Any]:
    if not client:
        raise HTTPException(status_code=500, detail="GEMINI_API_KEY не настроен.")
    anonymized = anonymize_text(request.text)
    return run_gemini_analysis(anonymized)

@app.post("/api/v1/analyze-file")
async def analyze_file(file: UploadFile = File(...)) -> Dict[str, Any]:
    if not client:
        raise HTTPException(status_code=500, detail="GEMINI_API_KEY не настроен.")
    content = await file.read()
    raw_text = extract_text_from_file(content, file.filename)
    anonymized = anonymize_text(raw_text)
    return run_gemini_analysis(anonymized)

def run_gemini_analysis(contract_text: str) -> Dict[str, Any]:
    prompt = f"Проанализируй договор:\n\n{contract_text}"
    
    try:
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
            config=types.GenerateContentConfig(
                system_instruction=SYSTEM_PROMPT,
                response_mime_type="application/json",
                temperature=0.2,
            ),
        )
        return json.loads(response.text)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка AI API: {str(e)}")
