import os
import io
import uuid
import json
from datetime import datetime
from fastapi import FastAPI, HTTPException, Request, UploadFile, File, Form
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from pydantic import BaseModel
from typing import Optional, List, Any
from google import genai
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pypdf

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")

gemini_api_key = os.environ.get("GEMINI_API_KEY")
client = genai.Client(api_key=gemini_api_key) if gemini_api_key else None

TEMP_FILES_DIR = "/tmp/lexify_docs"
os.makedirs(TEMP_FILES_DIR, exist_ok=True)

class AnalyzeRequest(BaseModel):
    text: str
    jurisdiction: Optional[str] = "Россия (РФ)"
    telegram_id: Optional[int] = 0

class GenerateRequest(BaseModel):
    doc_type: Optional[str] = "Договор услуг"
    jurisdiction: Optional[str] = "Россия (РФ)"
    executor: Optional[str] = "Иванов И.И."
    client: Optional[str] = "ТОО «Альфа Технологии»"
    description: Optional[str] = "Разработка программного модуля"
    payment: Optional[str] = "100% предоплата"
    deadline: Optional[str] = "30 дней"
    revisions: Optional[str] = "2"
    telegram_id: Optional[int] = 0
    risks: Optional[List[Any]] = []

@app.get("/", response_class=HTMLResponse)
async def read_index():
    with open("static/index.html", "r", encoding="utf-8") as f:
        return f.read()

@app.post("/api/analyze")
async def analyze_contract(req: AnalyzeRequest):
    if not client:
        return JSONResponse({"result": "Ошибка: GEMINI_API_KEY не настроен на сервере."})
    
    prompt = f"""
    Вы профессиональный юрист. Проведите экспертный анализ текста в юрисдикции: {req.jurisdiction}.
    Выявите скрытые риски, кабальные условия и потенциальные проблемы.
    
    Верните ответ СТРОГО в формате JSON со следующей структурой (без маркдаун-оберток вроде ```json):
    {{
      "overall_risk_level": "Высокий" или "Средний" или "Низкий",
      "summary": "Краткое резюме анализа договора",
      "risks": [
        {{
          "category": "red" или "yellow" или "green",
          "title": "Название риска",
          "quote": "Цитата из текста",
          "explanation": "Суть и опасность риска",
          "recommendation": "Как исправить пункт"
        }}
      ]
    }}
    
    Текст для анализа:
    {req.text}
    """
    try:
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=prompt,
        )
        clean_text = response.text.strip()
        if clean_text.startswith("```json"):
            clean_text = clean_text[7:]
        if clean_text.endswith("```"):
            clean_text = clean_text[:-3]
        
        try:
            parsed_json = json.loads(clean_text.strip())
            return {"result": parsed_json}
        except:
            return {"result": {"overall_risk_level": "Средний", "summary": response.text, "risks": []}}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/analyze-file")
async def analyze_file(
    file: UploadFile = File(...),
    jurisdiction: Optional[str] = Form("Россия (РФ)"),
    telegram_id: Optional[int] = Form(0)
):
    content = await file.read()
    extracted_text = ""
    
    if file.filename.endswith(".docx"):
        doc = docx.Document(io.BytesIO(content))
        extracted_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    elif file.filename.endswith(".pdf"):
        reader = pypdf.PdfReader(io.BytesIO(content))
        extracted_text = "\n".join([page.extract_text() for page in reader.pages if page.extract_text()])
    else:
        raise HTTPException(status_code=400, detail="Поддерживаются только .docx и .pdf файлы")

    if not extracted_text:
        raise HTTPException(status_code=400, detail="Не удалось извлечь текст из файла")

    req = AnalyzeRequest(text=extracted_text, jurisdiction=jurisdiction, telegram_id=telegram_id)
    return await analyze_contract(req)

@app.post("/generate-protocol")
async def generate_protocol(req: GenerateRequest, request: Request):
    doc = docx.Document()
    
    title = doc.add_paragraph()
    title_run = title.add_run(f"ДОКУМЕНТ: {req.doc_type.upper()}")
    title_run.bold = True
    title_run.font.size = Pt(14)
    title_run.font.color.rgb = RGBColor(31, 41, 55)
    
    doc.add_paragraph(f"Юрисдикция: {req.jurisdiction}")
    doc.add_paragraph(f"Дата формирования: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
    doc.add_paragraph("--------------------------------------------------")

    if req.doc_type in ["Результат анализа", "Протокол разногласий"]:
        doc.add_paragraph(f"Резюме / Описание: {req.description or 'Анализ правовых рисков выполнен.'}")
        doc.add_paragraph("\nВыявленные пункты и рекомендации:")
        
        if req.risks and isinstance(req.risks, list):
            for idx, r in enumerate(req.risks, 1):
                p = doc.add_paragraph()
                p.add_run(f"Пункт {idx}: {r.get('title', 'Риск')}\n").bold = True
                p.add_run(f"Цитата: «{r.get('quote', '')}»\n")
                p.add_run(f"Суть: {r.get('explanation', '')}\n")
                p.add_run(f"Рекомендация: {r.get('recommendation', '')}\n")
        else:
            doc.add_paragraph("Существенных критических замечаний не зафиксировано.")
    else:
        doc.add_paragraph(f"Исполнитель: {req.executor}")
        doc.add_paragraph(f"Заказчик: {req.client}")
        doc.add_paragraph(f"Предмет: {req.description}")
        doc.add_paragraph(f"Условия оплаты: {req.payment}")
        doc.add_paragraph(f"Срок выполнения: {req.deadline}")

    file_id = str(uuid.uuid4())
    safe_name = "Protocol.docx" if "Протокол" in req.doc_type else ("Result.docx" if "Результат" in req.doc_type else "Document.docx")
    filename = f"{file_id[:8]}_{safe_name}"
    file_path = os.path.join(TEMP_FILES_DIR, filename)
    
    doc.save(file_path)

    base_url = str(request.base_url).rstrip("/")
    download_url = f"{base_url}/api/download-temp/{file_id}"

    return {"download_url": download_url, "filename": safe_name}

@app.get("/api/download-temp/{file_id}")
async def download_temp_file(file_id: str):
    for filename in os.listdir(TEMP_FILES_DIR):
        if filename.startswith(file_id[:8]):
            file_path = os.path.join(TEMP_FILES_DIR, filename)
            if os.path.exists(file_path):
                original_name = filename.split("_", 1)[1] if "_" in filename else "Document.docx"
                return FileResponse(
                    file_path,
                    media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    filename=original_name
                )
    raise HTTPException(status_code=404, detail="Файл не найден или истек срок его хранения")

@app.get("/admin/stats/{telegram_id}")
async def admin_stats(telegram_id: int):
    return {
        "total_users": 1,
        "total_analyses": 12,
        "total_generations": 5,
        "total_payments": 0
    }
