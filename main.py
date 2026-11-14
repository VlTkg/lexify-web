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
from docx.oxml import OxmlElement
from docx.oxml.ns import nsdecls, qn
import pypdf

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")

gemini_api_key = os.environ.get("GEMINI_API_KEY")
client = genai.Client(api_key=gemini_api_key) if gemini_api_key else None

@app.get("/")
def read_root():
    with open("static/index.html", "r", encoding="utf-8") as f:
        html_content = f.read()
    return HTMLResponse(
        content=html_content,
        headers={
            "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0",
            "Pragma": "no-cache",
            "Expires": "0"
        }
    )

def run_gemini_analysis(text: str, jurisdiction: str) -> str:
    if not client:
        return json.dumps({
            "overall_risk_level": "Средний",
            "summary": "AI клиент не инициализирован (отсутствует GEMINI_API_KEY).",
            "risks": []
        }, ensure_ascii=False)
    
    prompt = f"""
    Проанализируй следующий текст договора/статьи с точки зрения законодательства юрисдикции: {jurisdiction}.
    Выяви скрытые риски, кабальные условия, штрафы и невыгодные пункты.
    Ответ верни СТРОГО в формате JSON без каких-либо обрамлений markdown, со следующей структурой:
    {{
        "overall_risk_level": "Высокий" / "Средний" / "Низкий",
        "summary": "Краткое резюме анализа текста",
        "risks": [
            {{
                "category": "red" (или "yellow" / "green"),
                "title": "Название риска",
                "quote": "Цитата из текста или пункт",
                "explanation": "Суть риска",
                "recommendation": "Рекомендация по изменению формулировки"
            }}
        ]
    }}
    Текст для анализа:
    {text}
    """
    try:
        response = client.models.generate_content(
            model='gemini-3.8-flash',
            contents=prompt
        )
        res_text = response.text.strip()
        if res_text.startswith("```json"):
            res_text = res_text[7:]
        if res_text.endswith("```"):
            res_text = res_text[:-3]
        return res_text.strip()
    except Exception as e:
        return json.dumps({
            "overall_risk_level": "Высокий",
            "summary": f"Ошибка обращения к модели ИИ: {str(e)}",
            "risks": []
        }, ensure_ascii=False)

@app.post("/api/analyze")
async def analyze_text(data: dict):
    try:
        text = data.get("text", "")
        jurisdiction = data.get("jurisdiction", "Россия (РФ)") or "Россия (РФ)"
        analysis = run_gemini_analysis(text, jurisdiction)
        return {"result": analysis}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/analyze-file")
async def analyze_file(
    file: UploadFile = File(...), 
    jurisdiction: Optional[str] = Form("Россия (РФ)"), 
    telegram_id: Optional[int] = Form(8804083078)
):
    try:
        contents = await file.read()
        filename = file.filename.lower()
        extracted_text = ""

        if filename.endswith(".docx"):
            doc = docx.Document(io.BytesIO(contents))
            extracted_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
        elif filename.endswith(".pdf"):
            reader = pypdf.PdfReader(io.BytesIO(contents))
            extracted_text = "\n".join([page.extract_text() for page in reader.pages if page and page.extract_text()])
        else:
            extracted_text = contents.decode("utf-8", errors="ignore")

        if not extracted_text.strip():
            return {"result": json.dumps({
                "overall_risk_level": "Средний",
                "summary": "Не удалось извлечь текст из файла. Убедитесь, что файл содержит текстовый слой.",
                "risks": []
            }, ensure_ascii=False)}

        jur = jurisdiction if jurisdiction else "Россия (РФ)"
        analysis_result = run_gemini_analysis(extracted_text, jur)
        return {"result": analysis_result}
    except Exception as e:
        return {"result": json.dumps({
            "overall_risk_level": " Высокий",
            "summary": f"Ошибка обработки файла: {str(e)}",
            "risks": []
        }, ensure_ascii=False)}

class ProtocolRequest(BaseModel):
    doc_type: Optional[str] = "Договор услуг"
    jurisdiction: Optional[str] = "Россия (РФ)"
    executor: Optional[str] = "Иванов И.И."
    client: Optional[str] = "ТОО «Альфа Технологии»"
    description: Optional[str] = "Разработка модуля"
    payment: Optional[str] = "100% предоплата"
    deadline: Optional[str] = "30 дней"
    revisions: Optional[str] = "2"
    telegram_id: Optional[int] = 0
    risks: Optional[List[Any]] = []

@app.post("/generate-protocol")
async def generate_protocol(req: ProtocolRequest):
    try:
        doc = docx.Document()
        
        sections = doc.sections
        for section in sections:
            section.top_margin = Inches(1)
            section.bottom_margin = Inches(1)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(f"ДОКУМЕНТ / ОТЧЕТ: {req.doc_type}")
        run.bold = True
        run.font.size = Pt(14)
        run.font.color.rgb = RGBColor(30, 58, 138)

        p_meta = doc.add_paragraph()
        p_meta.add_run(f"Юрисдикция: {req.jurisdiction or 'Россия (РФ)'}\n").bold = True
        p_meta.add_run(f"Исполнитель: {req.executor} | Заказчик: {req.client}\n")
        p_meta.add_run(f"Дата формирования: {datetime.now().strftime('%d.%m.%Y')}")

        doc.add_paragraph()

        if req.description and not req.risks:
            p_desc = doc.add_paragraph()
            p_desc.add_run(f"Описание / Резюме:\n{req.description}")

        if req.risks and len(req.risks) > 0:
            p_head = doc.add_paragraph()
            r_head = p_head.add_run("Выявленные риски и рекомендации:")
            r_head.bold = True
            r_head.font.size = Pt(12)

            for idx, risk in enumerate(req.risks, 1):
                p_r = doc.add_paragraph()
                if isinstance(risk, dict):
                    title = risk.get('title', 'Замечание')
                    quote = risk.get('quote', '')
                    expl = risk.get('explanation', '')
                    rec = risk.get('recommendation', '')
                    p_r.add_run(f"Риск #{idx}: {title}\n").bold = True
                    if quote:
                        p_r.add_run(f"• Цитата: «{quote}»\n")
                    if expl:
                        p_r.add_run(f"• Суть: {expl}\n")
                    if rec:
                        p_r.add_run(f"• Рекомендация: {rec}\n")
                else:
                    p_r.add_run(f"• {str(risk)}")
        elif not req.description:
            p_body = doc.add_paragraph()
            p_body.add_run(f"Предмет договора / описание: {req.description}\n")
            p_body.add_run(f"Условия оплаты: {req.payment}\n")
            p_body.add_run(f"Сроки выполнения: {req.deadline}\n")
            p_body.add_run("Документ сформирован автоматически с помощью Lexify AI.")

        file_stream = io.BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)

        filename = "Lexify_Document.docx"
        doc_type_lower = (req.doc_type or "").lower()
        if "протокол" in doc_type_lower:
            filename = "Lexify_Protokol_Raznoglasiy.docx"
        elif "анализ" in doc_type_lower or "результат" in doc_type_lower:
            filename = "Lexify_Rezultat_Analiza.docx"

        return StreamingResponse(
            file_stream,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/generate-document")
async def generate_document(req: ProtocolRequest):
    try:
        # Теперь конструктор документов использует ту же надежную генерацию Word-файла
        return await generate_protocol(req)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/admin/stats/{telegram_id}")
async def admin_stats(telegram_id: int):
    return {
        "total_users": 1,
        "total_analyses": 12,
        "total_generations": 5,
        "total_payments": 0
    }
