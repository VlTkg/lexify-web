import os
import io
import uuid
import json
import zipfile
import xml.etree.ElementTree as ET
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
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
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
    Вы профессиональный юрист. Проведите глубокий экспертный правовой анализ текста в юрисдикции: {req.jurisdiction}.
    Обязательно определите и выделите предмет договора, а также укажите конкретные статьи применимого законодательства (например, ГК РФ/РК/РБ), которые регулируют выявленные риски или нарушения.
    
    Верните ответ СТРОГО в формате JSON со следующей структурой (без маркдаун-оберток вроде ```json):
    {{
      "overall_risk_level": "Высокий" или "Средний" или "Низкий",
      "summary": "Четкое резюме анализа договора с описанием предмета и ключевых правовых выводов",
      "risks": [
        {{
          "category": "red" или "yellow" или "green",
          "title": "Название риска с указанием статьи закона",
          "quote": "Цитата из текста договора",
          "explanation": "Суть и опасность риска со ссылкой на статьи законов",
          "recommendation": "Как исправить пункт"
        }}
      ]
    }}
    
    Текст для анализа:
    {req.text}
    """
    try:
        response = client.models.generate_content(
            model="gemini-3.8-flash",
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
    filename_lower = file.filename.lower()
    
    try:
        if filename_lower.endswith(".docx"):
            doc = docx.Document(io.BytesIO(content))
            extracted_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
        elif filename_lower.endswith(".pdf"):
            reader = pypdf.PdfReader(io.BytesIO(content))
            extracted_text = "\n".join([page.extract_text() for page in reader.pages if page and page.extract_text()])
        elif filename_lower.endswith(".txt"):
            decoded = False
            for enc in ["utf-8", "cp1251", "windows-1251", "iso-8859-1", "koi8-r"]:
                try:
                    extracted_text = content.decode(enc)
                    decoded = True
                    break
                except UnicodeDecodeError:
                    continue
            if not decoded:
                extracted_text = content.decode("utf-8", errors="ignore")
            
            # Очистка нечитаемых символов/знаков вопроса для файлов с поврежденной кодировкой
            if "" in extracted_text or extracted_text.count("?") > 20:
                extracted_text = "[Внимание: исходный файл имеет поврежденную кодировку текста].\n" + extracted_text
        elif filename_lower.endswith(".odt"):
            with zipfile.ZipFile(io.BytesIO(content)) as z:
                if "content.xml" in z.namelist():
                    xml_content = z.read("content.xml")
                    root = ET.fromstring(xml_content)
                    texts = []
                    for elem in root.iter():
                        if elem.text:
                            texts.append(elem.text)
                    extracted_text = "\n".join(texts)
        else:
            raise HTTPException(status_code=400, detail="Поддерживаются только .docx, .pdf, .txt и .odt файлы")
    except Exception as e:
        if isinstance(e, HTTPException):
            raise e
        raise HTTPException(status_code=400, detail=f"Ошибка чтения файла: {str(e)}")

    if not extracted_text.strip():
        raise HTTPException(status_code=400, detail="Не удалось извлечь текст из файла или он пуст")

    req = AnalyzeRequest(text=extracted_text, jurisdiction=jurisdiction, telegram_id=telegram_id)
    res = await analyze_contract(req)
    if isinstance(res, dict) and "result" in res:
        return res
    return {"result": res}

@app.post("/generate-protocol")
async def generate_protocol(req: GenerateRequest, request: Request):
    doc = docx.Document()
    
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)
    font.color.rgb = RGBColor(38, 38, 38)

    if req.doc_type == "Результат анализа":
        title = doc.add_paragraph()
        title.paragraph_format.space_after = Pt(2)
        trun = title.add_run("ЭКСПЕРТНОЕ ЗАКЛЮЧЕНИЕ ПО РЕЗУЛЬТАТАМ АНАЛИЗА РИСКОВ")
        trun.bold = True
        trun.font.size = Pt(14)
        trun.font.color.rgb = RGBColor(20, 20, 20)

        meta = doc.add_paragraph()
        meta.paragraph_format.space_after = Pt(12)
        meta.add_run(f"Применимое право: {req.jurisdiction}\n").bold = True
        meta.add_run(f"Дата формирования: {datetime.now().strftime('%d.%m.%Y %H:%M')}")

        div = doc.add_paragraph()
        div.paragraph_format.space_after = Pt(12)
        drun = div.add_run("--------------------------------------------------------------------------------------------------")
        drun.font.color.rgb = RGBColor(120, 120, 120)

        sum_p = doc.add_paragraph()
        sum_p.paragraph_format.space_after = Pt(12)
        sum_p.add_run("1. ОБЩЕЕ ПРАВОВОЕ РЕЗЮМЕ И ПРЕДМЕТ АНАЛИЗА\n").bold = True
        sum_p.add_run(req.description or "Анализ рисков договора выполнен с учетом требований применимого законодательства.")

        head_p = doc.add_paragraph()
        head_p.paragraph_format.space_after = Pt(8)
        hrun = head_p.add_run("2. ДЕТАЛЬНЫЙ АНАЛИЗ ПУНКТОВ, СТАТЬИ ЗАКОНОВ И ВЫЯВЛЕННЫЕ РИСКИ")
        hrun.bold = True
        hrun.font.size = Pt(12)

        if req.risks and isinstance(req.risks, list):
            for idx, r in enumerate(req.risks, 1):
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(10)
                
                cat = r.get('category', 'yellow')
                lvl_str = "Высокий уровень риска" if cat == 'red' else ("Средний уровень риска" if cat == 'yellow' else "Низкий уровень риска")
                
                p.add_run(f"Пункт {idx}. {r.get('title', 'Замечание')} [{lvl_str}]\n").bold = True
                
                if r.get('quote'):
                    q = p.add_run(f"Реальная редакция: «{r.get('quote')}»\n")
                    q.italic = True
                    q.font.color.rgb = RGBColor(80, 80, 80)
                
                p.add_run("Правовая суть и статьи законов: ").bold = True
                p.add_run(f"{r.get('explanation', '')}\n")
                
                p.add_run("Рекомендуемое решение: ").bold = True
                p.add_run(f"{r.get('recommendation', '')}\n")
        else:
            doc.add_paragraph("Существенных правовых замечаний не зафиксировано.")

    elif req.doc_type == "Протокол разногласий":
        title = doc.add_paragraph()
        title.paragraph_format.space_after = Pt(2)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        trun = title.add_run("ПРОТОКОЛ РАЗНОГЛАСИЙ")
        trun.bold = True
        trun.font.size = Pt(14)
        trun.font.color.rgb = RGBColor(20, 20, 20)

        sub = doc.add_paragraph()
        sub.paragraph_format.space_after = Pt(12)
        sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sub.add_run(f"к проекту договора от {datetime.now().strftime('%d.%m.%Y')}")

        meta = doc.add_paragraph()
        meta.paragraph_format.space_after = Pt(10)
        meta.add_run(f"Юрисдикция: {req.jurisdiction}\n")

        table = doc.add_table(rows=1, cols=3)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = "Редакция по проекту договора"
        hdr_cells[1].text = "Предлагаемая редакция"
        hdr_cells[2].text = "Обоснование разногласий (со ссылками на статьи)"

        col_widths = [Inches(2.0), Inches(2.0), Inches(2.5)]
        for row in table.rows:
            for i, w in enumerate(col_widths):
                row.cells[i].width = w

        for cell in hdr_cells:
            shading = parse_xml(r'<w:shd {} w:fill="E6E6E6"/>'.format(nsdecls('w')))
            cell._tc.get_or_add_tcPr().append(shading)
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in p.runs:
                    run.bold = True
                    run.font.size = Pt(9.5)

        if req.risks and isinstance(req.risks, list):
            for r in req.risks:
                row_cells = table.add_row().cells
                for i, w in enumerate(col_widths):
                    row_cells[i].width = w
                
                row_cells[0].text = r.get('quote', 'Текст пункта отсутствует')
                row_cells[1].text = r.get('recommendation', 'Изложить в редакции стороны')
                row_cells[2].text = f"{r.get('title', 'Риск')}: {r.get('explanation', '')}"

                for cell in row_cells:
                    for p in cell.paragraphs:
                        for run in p.runs:
                            run.font.size = Pt(9.5)
                            run.font.color.rgb = RGBColor(40, 40, 40)
        else:
            row_cells = table.add_row().cells
            row_cells[0].text = "Пункты не указаны"
            row_cells[1].text = "Без изменений"
            row_cells[2].text = "Замечаний нет"

        doc.add_paragraph().paragraph_format.space_after = Pt(20)
        
        sig = doc.add_paragraph()
        sig.add_run("ЗАКАЗЧИК:                              ИСПОЛНИТЕЛЬ:\n\n")
        sig.add_run("_________________                      _________________")

    else:
        title = doc.add_paragraph()
        title.paragraph_format.space_after = Pt(4)
        trun = title.add_run(f"ДОКУМЕНТ: {req.doc_type.upper()}")
        trun.bold = True
        trun.font.size = Pt(14)

        doc.add_paragraph(f"Юрисдикция: {req.jurisdiction}")
        doc.add_paragraph(f"Дата: {datetime.now().strftime('%d.%m.%Y')}")
        doc.add_paragraph("--------------------------------------------------------------------------------------------------")
        doc.add_paragraph(f"Исполнитель: {req.executor}")
        doc.add_paragraph(f"Заказчик: {req.client}")
        doc.add_paragraph(f"Предмет: {req.description}")
        doc.add_paragraph(f"Условия оплаты: {req.payment}")
        doc.add_paragraph(f"Срок выполнения: {req.deadline}")

    file_id = str(uuid.uuid4())
    safe_name = "Protocol_Raznoglasiy.docx" if req.doc_type == "Протокол разногласий" else ("Result_Analysis.docx" if req.doc_type == "Результат анализа" else "Document.docx")
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
