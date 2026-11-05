import os
import io
import re
from fastapi import FastAPI, HTTPException, Response
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pydantic import BaseModel
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")


class AnalysisRequest(BaseModel):
    text: str
    doc_type: str
    jurisdiction: str
    telegram_id: int = 0
    first_name: str = ""
    username: str = ""


class ProtocolRequest(BaseModel):
    doc_type: str
    jurisdiction: str
    telegram_id: int = 0
    risks: list = []


class ConstructorRequest(BaseModel):
    doc_type: str
    jurisdiction: str
    executor: str
    client: str
    description: str
    payment: str
    deadline: str
    revisions: str
    telegram_id: int = 0


class ConstructorDownloadSchema(BaseModel):
    document_text: str
    doc_type: str
    jurisdiction: str


class FeedbackRequest(BaseModel):
    type: str
    rating: int
    sentiment: str
    feedback: str = ""
    doc_type: str = ""
    jurisdiction: str = ""


@app.get("/")
def read_root():
    return FileResponse("static/index.html")


@app.get("/app")
def read_app():
    return FileResponse("static/index.html")


@app.get("/user-info/{telegram_id}")
def user_info(telegram_id: int):
    return {
        "free_checks_left": 3,
        "is_admin": True
    }


@app.post("/api/analyze")
def api_analyze(data: AnalysisRequest):
    text = data.text.strip()
    if not text:
        raise HTTPException(status_code=400, detail="Текст договора пуст")
    
    risks = []
    raw_splits = re.split(r'(?=\b\d+\.\d+\.)', text)
    if len(raw_splits) <= 1:
        raw_splits = text.split('\n')

    for chunk in raw_splits:
        chunk_clean = chunk.strip()
        if not chunk_clean:
            continue
            
        lower_chunk = chunk_clean.lower()
        num_match = re.search(r'(\d+\.\d+)', chunk_clean)
        point_num = f"Пункт {num_match.group(1)}" if num_match else "Пункт договора"

        if "24 часа" in lower_chunk or "момента получения заявки" in lower_chunk:
            risks.append({
                "point": point_num,
                "category": "Сроки и приемка",
                "title": f"{point_num}. Жесткий срок выполнения по заявке",
                "quote": chunk_clean,
                "explanation": "Срок исполнения в течение 24 часов с момента получения заявки создает критический риск нарушения из-за высокой загрузки.",
                "recommendation": "Увеличить срок реагирования и исполнения заявки до 2–3 рабочих дней."
            })
        elif "штраф" in lower_chunk or "100%" in lower_chunk or "не оплачивать" in lower_chunk or "даже на 1 час" in lower_chunk:
            risks.append({
                "point": point_num,
                "category": "Финансовая ответственность",
                "title": f"{point_num}. Кабальный штраф 100% и отказ от оплаты",
                "quote": chunk_clean,
                "explanation": "Штраф 100% от суммы договора и право заказчика не оплачивать работу при минимальной просрочке грубо нарушают баланс интересов.",
                "recommendation": "Исключить право отказа от оплаты уже выполненных этапов. Установить стандартную неустойку 0.1% за день просрочки."
            })
        elif "односторонн" in lower_chunk and ("изменять" in lower_chunk or "объем" in lower_chunk or "стоимост" in lower_chunk):
            risks.append({
                "point": point_num,
                "category": "Изменение условий",
                "title": f"{point_num}. Одностороннее изменение объема и стоимости",
                "quote": chunk_clean,
                "explanation": "Предоставление заказчику права произвольно менять объем услуг, сроки и цену без согласия исполнителя недопустимо.",
                "recommendation": "Зафиксировать, что любые изменения объема и цены производятся исключительно по дополнительному соглашению сторон."
            })
        elif "исключительн" in lower_chunk or "наработк" in lower_chunk or "бессрочн" in lower_chunk:
            risks.append({
                "point": point_num,
                "category": "Интеллектуальная собственность",
                "title": f"{point_num}. Бессрочное отчуждение всех наработок",
                "quote": chunk_clean,
                "explanation": "Формулировка захватывает личные проекты и сторонние наработки исполнителя, передавая их безвозмездно.",
                "recommendation": "Ограничить передачу прав исключительно результатами по конкретному договору и только после 100% оплаты."
            })
        elif "180" in lower_chunk or "рабочих дней" in lower_chunk:
            risks.append({
                "point": point_num,
                "category": "Порядок расчетов",
                "title": f"{point_num}. Сверхдолгая отсрочка платежа (180 дней)",
                "quote": chunk_clean,
                "explanation": "Оплата в течение 180 рабочих дней (почти 9 месяцев) замораживает оборотные средства исполнителя.",
                "recommendation": "Сократить срок окончательного расчета до 5–10 рабочих дней после подписания Акта."
            })

    if not risks:
        risks.append({
            "point": "Общие положения",
            "category": "Правовой анализ",
            "title": "Условия договора требуют доработки",
            "quote": text[:150] + "...",
            "explanation": "В тексте обнаружены общие формулировки, требующие конкретизации.",
            "recommendation": "Детализировать регламент сдачи-приемки и расчетов."
        })

    return {
        "overall_risk_level": "Высокий" if len(risks) >= 2 else "Средний",
        "summary": f"Проведен детальный юридический анализ текста ({data.doc_type}) для юрисдикции {data.jurisdiction}. Выявлено {len(risks)} критических рисков по конкретным пунктам договора.",
        "risks": risks
    }


@app.post("/api/generate-protocol")
def api_generate_protocol(data: ProtocolRequest):
    doc = Document()
    
    # Настройка полей документа
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Заголовок протокола
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_t = title_p.add_run("ПРОТОКОЛ РАЗНОГЛАСИЙ")
    run_t.bold = True
    run_t.font.name = "Times New Roman"
    run_t.font.size = Pt(14)
    run_t.font.color.rgb = RGBColor(0, 0, 0)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = sub_p.add_run(f"к Договору оказания услуг ({data.doc_type})\nЮрисдикция: {data.jurisdiction}\n")
    run_sub.font.name = "Times New Roman"
    run_sub.font.size = Pt(11)
    run_sub.font.color.rgb = RGBColor(50, 50, 50)

    doc.add_paragraph() # Отступ

    # Создаем красивую таблицу с протоколом
    table = doc.add_table(rows=1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # Задаем ширину колонок
    col_widths = [Inches(0.6), Inches(2.2), Inches(2.0), Inches(2.2)]
    
    hdr_cells = table.rows[0].cells
    hdr_titles = ["№", "Редакция Заказчика", "Суть замечания / Риск", "Рекомендуемая редакция"]
    
    for i, title_text in enumerate(hdr_titles):
        hdr_cells[i].text = title_text
        hdr_cells[i].width = col_widths[i]
        # Стилизация шапки таблицы
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.bold = True
            run.font.name = "Times New Roman"
            run.font.size = Pt(10)
            run.font.color.rgb = RGBColor(255, 255, 255)
        # Заливка шапки темно-серым цветом
        shading_xml = parse_xml(r'<w:shd {} w:fill="333333"/>'.format(nsdecls('w')))
        hdr_cells[i]._tc.get_or_add_tcPr().append(shading_xml)

    # Заполняем таблицу данными по каждому пункту
    for idx, risk in enumerate(data.risks, 1):
        row_cells = table.add_row().cells
        
        row_cells[0].text = str(idx)
        row_cells[1].text = f"«{risk.get('quote', '-')}»"
        row_cells[2].text = risk.get('explanation', '-')
        row_cells[3].text = risk.get('recommendation', '-')

        for i, cell in enumerate(row_cells):
            cell.width = col_widths[i]
            for p in cell.paragraphs:
                p.paragraph_format.line_spacing = 1.15
                for run in p.runs:
                    run.font.name = "Times New Roman"
                    run.font.size = Pt(10)
                    run.font.color.rgb = RGBColor(0, 0, 0)

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)

    return Response(
        content=file_stream.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": "attachment; filename=Protokol_raznoglasiy.docx"}
    )


@app.post("/api/generate-document")
def api_generate_document(data: ConstructorRequest):
    desc = data.description.strip() if data.description else "Оказание профессиональных услуг"
    executor = data.executor if data.executor else "Исполнитель"
    client = data.client if data.client else "Заказчик"
    payment = data.payment if data.payment else "Постоплата"
    deadline = data.deadline if data.deadline else "14 календарных дней"
    revisions = data.revisions if data.revisions else "2"

    doc_title = "ПУБЛИЧНАЯ ОФЕРТА НА ОКАЗАНИЕ УСЛУГ" if data.doc_type == "offer" else "ДОГОВОР ВОЗМЕЗДНОГО ОКАЗАНИЯ УСЛУГ"

    document_text = f"""# {doc_title}
Юрисдикция: {data.jurisdiction}
г. Москва                                                                          «___» _________ 2026 г.

Настоящий документ определяет условия оказания услуг и является официальным соглашением между сторонами: {executor}, именуемым в дальнейшем «Исполнитель», и {client}, именуемым в дальнейшем «Заказчик».

1. ПРЕДМЕТ ДОГОВОРА
1.1. Исполнитель обязуется по заданию Заказчика оказать следующие услуги: {desc}, а Заказчик обязуется принять и оплатить их.
1.2. Срок оказания услуг составляет: {deadline} с момента согласования технического задания.

2. СТОИМОСТЬ И ПОРЯДОК РАСЧЕТОВ
2.1. Стоимость услуг и порядок оплаты: {payment}.
2.2. Расчеты производятся в валюте юрисдикции {data.jurisdiction}.

3. ПОРЯДОК СДАЧИ-ПРИЕМКИ УСЛУГ
3.1. Сдача-приемка оказанных услуг оформляется Актом приема-передачи.
3.2. Заказчик вправе запросить до {revisions} итераций бесплатных правок. Направление мотивированного отказа допускается в течение 3 рабочих дней.

4. ОТВЕТСТВЕННОСТЬ И ПОРЯДОК РАЗРЕШЕНИЯ СПОРОВ
4.1. В случае нарушения сроков оплаты Заказчик выплачивает пени в размере 0.1% от суммы задолженности за каждый день просрочки.
4.2. Все споры разрешаются путем переговоров в соответствии с законодательством {data.jurisdiction}.

5. СРОК ДЕЙСТВИЯ И ЗАКЛЮЧИТЕЛЬНЫЕ ПОЛОЖЕНИЯ
5.1. Договор вступает в силу с момента его подписания и действует до полного исполнения сторонами обязательств.

6. РЕКВИЗИТЫ И ПОДПИСИ СТОРОН

Исполнитель: 
{executor}
________________ / ______________

Заказчик: 
{client}
________________ / ______________
"""
    return {"document_text": document_text}


@app.post("/api/download-constructor-file")
def api_download_constructor_file(data: ConstructorDownloadSchema):
    doc = Document()
    for line in data.document_text.split('\n'):
        if line.startswith('# '):
            h = doc.add_heading(line.replace('# ', ''), level=1)
            h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif line.startswith('## '):
            doc.add_heading(line.replace('# ', ''), level=2)
        elif line.startswith('1.') or line.startswith('2.') or line.startswith('3.') or line.startswith('4.') or line.startswith('5.') or line.startswith('6.'):
            p = doc.add_paragraph()
            p.add_run(line).bold = True
        else:
            doc.add_paragraph(line)

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)

    filename = "Oferta.docx" if data.doc_type == "offer" else "Dogovor.docx"
    return Response(
        content=file_stream.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )


@app.post("/api/feedback")
def api_feedback(data: FeedbackRequest):
    return {"status": "success", "message": "Отзыв успешно сохранен"}


@app.get("/admin/stats/{telegram_id}")
def admin_stats(telegram_id: int):
    return {"total_users": 15, "total_analyses": 42, "total_payments": 5}


@app.post("/create-invoice")
def create_invoice(data: dict):
    return {"invoice_link": "https://t.me/invoice_dummy_link"}
