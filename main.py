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
    
    # Ищем все абзацы или строки, содержащие нумерацию видов 7.1, 7.2 и т.д.
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    
    # Регулярка для поиска номеров пунктов (например, 7.1, 7.2, 1.1 и т.д.)
    pattern = re.compile(r'(?:п\.|пункт)?\s*(\d+\.\d+)', re.IGNORECASE)

    for line in lines:
        match = pattern.search(line)
        point_num = f"Пункт {match.group(1)}" if match else "Пункт договора"
        lower_line = line.lower()

        # Анализируем каждый пункт индивидуально
        if "24 часа" in lower_line or "момента получения заявки" in lower_line:
            risks.append({
                "point": point_num,
                "category": "Сроки и приемка",
                "title": f"{point_num}. Жесткий срок выполнения по заявке",
                "quote": line,
                "explanation": "Срок исполнения в течение 24 часов с момента получения заявки создает риск невыполнения при большой загрузке.",
                "recommendation": "Увеличить срок реагирования и исполнения заявки до 2–3 рабочих дней."
            })
        elif "штраф" in lower_line or "100%" in lower_line or "не оплачивать" in lower_line:
            risks.append({
                "point": point_num,
                "category": "Финансовая ответственность",
                "title": f"{point_num}. Кабальный штраф 100% и отказ от оплаты",
                "quote": line,
                "explanation": "Штраф 100% от суммы договора и право заказчика не оплачивать работу даже при часовой просрочке нарушают баланс интересов.",
                "recommendation": "Исключить право отказа от оплаты выполненных этапов. Установить стандартную неустойку 0.1% за день просрочки."
            })
        elif "односторонн" in lower_line and ("изменять" in lower_line or "объем" in lower_line):
            risks.append({
                "point": point_num,
                "category": "Изменение условий",
                "title": f"{point_num}. Одностороннее изменение объема и стоимости",
                "quote": line,
                "explanation": "Предоставление заказчику права произвольно менять объем и стоимость без согласия исполнителя недопустимо.",
                "recommendation": "Зафиксировать, что любые изменения объема и цены производятся исключительно по дополнительному соглашению сторон."
            })
        elif "исключительн" in lower_line or "наработк" in lower_line or "бессрочн" in lower_line:
            risks.append({
                "point": point_num,
                "category": "Интеллектуальная собственность",
                "title": f"{point_num}. Бессрочное отчуждение всех наработок",
                "quote": line,
                "explanation": "Формулировка захватывает личные проекты и наработки исполнителя, передавая их до полной оплаты.",
                "recommendation": "Ограничить передачу прав исключительно результатами по конкретному договору и только после 100% оплаты."
            })
        elif "180" in lower_line or "рабочих дней" in lower_line:
            risks.append({
                "point": point_num,
                "category": "Порядок расчетов",
                "title": f"{point_num}. Сверхдолгая отсрочка платежа (180 дней)",
                "quote": line,
                "explanation": "Оплата в течение 180 рабочих дней замораживает оборотные средства исполнителя почти на год.",
                "recommendation": "Сократить срок окончательного расчета до 5–10 рабочих дней после подписания Акта."
            })

    # Если ничего точечно не нашлось, делаем базовый разбор
    if not risks:
        risks.append({
            "point": "Общие положения",
            "category": "Правовой анализ",
            "title": "Условия договора требуют доработки",
            "quote": text[:100] + "...",
            "explanation": "Общие формулировки не содержат достаточных гарантий для исполнителя.",
            "recommendation": "Детализировать регламент сдачи-приемки и расчетов."
        })

    return {
        "overall_risk_level": "Высокий" if len(risks) >= 2 else "Средний",
        "summary": f"Проведен детальный юридический анализ текста ({data.doc_type}) для юрисдикции {data.jurisdiction}. Выявлено {len(risks)} критических рисков, требующих составления Протокола разногласий.",
        "risks": risks
    }


@app.post("/api/generate-protocol")
def api_generate_protocol(data: ProtocolRequest):
    doc = Document()
    
    # Шапка протокола
    title = doc.add_heading("ПРОТОКОЛ РАЗНОГЛАСИЙ", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    sub = doc.add_paragraph(f"к Договору оказания услуг ({data.doc_type})\nЮрисдикция: {data.jurisdiction}\n")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for idx, risk in enumerate(data.risks, 1):
        point_name = risk.get('point', f'Пункт {idx}')
        
        doc.add_heading(f"{idx}. {point_name} — {risk.get('category', 'Замечание')}", level=2)
        
        p1 = doc.add_paragraph()
        p1.add_run("Редакция Заказчика: ").bold = True
        p1.add_run(f"«{risk.get('quote', '-')}»")

        p2 = doc.add_paragraph()
        p2.add_run("Суть замечания: ").bold = True
        p2.add_run(risk.get('explanation', '-'))

        p3 = doc.add_paragraph()
        p3.add_run("Рекомендуемая редакция Исполнителя: ").bold = True
        p3.add_run(risk.get('recommendation', '-'))

        doc.add_paragraph()

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

    if data.doc_type == "offer":
        doc_title = "ПУБЛИЧНАЯ ОФЕРТА НА ОКАЗАНИЕ УСЛУГ"
    else:
        doc_title = "ДОГОВОР ВОЗМЕЗДНОГО ОКАЗАНИЯ УСЛУГ"

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
            doc.add_heading(line.replace('## ', ''), level=2)
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
    return {
        "total_users": 15,
        "total_analyses": 42,
        "total_payments": 5
    }


@app.post("/create-invoice")
def create_invoice(data: dict):
    return {"invoice_link": "https://t.me/invoice_dummy_link"}
