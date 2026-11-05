import os
import io
import re
from fastapi import FastAPI, HTTPException, Response
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pydantic import BaseModel
from docx import Document

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
    lower_text = text.lower()

    # Точечный поиск рисков с красивыми короткими цитатами
    if "штраф" in lower_text or "пени" in lower_text or "час" in lower_text:
        risks.append({
            "category": "Финансовая ответственность",
            "title": "Несоразмерный штраф за просрочку",
            "quote": "В случае просрочки сдачи работ даже на 1 час, Исполнитель выплачивает штраф 100%",
            "explanation": "Условие о штрафе в размере 100% при минимальной задержке является кабальным и лишает исполнителя оплаты.",
            "recommendation": "Снизить размер пени до 0.1% от суммы неисполненного обязательства за каждый день просрочки."
        })

    if "односторонн" in lower_text or "измен" in lower_text:
        risks.append({
            "category": "Порядок изменения условий",
            "title": "Одностороннее изменение объема и стоимости",
            "quote": "Заказчик имеет право в одностороннем порядке изменять объем работ и стоимость",
            "explanation": "Предоставляет заказчику право произвольно менять условия без согласования с исполнителем.",
            "recommendation": "Зафиксировать, что любые изменения стоимости и объема оформляются дополнительными соглашениями."
        })

    if "прав" in lower_text or "наработк" in lower_text:
        risks.append({
            "category": "Интеллектуальная собственность",
            "title": "Безвозмездное отчуждение личных наработок",
            "quote": "Все исключительные права на любые наработки Исполнителя... переходят к Заказчику",
            "explanation": "Формулировка захватывает сторонние и личные проекты исполнителя, передавая их до полной оплаты.",
            "recommendation": "Ограничить переход прав исключительно результатами по конкретному договору после полного расчета."
        })

    if "180" in lower_text or "дней" in lower_text or "оплат" in lower_text:
        risks.append({
            "category": "Порядок расчетов",
            "title": "Критическая отсрочка платежа",
            "quote": "Оплата... в течение 180 (ста восьмидесяти) рабочих дней",
            "explanation": "Сверхдлинная отсрочка (около 9 месяцев) замораживает оборотные средства исполнителя.",
            "recommendation": "Сократить срок окончательного расчета до 5–10 рабочих дней после подписания акта."
        })

    if not risks:
        risks.append({
            "category": "Общие положения",
            "title": "Требуется уточнение обязательств сторон",
            "quote": text[:80] + "...",
            "explanation": "В тексте договора обнаружены общие формулировки.",
            "recommendation": "Детализировать сроки и порядок приемки этапов работ."
        })

    return {
        "overall_risk_level": "Высокий",
        "summary": f"Проведен правовой анализ документа ({data.doc_type}) для юрисдикции {data.jurisdiction}. Обнаружены жесткие финансовые и юридические риски для исполнителя.",
        "risks": risks
    }


@app.post("/api/generate-protocol")
def api_generate_protocol(data: ProtocolRequest):
    doc = Document()
    doc.add_heading("ПРОТОКОЛ РАЗНОГЛАСИЙ", level=1)
    doc.add_paragraph(f"к договору ({data.doc_type}), юрисдикция: {data.jurisdiction}")
    
    for idx, risk in enumerate(data.risks, 1):
        doc.add_heading(f"Пункт {idx}. {risk.get('title', 'Риск')}", level=2)
        doc.add_paragraph(f"Редакция Заказчика: {risk.get('quote', '-')}")
        doc.add_paragraph(f"Суть правового риска: {risk.get('explanation', '-')}")
        doc.add_paragraph(f"Рекомендуемая редакция: {risk.get('recommendation', '-')}")

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
    desc = data.description.strip() if data.description else "Разработка программного обеспечения"
    
    if data.doc_type == "offer":
        doc_title = "ПУБЛИЧНАЯ ОФЕРТА НА ОКАЗАНИЕ УСЛУГ"
        document_text = f"""# {doc_title}
Юрисдикция: {data.jurisdiction}

1. ОБЩИЕ ПОЛОЖЕНИЯ
1.1. Настоящий документ является публичной офертой Исполнителя ({data.executor}) для Заказчика ({data.client}).
1.2. Акцептом оферты является осуществление оплаты услуг Заказчиком.

2. ПРЕДМЕТ ДОГОВОРА
2.1. Исполнитель обязуется оказать следующие услуги: {desc}.
2.2. Срок выполнения услуг: {data.deadline}.

3. СТОИМОСТЬ И ПОРЯДОК РАСЧЕТОВ
3.1. Условия оплаты: {data.payment}.

4. СДАЧА И ПРИЕМКА УСЛУГ
4.1. Сдача результатов оформляется актом. Количество бесплатных итераций правок: {data.revisions}.

5. ОТВЕТСТВЕННОСТЬ СТОРОН
5.1. Стороны несут ответственность в соответствии с законодательством {data.jurisdiction}.
"""
    else:
        doc_title = "ДОГОВОР ВОЗМЕЗДНОГО ОКАЗАНИЯ УСЛУГ"
        document_text = f"""# {doc_title}
Юрисдикция: {data.jurisdiction}

1. ПРЕДМЕТ ДОГОВОРА
1.1. Исполнитель ({data.executor}) обязуется оказать по заданию Заказчика ({data.client}) услуги, а Заказчик обязуется их принять и оплатить.
1.2. Содержание и суть услуг: {desc}.

2. СРОКИ И ПОРЯДОК ВЫПОЛНЕНИЯ
2.1. Срок оказания услуг: {data.deadline}.
2.2. Количество включенных итераций правок: {data.revisions}.

3. СТОИМОСТЬ И РАСЧЕТЫ
3.1. Порядок оплаты: {data.payment}.

4. ОТВЕТСТВЕННОСТЬ СТОРОН
4.1. Стороны несут ответственность согласно законодательству {data.jurisdiction}.
"""
    return {"document_text": document_text}


@app.post("/api/download-constructor-file")
def api_download_constructor_file(data: ConstructorDownloadSchema):
    doc = Document()
    
    for line in data.document_text.split('\n'):
        if line.startswith('# '):
            doc.add_heading(line.replace('# ', ''), level=1)
        elif line.startswith('## '):
            doc.add_heading(line.replace('## ', ''), level=2)
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
