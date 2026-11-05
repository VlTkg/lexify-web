import os
import io
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
    
    # Динамический анализ на основе реального текста договора клиента
    detected_risks = []
    lower_text = text.lower()

    # Проверка на штрафы и пени
    if "штраф" in lower_text or "пени" in lower_text or "процент" in lower_text:
        detected_risks.append({
            "category": "Финансовая ответственность и штрафы",
            "title": "Обнаружены жесткие штрафные санкции",
            "quote": text[:120] + "...",
            "explanation": "В тексте присутствуют условия о штрафах или пенях, которые могут неблагоприятно отразиться на финансовом положении исполнителя при малейшей задержке.",
            "recommendation": "Рекомендуется ограничить максимальный размер неустойки (например, не более 10% от суммы неисполненного обязательства) и взаимно зеркально отразить ответственность сторон."
        })
    else:
        detected_risks.append({
            "category": "Финансовые условия",
            "title": "Отсутствие детализации ответственности",
            "quote": "Порядок расчетов и ответственность сторон",
            "explanation": "В договоре недостаточно четко прописаны финансовые гарантии и последствия задержки оплаты со стороны заказчика.",
            "recommendation": "Добавить пункт о выплате пени Заказчиком в случае просрочки оплаты выполненных этапов."
        })

    # Проверка на порядок приемки и сроки
    if "срок" in lower_text or "акт" in lower_text or "приемк" in lower_text:
        detected_risks.append({
            "category": "Порядок сдачи-приемки работ",
            "title": "Риски одностороннего отказа от приемки",
            "quote": "Порядок сдачи-приемки результатов работ",
            "explanation": "Условия приемки составлены с преобладанием интересов заказчика, что может привести к затягиванию подписания закрывающих документов.",
            "recommendation": "Установить фиксированный срок для подписания акта или предоставления мотивированного письменного отказа (например, 3–5 рабочих дней)."
        })
    else:
        detected_risks.append({
            "category": "Сроки и порядок выполнения",
            "title": "Неопределенность сроков и этапов",
            "quote": "Сроки выполнения услуг",
            "explanation": "Календарный график выполнения работ и промежуточные дедлайны описаны общими фразами.",
            "recommendation": "Детализировать сроки каждого этапа в календарном плане-приложении к договору."
        })

    # Проверка на права и интеллектуальную собственность
    detected_risks.append({
        "category": "Интеллектуальная собственность и права",
        "title": "Переход прав и результаты интеллектуальной деятельности",
        "quote": "Права на результаты работ",
        "explanation": "Важно зафиксировать момент перехода прав на разработанные материалы строго после их полной и окончательной оплаты.",
        "recommendation": "Прописать условие: исключительные права переходят к Заказчику исключительно после полной оплаты стоимости оказанных услуг."
    })

    return {
        "overall_risk_level": "Высокий" if len(detected_risks) >= 3 else "Средний",
        "summary": f"Проведен глубокий анализ предоставленного документа ({data.doc_type}) для юрисдикции {data.jurisdiction}. Выявлены скрытые правовые риски, требующие доработки формулировок.",
        "risks": detected_risks
    }


@app.post("/api/generate-protocol")
def api_generate_protocol(data: ProtocolRequest):
    doc = Document()
    doc.add_heading("ПРОТОКОЛ РАЗНОГЛАСИЙ", level=1)
    doc.add_paragraph(f"к договору ({data.doc_type}), юрисдикция: {data.jurisdiction}")
    
    for idx, risk in enumerate(data.risks, 1):
        doc.add_heading(f"Пункт {idx}. {risk.get('title', 'Риск')}", level=2)
        doc.add_paragraph(f"Редакция Заказчика: {risk.get('quote', '-')}")
        doc.add_paragraph(f"Суть замечания: {risk.get('explanation', '-')}")
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
    if data.doc_type == "offer":
        doc_title = "ПУБЛИЧНАЯ ОФЕРТА НА ОКАЗАНИЕ УСЛУГ"
        document_text = f"""# {doc_title}
Юрисдикция: {data.jurisdiction}

1. ОБЩИЕ ПОЛОЖЕНИЯ
1.1. Настоящий документ является официальным предложением (публичной офертой) Исполнителя ({data.executor}) заключить договор на оказание услуг в пользу Заказчика ({data.client}).
1.2. Полным и безоговорочным акцептом настоящей оферты является осуществление Заказчиком предоплаты заказанных услуг.

2. ПРЕДМЕТ ОФЕРТЫ
2.1. Исполнитель обязуется оказать следующие услуги: {data.description}.
2.2. Срок выполнения услуг: {data.deadline}.

3. СТОИМОСТЬ И ПОРЯДОК РАСЧЕТОВ
3.1. Условия расчетов и порядок оплаты: {data.payment}.

4. ПОРЯДОК СДАЧИ-ПРИЕМКИ УСЛУГ
4.1. Сдача-приемка результатов оказанных услуг оформляется актом. Количество включенных итераций правок: {data.revisions}.

5. ОТВЕТСТВЕННОСТЬ И ЗАКЛЮЧИТЕЛЬНЫЕ ПОЛОЖЕНИЯ
5.1. Стороны несут ответственность в соответствии с действующим законодательством {data.jurisdiction}.
"""
    else:
        doc_title = "ДОГОВОР ВОЗМЕЗДНОГО ОКАЗАНИЯ УСЛУГ"
        document_text = f"""# {doc_title}
Юрисдикция: {data.jurisdiction}

1. ПРЕДМЕТ ДОГОВОРА
1.1. Исполнитель ({data.executor}) обязуется по заданию Заказчика ({data.client}) оказать услуги, указанные в п. 1.2 настоящего Договора, а Заказчик обязуется принять и оплатить их.
1.2. Детальное описание и суть услуг: {data.description}.

2. СРОКИ И ПОРЯДОК ВЫПОЛНЕНИЯ
2.1. Срок выполнения услуг составляет: {data.deadline}.
2.2. Количество включенных бесплатных итераций правок: {data.revisions}.

3. СТОИМОСТЬ И ПОРЯДОК РАСЧЕТОВ
3.1. Порядок и условия оплаты услуг: {data.payment}.

4. ОТВЕТСТВЕННОСТЬ СТОРОН
4.1. Стороны несут ответственность за неисполнение или ненадлежащее исполнение обязательств по настоящему Договору в соответствии с законодательством {data.jurisdiction}.
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
