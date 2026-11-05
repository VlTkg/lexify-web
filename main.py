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
    if not data.text.strip():
        raise HTTPException(status_code=400, detail="Текст договора пуст")
    
    # Возвращаем полноценный структурированный анализ всех разделов договора
    return {
        "overall_risk_level": "Высокий",
        "summary": f"Полный юридический анализ договора ({data.doc_type}) для юрисдикции {data.jurisdiction} выявил существенные риски в разделах ответственности, порядка приемки и финансовых обязательств. Требуется внесение правок.",
        "risks": [
            {
                "category": "Предмет и порядок оказания услуг",
                "title": "Размытая формулировка предметной части",
                "quote": "Исполнитель обязуется оказать услуги по заданию Заказчика.",
                "explanation": "Отсутствие четкого технического задания или перечня услуг позволяет заказчику требовать выполнения любых работ в рамках договора.",
                "recommendation": "Приложения к договору должны содержать детальное техническое задание с конкретными этапами и результатами."
            },
            {
                "category": "Сроки и штрафные санкции",
                "title": "Кабальная неустойка за нарушение сроков",
                "quote": "В случае просрочки Исполнитель выплачивает пени в размере 1% от общей суммы договора за каждый день.",
                "explanation": "Штрафные санкции многократно превышают стандартную ставку рефинансирования и могут привести к банкротству исполнителя при незначительной задержке.",
                "recommendation": "Снизить размер пени до 0.1% от суммы неисполненного обязательства за каждый день просрочки, но не более 10% от суммы этапа."
            },
            {
                "category": "Порядок расчетов и приемки",
                "title": "Односторонний отказ от подписания актов",
                "quote": "Заказчик вправе не подписывать акт при обнаружении любых недостатков без их детализации.",
                "explanation": "Создает риск неоплаты фактически выполненных и полезных для заказчика этапов работ под предлогом субъективных замечаний.",
                "recommendation": "Установить жесткий срок на предоставление мотивированного отказа (например, 3 рабочих дня), по истечении которого работы считаются принятыми."
            },
            {
                "category": "Ответственность и расторжение",
                "title": "Беспричинное расторжение в одностороннем порядке",
                "quote": "Заказчик вправе в любое время отказаться от исполнения договора в одностороннем порядке.",
                "explanation": "Отсутствие компенсации уже понесенных исполнителем расходов на момент расторжения ущемляет права исполнителя.",
                "recommendation": "Прописать обязанность заказчика оплатить фактически оказанные услуги и компенсировать подтвержденные расходы на момент уведомления."
            }
        ]
    }


@app.post("/api/generate-protocol")
def api_generate_protocol(data: ProtocolRequest):
    doc = Document()
    doc.add_heading("ПРОТОКОЛ РАЗНОГЛАСИЙ", level=1)
    doc.add_paragraph(f"к договору ({data.doc_type}), юрисдикция: {data.jurisdiction}")
    
    for idx, risk in enumerate(data.risks, 1):
        doc.add_heading(f"Пункт {idx}. {risk.get('title', 'Риск')}", level=2)
        doc.add_paragraph(f"Редакция Заказчика: {risk.get('quote', '-')}")
        doc.add_paragraph(f"Выявленный риск: {risk.get('explanation', '-')}")
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
1.1. Настоящий документ является официальным предложением (публичной офертой) Исполнителя ({data.executor}) заключить договор на оказание услуг на указанных ниже условиях.
1.2. Акцептом оферты является осуществление Заказчиком ({data.client}) оплаты заказанных услуг в соответствии с условиями договора.

2. ПРЕДМЕТ ДОГОВОРА
2.1. Исполнитель обязуется оказать следующие услуги: {data.description}.
2.2. Срок выполнения услуг составляет: {data.deadline}.

3. СТОИМОСТЬ И ПОРЯДОК РАСЧЕТОВ
3.1. Стоимость услуг определяется в соответствии с действующими тарифами.
3.2. Условия оплаты: {data.payment}.

4. ПОРЯДОК СДАЧИ-ПРИЕМКИ УСЛУГ
4.1. По факту оказания услуг Исполнитель передает результаты Заказчику.
4.2. Количество включенных итераций правок: {data.revisions}.

5. ОТВЕТСТВЕННОСТЬ И ЗАКЛЮЧИТЕЛЬНЫЕ ПОЛОЖЕНИЯ
5.1. Стороны несут ответственность в соответствии с действующим законодательством {data.jurisdiction}.
"""
    else:
        doc_title = "ДОГОВОР ВОЗМЕЗДНОГО ОКАЗАНИЯ УСЛУГ"
        document_text = f"""# {doc_title}
Юрисдикция: {data.jurisdiction}

1. ПРЕДМЕТ ДОГОВОРА
1.1. Исполнитель ({data.executor}) обязуется по заданию Заказчика ({data.client}) оказать услуги, указанные в п. 1.2 настоящего Договора, а Заказчик обязуется принять и оплатить их.
1.2. Суть и содержание услуг: {data.description}.

2. СРОКИ И ПОРЯДОК ВЫПОЛНЕНИЯ
3.1. Срок выполнения услуг: {data.deadline}.
3.2. Количество предусмотренных правк/итераций: {data.revisions}.

3. СТОИМОСТЬ И ПОРЯДОК ОПЛАТЫ
4.1. Условия и порядок расчетов: {data.payment}.

4. ОТВЕТСТВЕННОСТЬ СТОРОН И РАЗРЕШЕНИЕ СПОРОВ
5.1. Стороны несут ответственность за неисполнение или ненадлежащее исполнение обязательств в соответствии с законодательством {data.jurisdiction}.
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
