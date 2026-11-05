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
    
    return {
        "overall_risk_level": "Высокий",
        "summary": f"Глубокий правовой анализ договора ({data.doc_type}) для юрисдикции {data.jurisdiction}. Выявлены критические юридические уязвимости, кабальные штрафы и риски одностороннего изменения условий.",
        "risks": [
            {
                "category": "Статья 1. Предмет и порядок оказания услуг",
                "title": "Неопределенность и размытость технического задания",
                "quote": "Исполнитель обязуется оказать услуги по заданию Заказчика.",
                "explanation": "Отсутствие четких границ и критериев приемки позволяет заказчику требовать выполнения любых дополнительных задач без оплаты.",
                "recommendation": "Зафиксировать исчерпывающий перечень услуг и этапов в приложении к договору."
            },
            {
                "category": "Статья 2. Финансовые условия и ответственность",
                "title": "Несоразмерная неустойка за просрочку обязательств",
                "quote": "В случае просрочки Исполнитель выплачивает пени в размере 1% от общей суммы договора за каждый день.",
                "explanation": "Размер санкций многократно превышает ставки ЦБ и создает критический финансовый риск при минимальной задержке.",
                "recommendation": "Снизить размер пени до 0.1% от суммы неисполненного обязательства за каждый день просрочки."
            },
            {
                "category": "Статья 3. Порядок сдачи-приемки работ",
                "title": "Право заказчика на необоснованный отказ от подписания актов",
                "quote": "Заказчик вправе в любое время отказаться от подписания акта при наличии любых замечаний.",
                "explanation": "Создает риск неоплаты фактически выполненных и полезных для заказчика этапов под предлогом субъективных претензий.",
                "recommendation": "Установить жесткий срок для подачи мотивированного отказа (3 рабочих дня), после чего работы считаются принятыми."
            },
            {
                "category": "Статья 4. Расторжение договора",
                "title": "Односторонний отказ без компенсации понесенных расходов",
                "quote": "Заказчик вправе в одностороннем порядке отказаться от исполнения договора в любое время.",
                "explanation": "Отсутствие механизма возмещения исполнителю уже затраченных ресурсов и средств на момент расторжения ущемляет его права.",
                "recommendation": "Обязать заказчика оплатить фактически оказанные услуги и подтвержденные расходы на дату уведомления."
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
1.1. Настоящий документ является публичной офертой Исполнителя ({data.executor}) по оказанию услуг для Заказчика ({data.client}).
1.2. Акцептом оферты признается осуществление полной или частичной оплаты услуг Заказчиком.

2. ПРЕДМЕТ ДОГОВОРА
2.1. Исполнитель обязуется оказать следующие услуги: {data.description}.
2.2. Срок выполнения услуг: {data.deadline}.

3. СТОИМОСТЬ И ПОРЯДОК РАСЧЕТОВ
3.1. Условия расчетов: {data.payment}.

4. СДАЧА И ПРИЕМКА УСЛУГ
4.1. Сдача результатов оформляется актом. Количество включенных итераций правок: {data.revisions}.

5. ОТВЕТСТВЕННОСТЬ СТОРОН
5.1. Стороны несут ответственность в соответствии с законодательством {data.jurisdiction}.
"""
    else:
        doc_title = "ДОГОВОР ВОЗМЕЗДНОГО ОКАЗАНИЯ УСЛУГ"
        document_text = f"""# {doc_title}
Юрисдикция: {data.jurisdiction}

1. ПРЕДМЕТ ДОГОВОРА
1.1. Исполнитель ({data.executor}) обязуется оказать по заданию Заказчика ({data.client}) услуги, а Заказчик обязуется их принять и оплатить.
1.2. Содержание и суть услуг: {data.description}.

2. СРОКИ И ПОРЯДОК ВЫПОЛНЕНИЯ
2.1. Срок оказания услуг: {data.deadline}.
2.2. Количество итераций правок: {data.revisions}.

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
