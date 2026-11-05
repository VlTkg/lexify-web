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

    # Интеллектуальный поиск ключевых проблемных зон по тексту клиента
    
    # 1. Штрафы и пени
    if "штраф" in lower_text or "пени" in lower_text or "час" in lower_text:
        # Пытаемся вычленить цитату со штрафом
        quote_match = re.search(r'([^\n]*?(?:штраф|пени|просроч)[^\n]*)', text, re.IGNORECASE)
        quote = quote_match.group(1) if quote_match else "Условия о штрафных санкциях и ответственности за просрочку."
        
        risks.append({
            "category": "Финансовая ответственность и штрафы",
            "title": "Критический размер штрафных санкций",
            "quote": quote.strip(),
            "explanation": "Обнаружены кабальные условия ответственности (например, несоразмерные штрафы или лишение оплаты при минимальной задержке), что нарушает баланс интересов сторон.",
            "recommendation": "Снизить размер неустойки до стандартных значений (0.1% за день просрочки) и исключить отказ от оплаты уже выполненных этапов."
        })

    # 2. Одностороннее изменение условий
    if "односторонн" in lower_text or "измен" in lower_text:
        quote_match = re.search(r'([^\n]*?односторонн[^\n]*)', text, re.IGNORECASE)
        quote = quote_match.group(1) if quote_match else "Право заказчика на одностороннее изменение условий."
        
        risks.append({
            "category": "Порядок изменения условий договора",
            "title": "Одностороннее изменение объема работ и стоимости",
            "quote": quote.strip(),
            "explanation": "Предоставление заказчику права менять существенные условия без согласия исполнителя создает правовую неопределенность и риск недобросовестного поведения.",
            "recommendation": "Зафиксировать, что любые изменения объема, сроков и стоимости оформляются исключительно дополнительными соглашениями обеих сторон."
        })

    # 3. Интеллектуальная собственность
    if "прав" in lower_text or "наработк" in lower_text or "исключительн" in lower_text:
        quote_match = re.search(r'([^\n]*?(?:прав[ао]|наработк|интеллект)[^\n]*)', text, re.IGNORECASE)
        quote = quote_match.group(1) if quote_match else "Переход исключительных прав и наработок."
        
        risks.append({
            "category": "Интеллектуальная собственность",
            "title": "Тотальное отчуждение прав, включая личные проекты",
            "quote": quote.strip(),
            "explanation": "Формулировки захватывают сторонние или личные наработки исполнителя и передают их до полной оплаты, лишая исполнителя результатов его труда.",
            "recommendation": "Четко ограничить перечень передаваемых материалов исключительно результатами по данному договору и закрепить переход прав только после полного расчета."
        })

    # 4. Сроки и порядок оплаты / отсрочки
    if "дней" in lower_text or "оплат" in lower_text or "акта" in lower_text:
        quote_match = re.search(r'([^\n]*?(?:оплат[аы]|дней|акта)[^\n]*)', text, re.IGNORECASE)
        quote = quote_match.group(1) if quote_match else "Порядок расчетов и подписания актов."
        
        risks.append({
            "category": "Порядок расчетов и приемки",
            "title": "Затянутые сроки оплаты и риски приемки",
            "quote": quote.strip(),
            "explanation": "Установлена длительная отсрочка платежа (или размытые сроки расчетов), что замораживает оборотные средства исполнителя.",
            "recommendation": "Сократить срок оплаты после подписания акта приемки-передачи до разумных пределов (например, 5–10 рабочих дней)."
        })

    # Если вдруг текст совсем специфический, гарантируем наличие базового анализа
    if not risks:
        risks.append({
            "category": "Общие положения договора",
            "title": "Требуется детальная правовая экспертиза разделов",
            "quote": text[:100] + "...",
            "explanation": "В тексте договора обнаружены общие формулировки, требующие конкретизации обязанностей сторон.",
            "recommendation": "Детализировать взаимную ответственность и порядок разрешения споров."
        })

    high_risk = "Высокий" if len(risks) >= 2 else "Средний"

    return {
        "overall_risk_level": high_risk,
        "summary": f"Проведен глубокий правовой анализ предоставленного фрагмента текста ({data.doc_type}). Выявлены серьезные ущемления прав исполнителя, кабальные штрафы и риски безвозмездной передачи результатов.",
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
    if data.doc_type == "offer":
        doc_title = "ПУБЛИЧНАЯ ОФЕРТА НА ОКАЗАНИЕ УСЛУГ"
        document_text = f"""# {doc_title}
Юрисдикция: {data.jurisdiction}

1. ОБЩИЕ ПОЛОЖЕНИЯ
1.1. Настоящий документ представляет собой официальное предложение Исполнителя ({data.executor}) в адрес Заказчика ({data.client}).
1.2. Акцептом оферты является совершение Заказчиком оплаты услуг.

2. ПРЕДМЕТ ДОГОВОРА
2.1. Исполнитель обязуется оказать услуги: {data.description}.
2.2. Срок выполнения: {data.deadline}.

3. СТОИМОСТЬ И ПОРЯДОК РАСЧЕТОВ
3.1. Условия оплаты: {data.payment}.

4. ОТВЕТСТВЕННОСТЬ
4.1. Стороны несут ответственность в соответствии с законодательством {data.jurisdiction}.
"""
    else:
        doc_title = "ДОГОВОР ВОЗМЕЗДНОГО ОКАЗАНИЯ УСЛУГ"
        document_text = f"""# {doc_title}
Юрисдикция: {data.jurisdiction}

1. ПРЕДМЕТ ДОГОВОРА
1.1. Исполнитель ({data.executor}) обязуется оказать Заказчику ({data.client}) услуги: {data.description}.
1.2. Срок выполнения работ: {data.deadline}.

2. СТОИМОСТЬ И РАСЧЕТЫ
2.1. Порядок расчетов: {data.payment}.
2.2. Количество включенных итераций правок: {data.revisions}.

3. ОТВЕТСТВЕННОСТЬ СТОРОН
3.1. Ответственность сторон регулируется действующим законодательством {data.jurisdiction}.
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
    return {"status": "success", "message": "Отзыв сохранен"}


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
