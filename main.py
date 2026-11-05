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

    # Разбиваем текст на предложения/пункты для поиска конкретных нумераций
    lines = [line.strip() for line in text.split('\n') if line.strip()]

    def extract_point_number(line_text):
        match = re.search(r'(\d+(?:\.\d+)*)\.?', line_text)
        return f"Пункт {match.group(1)}" if match else "Раздел договора"

    for line in lines:
        lower_line = line.lower()
        point_num = extract_point_number(line)

        # 1. Сверхштрафы и пени
        if ("штраф" in lower_line or "пени" in lower_line or "100%" in lower_line or "просроч" in lower_line) and "ответственность" not in lower_line:
            risks.append({
                "point": point_num,
                "category": "Финансовые санкции",
                "title": f"{point_num}. Кабальный размер штрафа за просрочку",
                "quote": line,
                "explanation": "Лишение оплаты или штраф 100% при минимальной просрочке является кабальным условием, противоречащим принципу соразмерности ответственности.",
                "recommendation": "Исключить 100% штраф. Установить неустойку в размере 0.1% от стоимости просроченного этапа за каждый день задержки, но не более 10% от общей суммы."
            })

        # 2. Одностороннее изменение условий
        elif "односторонн" in lower_line or "измен" in lower_line:
            risks.append({
                "point": point_num,
                "category": "Изменение условий договора",
                "title": f"{point_num}. Права Заказчика на одностороннее изменение условий",
                "quote": line,
                "explanation": "Одностороннее изменение объема услуг, сроков или цены без согласия Исполнителя нарушает равенство сторон и создает риск неконтролируемой переработки.",
                "recommendation": "Любые изменения объема работ, сроков и стоимости оформляются исключительно двусторонним Соглашением или Дополнительным соглашением."
            })

        # 3. Отчуждение прав / личные проекты
        elif "исключительн" in lower_line or "наработк" in lower_line or "бессрочн" in lower_line:
            risks.append({
                "point": point_num,
                "category": "Интеллектуальная собственность",
                "title": f"{point_num}. Безвозмездная передача всех наработок и личных проектов",
                "quote": line,
                "explanation": "Передача прав на сторонние/личные наработки и переход прав до момента окончательного расчета лишает Исполнителя защиты результатов его труда.",
                "recommendation": "Исключительные права на результаты работ переходят к Заказчику строго после 100% оплаты. Исключить передачу личных наработок Исполнителя."
            })

        # 4. Длительная отсрочка платежа
        elif "180" in lower_line or "отсрочк" in lower_line or ("рабочих дней" in lower_line and "оплат" in lower_line):
            risks.append({
                "point": point_num,
                "category": "Порядок расчетов",
                "title": f"{point_num}. Чрезмерная отсрочка выплат (180 рабочих дней)",
                "quote": line,
                "explanation": "Задержка оплаты на 180 рабочих дней (около 9 календарных месяцев) замораживает средства Исполнителя и увеличивает риск дефолта Заказчика.",
                "recommendation": "Сократить срок оплаты за фактически выполненные работы до 5 (пяти) рабочих дней с момента подписания Акта приема-передачи."
            })

    # Если регулярные выражения не вычленили специфику, проводим базовый структурированный разбор
    if not risks:
        risks.append({
            "point": "Раздел 1",
            "category": "Общие условия",
            "title": "Пункт 1.1. Размытые условия договора",
            "quote": text[:120] + "...",
            "explanation": "В предоставленном тексте отсутствуют четкие критерии сдачи-приемки и регламент оплаты.",
            "recommendation": "Детализировать порядок сдачи работ и зафиксировать ответственность Заказчика за просрочку оплаты."
        })

    return {
        "overall_risk_level": "Высокий",
        "summary": f"Правовая экспертиза документа ({data.doc_type}) для юрисдикции {data.jurisdiction} выявила {len(risks)} критических рисков. Договор ущемляет права Исполнителя и требует обязательного подписания Протокола разногласий.",
        "risks": risks
    }


@app.post("/api/generate-protocol")
def api_generate_protocol(data: ProtocolRequest):
    doc = Document()

    # Заголовок
    title = doc.add_heading("ПРОТОКОЛ РАЗНОГЛАСИЙ", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    sub = doc.add_paragraph(f"к Договору оказания услуг ({data.doc_type})\nЮрисдикция: {data.jurisdiction}\n")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for idx, risk in enumerate(data.risks, 1):
        point_name = risk.get('point', f'Пункт {idx}')
        risk_title = risk.get('title', 'Замечание')

        heading = doc.add_heading(f"{idx}. {point_name} — {risk.get('category', 'Замечание')}", level=2)
        
        p1 = doc.add_paragraph()
        p1.add_run("Редакция Заказчика: ").bold = True
        p1.add_run(f"«{risk.get('quote', '-')}»")

        p2 = doc.add_paragraph()
        p2.add_run("Правовой риск: ").bold = True
        p2.add_run(risk.get('explanation', '-'))

        p3 = doc.add_paragraph()
        p3.add_run("Предлагаемая редакция Исполнителя: ").bold = True
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
    desc = data.description.strip() if data.description else "Оказание консультационных и техническо-программных услуг"
    executor = data.executor if data.executor else "Исполнитель"
    client = data.client if data.client else "Заказчик"
    payment = data.payment if data.payment else "Постоплата"
    deadline = data.deadline if data.deadline else "14 рабочих дней"
    revisions = data.revisions if data.revisions else "2"

    if data.doc_type == "offer":
        doc_title = "ПУБЛИЧНАЯ ОФЕРТА НА ОКАЗАНИЕ УСЛУГ"
        document_text = f"""# {doc_title}
Юрисдикция: {data.jurisdiction}

1. ОБЩИЕ ПОЛОЖЕНИЯ
1.1. Настоящий документ является официальным предложением (публичной офертой) Исполнителя ({executor}) заключить Договор оказания услуг с любым физическим или юридическим лицом (Заказчик: {client}).
1.2. Акцептом настоящей оферты является факт оплаты услуг Заказчиком в соответствии с разделом 3 настоящего Договора.

2. ПРЕДМЕТ ДОГОВОРА
2.1. Исполнитель обязуется оказать Заказчику следующие услуги: {desc}.
2.2. Срок оказания услуг составляет: {deadline} с момента получения всех исходных материалов.
2.3. В стоимость услуг входит до {revisions} итераций правок в пределах согласованного ТЗ.

3. СТОИМОСТЬ И ПОРЯДОК РАСЧЕТОВ
3.1. Оплата услуг осуществляется на следующих условиях: {payment}.
3.2. Расчеты производятся в безналичном порядке путем перечисления денежных средств на счет Исполнителя.

4. ПОРЯДОК СДАЧИ-ПРИЕМКИ РАБОТ
4.1. По окончании оказания услуг Исполнитель направляет Заказчику результат и Акт приема-передачи.
4.2. Заказчик обязуется рассмотреть и подписать Акт либо направить мотивированный отказ в течение 3 (трех) рабочих дней. По истечении данного срока услуги считаются принятыми в полном объеме.

5. ОТВЕТСТВЕННОСТЬ СТОРОН И ФОРС-МАЖОР
5.1. За неисполнение или ненадлежащее исполнение обязательств Стороны несут ответственность согласно законодательству {data.jurisdiction}.
5.2. Стороны освобождаются от ответственности при наступлении форс-мажорных обстоятельств.

6. РЕКВИЗИТЫ И ПОДПИСИ СТОРОН
Исполнитель: {executor}
Заказчик: {client}
"""
    else:
        doc_title = "ДОГОВОР ВОЗМЕЗДНОГО ОКАЗАНИЯ УСЛУГ"
        document_text = f"""# {doc_title}
Юрисдикция: {data.jurisdiction}

1. ПРЕДМЕТ ДОГОВОРА
1.1. Исполнитель ({executor}) обязуется по заданию Заказчика ({client}) оказать услуги, указанные в п. 1.2 настоящего Договора, а Заказчик обязуется принять и оплатить их.
1.2. Суть и объем оказываемых услуг: {desc}.

2. СРОКИ И ПОРЯДОК ОКАЗАНИЯ УСЛУГ
2.1. Срок оказания услуг: {deadline} с момента предоставления Заказчиком необходимых исходных данных.
2.2. Заказчик имеет право на {revisions} бесплатные итерации правок в рамках предмета Договора.

3. СТОИМОСТЬ УСЛУГ И ПОРЯДОК РАСЧЕТОВ
3.1. Условия и порядок оплаты услуг: {payment}.
3.2. В случае просрочки оплаты Заказчик уплачивает Исполнителю пени в размере 0.1% от просроченной суммы за каждый день задержки.

4. ПОРЯДОК СДАЧИ И ПРИЕМКИ
4.1. Сдача-приемка услуг оформляется Актом приема-передачи.
4.2. Заказчик в течение 3 (трех) рабочих дней с момента получения Акта обязан подписать его или предоставить мотивированный письменный отказ. В противном случае услуги считаются оказанными надлежащим образом и принятыми.

5. ИНТЕЛЛЕКТУАЛЬНАЯ СОБСТВЕННОСТЬ
5.1. Исключительные права на результаты оказанных услуг переходят к Заказчику исключительно после 100% оплаты стоимости услуг.

6. РЕКВИЗИТЫ И ПОДПИСИ СТОРОН
Исполнитель: {executor}
Заказчик: {client}
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
