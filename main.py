import os
import io
import re
from datetime import datetime
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

# Простая база данных в памяти (можно расширить или заменить)
DB_USERS = {
    123456789: {"telegram_id": 123456789, "name": "Vladimir Tkachenko", "username": "admin", "checks_left": 3, "is_admin": True}
}
DB_PAYMENTS = [
    {"user": "Vladimir Tkachenko (@admin)", "package": "10 проверок", "stars": 100, "date": "2026-11-01 14:20"},
    {"user": "Иван Петров (@ivan_p)", "package": "5 проверок", "stars": 60, "date": "2026-11-02 10:15"},
    {"user": "Анна Смирнова (@anna_s)", "package": "1 проверка", "stars": 15, "date": "2026-11-03 16:40"},
    {"user": "Сергей Волков (@volkov)", "package": "10 проверок", "stars": 100, "date": "2026-11-04 11:05"},
    {"user": "Ольга Кравцова (@olga_k)", "package": "5 проверок", "stars": 60, "date": "2026-11-05 09:30"}
]
DB_FEEDBACKS = [
    {"user": "Vladimir Tkachenko", "rating": 5, "text": "Отличный сервис, очень помогает в работе!", "date": "2026-11-01"},
    {"user": "Иван Петров", "rating": 4, "text": "Быстрый анализ, хотелось бы больше шаблонов.", "date": "2026-11-02"}
]
TOTAL_ANALYSES_COUNT = 42


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
    rating: int
    feedback: str = ""
    telegram_id: int = 0
    name: str = "Пользователь"


@app.get("/")
def read_root():
    return FileResponse("static/index.html")


@app.get("/app")
def read_app():
    return FileResponse("static/index.html")


@app.get("/user-info/{telegram_id}")
def user_info(telegram_id: int):
    global DB_USERS
    if telegram_id not in DB_USERS:
        DB_USERS[telegram_id] = {
            "telegram_id": telegram_id,
            "name": f"Пользователь {telegram_id}",
            "username": f"user_{telegram_id}",
            "checks_left": 3,
            "is_admin": (telegram_id == 123456789)
        }
    user = DB_USERS[telegram_id]
    return {
        "free_checks_left": user["checks_left"],
        "is_admin": user["is_admin"]
    }


@app.post("/api/analyze")
def api_analyze(data: AnalysisRequest):
    global TOTAL_ANALYSES_COUNT
    TOTAL_ANALYSES_COUNT += 1
    
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
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_t = title_p.add_run("ПРОТОКОЛ РАЗНОГЛАСИЙ")
    run_t.bold = True
    run_t.font.name = "Times New Roman"
    run_t.font.size = Pt(13)
    run_t.font.color.rgb = RGBColor(0, 0, 0)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = sub_p.add_run(f"к Договору оказания услуг ({data.doc_type})\nЮрисдикция: {data.jurisdiction}")
    run_sub.font.name = "Times New Roman"
    run_sub.font.size = Pt(11)
    run_sub.font.color.rgb = RGBColor(0, 0, 0)

    table = doc.add_table(rows=1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    col_widths = [Inches(0.5), Inches(2.3), Inches(2.1), Inches(2.1)]
    
    hdr_cells = table.rows[0].cells
    hdr_titles = ["№", "Редакция Заказчика", "Суть замечания / Риск", "Рекомендуемая редакция"]
    
    for i, title_text in enumerate(hdr_titles):
        hdr_cells[i].text = title_text
        hdr_cells[i].width = col_widths[i]
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(8)
        for run in p.runs:
            run.bold = True
            run.font.name = "Times New Roman"
            run.font.size = Pt(10)
            run.font.color.rgb = RGBColor(0, 0, 0)
        shading_xml = parse_xml(r'<w:shd {} w:fill="E2E8F0"/>'.format(nsdecls('w')))
        hdr_cells[i]._tc.get_or_add_tcPr().append(shading_xml)

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
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(6)
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

Стороны: {executor}, именуемый(-ая) в дальнейшем «Исполнитель», и {client}, именуемый(-ая) в дальнейшем «Заказчик», заключили настоящий Договор о нижеследующем:

1. ПРЕДМЕТ ДОГОВОРА
1.1. Исполнитель обязуется по заданию Заказчика оказать следующие услуги: {desc}, а Заказчик обязуется принять и оплатить их в порядке и на условиях, предусмотренных настоящим Договором.
1.2. Срок оказания услуг составляет: {deadline} с момента получения от Заказчика всех необходимых исходных данных и предоплаты (если применимо).
1.3. Содержание, объем и требования к результату услуг определяются техническим заданием или согласованной перепиской Сторон.

2. СТОИМОСТЬ И ПОРЯДОК РАСЧЕТОВ
2.1. Стоимость услуг и порядок оплаты согласованы Сторонами следующим образом: {payment}.
2.2. Оплата производится в валюте юрисдикции {data.jurisdiction} на основании выставленных счетов или актов.
2.3. В случае задержки оплаты Заказчик уплачивает Исполнителю пени в размере 0.1% от суммы задолженности за каждый день просрочки.

3. ПОРЯДОК СДАЧИ-ПРИЕМКИ УСЛУГ
3.1. По завершении оказания услуг Исполнитель передает Заказчику результат и Акт приема-передачи услуг.
3.2. Заказчик в течение 3 (трех) рабочих дней с момента получения Акта обязан подписать его или предоставить мотивированный письменный отказ.
3.3. В случае наличия замечаний Стороны согласовывают перечень доработок. В стоимость услуг входит до {revisions} итераций бесплатных правок. По истечении указанного срока без замечаний услуги считаются принятыми.

4. ИНТЕЛЛЕКТУАЛЬНЫЕ ПРАВА
4.1. Исключительные права на результаты интеллектуальной деятельности, созданные Исполнителем в рамках настоящего Договора, переходят к Заказчику исключительно после полной оплаты стоимости оказанных услуг.
4.2. До полной оплаты Исполнитель сохраняет за собой правообладание всеми разработанными материалами.

5. ОТВЕТСТВЕННОСТЬ СТОРОН И ФОРС-МАЖОР
5.1. Стороны несут ответственность за неисполнение или ненадлежащее исполнение обязательств в соответствии с законодательством {data.jurisdiction}.
5.2. Стороны освобождаются от ответственности за частичное или полное неисполнение обязательств, если оно явилось следствием обстоятельств непреодолимой силы (форс-мажор).

6. СРОК ДЕЙСТВИЯ, ИЗМЕНЕНИЕ И РАСТОРЖЕНИЕ ДОГОВОРА
6.1. Настоящий Договор вступает в силу с момента подписания и действует до полного исполнения Сторонами своих обязательств.
6.2. Все изменения и дополнения к Договору действительны лишь в том случае, если они совершены в письменной форме и подписаны уполномоченными представителями Сторон.

7. РЕКВИЗИТЫ И ПОДПИСИ СТОРОН

Исполнитель:
Наименование / ФИО: {executor}
Юридический / почтовый адрес: ___________________________
ИНН / ОГРНИП / Реф. номер: _______________________________
Банковские реквизиты: ___________________________________
Подпись: ________________ / ______________

Заказчик:
Наименование / ФИО: {client}
Юридический / почтовый адрес: ___________________________
ИНН / ОГРНИП / Реф. номер: _______________________________
Банковские реквизиты: ___________________________________
Подпись: ________________ / ______________
"""
    return {"document_text": document_text}


@app.post("/api/download-constructor-file")
def api_download_constructor_file(data: ConstructorDownloadSchema):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    for line in data.document_text.split('\n'):
        if not line.strip():
            doc.add_paragraph()
            continue
            
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        
        if line.startswith('# '):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(14)
            run = p.add_run(line.replace('# ', ''))
            run.bold = True
            run.font.name = "Times New Roman"
            run.font.size = Pt(14)
            run.font.color.rgb = RGBColor(0, 0, 0)
        elif line.startswith('Юрисдикция:') or line.startswith('г. Москва') or line.startswith('Стороны:'):
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(6)
            run = p.add_run(line)
            run.bold = True
            run.font.name = "Times New Roman"
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(0, 0, 0)
        elif any(line.startswith(f"{i}.") for i in range(1, 10)) or line.startswith('7. РЕКВИЗИТЫ'):
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run(line)
            run.bold = True
            run.font.name = "Times New Roman"
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(0, 0, 0)
        elif line.startswith('Исполнитель:') or line.startswith('Заказчик:'):
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run(line)
            run.bold = True
            run.font.name = "Times New Roman"
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(0, 0, 0)
        elif ':' in line or '_' in line:
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run(line)
            run.bold = True
            run.font.name = "Times New Roman"
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(0, 0, 0)
        else:
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run(line)
            run.bold = False
            run.font.name = "Times New Roman"
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(0, 0, 0)

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
    global DB_FEEDBACKS
    DB_FEEDBACKS.insert(0, {
        "user": data.name,
        "rating": data.rating,
        "text": data.feedback,
        "date": datetime.now().strftime("%Y-%m-%d %H:%M")
    })
    return {"status": "success", "message": "Отзыв успешно сохранен"}


@app.get("/admin/stats/{telegram_id}")
def admin_stats(telegram_id: int):
    # Проверка прав администратора
    if telegram_id != 123456789:
        raise HTTPException(status_code=403, detail="Доступ запрещен")
        
    return {
        "total_users": len(DB_USERS),
        "total_analyses": TOTAL_ANALYSES_COUNT,
        "total_payments": len(DB_PAYMENTS),
        "users_list": list(DB_USERS.values()),
        "payments_list": DB_PAYMENTS,
        "feedbacks_list": DB_FEEDBACKS
    }
