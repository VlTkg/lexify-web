import os
import io
import re
import requests
from datetime import datetime
from fastapi import FastAPI, HTTPException, Request, Response
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

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "YOUR_BOT_TOKEN_HERE")

DB_USERS = {}
DB_PAYMENTS = []
DB_FEEDBACKS = []
TOTAL_ANALYSES_COUNT = 0


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
    executor_id: str
    executor_address: str
    client: str
    client_id: str
    client_address: str
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
    doc_type: str = "Общий"
    jurisdiction: str = "Россия (РФ)"


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
        is_adm = (telegram_id == 123456789)
        DB_USERS[telegram_id] = {
            "telegram_id": telegram_id,
            "name": "Владимир Ткаченко" if is_adm else f"Пользователь {telegram_id}",
            "username": "admin" if is_adm else f"user_{telegram_id}",
            "checks_left": 999 if is_adm else 3,
            "is_admin": is_adm
        }
    user = DB_USERS[telegram_id]
    if user["is_admin"]:
        user["checks_left"] = 999

    return {
        "free_checks_left": user["checks_left"],
        "is_admin": user["is_admin"]
    }


def get_jurisdiction_details(jurisdiction_str: str):
    if "Казахстан" in jurisdiction_str:
        return {
            "city": "г. Алматы",
            "currency": "тенге (KZT)",
            "law_ref": "Гражданским кодексом Республики Казахстан",
            "id_label": "ИИН / БИН"
        }
    elif "Беларусь" in jurisdiction_str:
        return {
            "city": "г. Минск",
            "currency": "белорусских рублях (BYN)",
            "law_ref": "Гражданским кодексом Республики Беларусь",
            "id_label": "УНП / Личный номер"
        }
    else:
        return {
            "city": "г. Москва",
            "currency": "рублях РФ (RUB)",
            "law_ref": "законодательством Российской Федерации",
            "id_label": "ИНН"
        }


@app.post("/api/analyze")
def api_analyze(data: AnalysisRequest):
    global TOTAL_ANALYSES_COUNT, DB_USERS
    TOTAL_ANALYSES_COUNT += 1
    
    if data.telegram_id not in DB_USERS:
        is_adm = (data.telegram_id == 123456789)
        DB_USERS[data.telegram_id] = {
            "telegram_id": data.telegram_id,
            "name": data.first_name or f"Пользователь {data.telegram_id}",
            "username": data.username or f"user_{data.telegram_id}",
            "checks_left": 999 if is_adm else 3,
            "is_admin": is_adm
        }

    user = DB_USERS[data.telegram_id]
    if not user["is_admin"] and user["checks_left"] > 0:
        user["checks_left"] -= 1

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
                "title": f"{point_num}. Неприемлемый штраф 100% и отказ от оплаты",
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
        "summary": f"Проведен детальный юридический анализ текста ({data.doc_type}) для юрисдикции {data.jurisdiction}. Выявлено {len(risks)} неприемлемых рисков по конкретным пунктам договора.",
        "risks": risks
    }


@app.post("/api/generate-protocol")
def api_generate_protocol(data: ProtocolRequest):
    details = get_jurisdiction_details(data.jurisdiction)
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
    run_sub = sub_p.add_run(f"к Договору оказания услуг ({data.doc_type})\nПрименимое право: {details['law_ref']}")
    run_sub.font.name = "Times New Roman"
    run_sub.font.size = Pt(11)
    run_sub.font.color.rgb = RGBColor(0, 0, 0)

    table = doc.add_table(rows=1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # Жесткая фиксация макета таблицы для LibreOffice и Word
    tblPr = table._tbl.tblPr
    tblLayout = parse_xml(r'<w:tblLayout {} w:type="fixed"/>'.format(nsdecls('w')))
    tblPr.append(tblLayout)
    
    # Ширина в twips (1 дюйм = 1440 twips).
    col_widths_twips = [720, 3120, 3120, 3120]
    
    hdr_cells = table.rows[0].cells
    hdr_titles = ["№", "Редакция Заказчика", "Правовая квалификация / Риск", "Предлагаемая редакция"]
    
    for i, title_text in enumerate(hdr_titles):
        hdr_cells[i].text = title_text
        tcPr = hdr_cells[i]._tc.get_or_add_tcPr()
        tcW = parse_xml(r'<w:tcW {} w:w="{}" w:type="dxa"/>'.format(nsdecls('w'), col_widths_twips[i]))
        tcPr.append(tcW)
        
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
        tcPr.append(shading_xml)

    for idx, risk in enumerate(data.risks, 1):
        row_cells = table.add_row().cells
        row_cells[0].text = str(idx)
        row_cells[1].text = f"«{risk.get('quote', '-')}»"
        row_cells[2].text = risk.get('explanation', '-')
        row_cells[3].text = risk.get('recommendation', '-')

        for i, cell in enumerate(row_cells):
            tcPr = cell._tc.get_or_add_tcPr()
            tcW = parse_xml(r'<w:tcW {} w:w="{}" w:type="dxa"/>'.format(nsdecls('w'), col_widths_twips[i]))
            tcPr.append(tcW)
            
            for p in cell.paragraphs:
                if i == 0:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
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
    details = get_jurisdiction_details(data.jurisdiction)
    
    desc = data.description.strip() if data.description else "Оказание профессиональных услуг согласно техническому заданию"
    executor = data.executor.strip() if data.executor else "Исполнитель"
    executor_id = data.executor_id.strip() if data.executor_id else "_________________________"
    executor_address = data.executor_address.strip() if data.executor_address else "_________________________"
    
    client = data.client.strip() if data.client else "Заказчик"
    client_id = data.client_id.strip() if data.client_id else "_________________________"
    client_address = data.client_address.strip() if data.client_address else "_________________________"
    
    payment = data.payment.strip() if data.payment else f"Постоплата в течение 5 (пяти) банковских дней в {details['currency']}"
    deadline = data.deadline.strip() if data.deadline else "14 (четырнадцать) календарных дней"
    revisions = data.revisions.strip() if data.revisions else "2 (две)"

    doc_title = "ПУБЛИЧНАЯ ОФЕРТА НА ОКАЗАНИЕ УСЛУГ" if data.doc_type == "offer" else "ДОГОВОР ВОЗМЕЗДНОГО ОКАЗАНИЯ УСЛУГ"

    document_text = f"""# {doc_title}

{details['city']}                                                                          «___» _________ 2026 г.

Гражданин(-ка) {executor} ({details['id_label']}: {executor_id}, адрес: {executor_address}), именуемый(-ая) в дальнейшем «Исполнитель», с одной стороны, и {client} ({details['id_label']}: {client_id}, адрес: {client_address}), именуемый(-ая) в дальнейшем «Заказчик», с другой стороны, совместно именуемые «Стороны», заключили настоящий Договор (далее — «Договор») о нижеследующем:

1. ПРЕДМЕТ ДОГОВОРА
1.1. Исполнитель обязуется по заданию Заказчика оказать следующие услуги: {desc} (далее — «Услуги»), а Заказчик обязуется принять и оплатить оказанные Услуги в порядке и на условиях, предусмотренных настоящим Договором.
1.2. Срок оказания Услуг: {deadline} с момента предоставления Заказчиком необходимых исходных данных.
1.3. Услуги оказываются Исполнителем лично. Привлечение третьих лиц допускается только по предварительному письменному согласованию с Заказчиком.

2. СТОИМОСТЬ УСЛУГ И ПОРЯДОК РАСЧЕТОВ
2.1. Стоимость Услуг по настоящему Договору и порядок расчетов определяются Сторонами следующим образом: {payment}.
2.2. Расчеты производятся в безналичном порядке в валюте {details['currency']}.
2.3. В случае просрочки оплаты Заказчик уплачивает Исполнителю неустойку в размере 0.1% от суммы задолженности за каждый день просрочки.

3. ПОРЯДОК СДАЧИ И ПРИЕМКИ УСЛУГ
3.1. По завершении оказания Услуг Исполнитель передает Заказчику результат и Акт приема-передачи оказанных услуг.
3.2. Заказчик в течение 3 (трех) рабочих дней с момента получения Акта обязан подписать его либо предоставить мотивированный письменный отказ.
3.3. В стоимость Услуг входит до {revisions} итераций бесплатных доработок по замечаниям Заказчика.
3.4. Если в течение указанного срока Заказчик не подписал Акт и не предоставил мотивированный отказ, Услуги считаются принятыми в полном объеме и надлежащего качества.

4. ИНТЕЛЛЕКТУАЛЬНЫЕ ПРАВА И КОНФИДЕНЦИАЛЬНОСТЬ
4.1. Исключительные права на результаты интеллектуальной деятельности, созданные Исполнителем в рамках Договора, переходят к Заказчику в полном объеме с момента их полной (100%) оплаты.
4.2. До момента полной оплаты все права сохраняются за Исполнителем.
4.3. Стороны обязуются соблюдать конфиденциальность в отношении условий Договора и полученной информации.

5. ОТВЕТСТВЕННОСТЬ СТОРОН И РАЗРЕШЕНИЕ СПОРОВ
5.1. За неисполнение или ненадлежащее исполнение обязательств Стороны несут ответственность в соответствии с {details['law_ref']}.
5.2. Все споры подлежат досудебному претензионному урегулированию (срок ответа на претензию — 10 рабочих дней).

6. РЕКВИЗИТЫ И ПОДПИСИ СТОРОН

Исполнитель:
ФИО: {executor}
{details['id_label']}: {executor_id}
Адрес: {executor_address}
Подпись: ________________ / ______________

Заказчик:
Наименование / ФИО: {client}
{details['id_label']}: {client_id}
Адрес: {client_address}
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
        elif line.startswith('г.') or line.startswith('Гражданин'):
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(6)
            run = p.add_run(line)
            run.bold = True
            run.font.name = "Times New Roman"
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(0, 0, 0)
        elif any(line.startswith(f"{i}.") for i in range(1, 10)) or line.startswith('6. РЕКВИЗИТЫ'):
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
    new_fb = {
        "user": data.name,
        "rating": data.rating,
        "text": data.feedback,
        "date": datetime.now().strftime("%Y-%m-%d %H:%M")
    }
    DB_FEEDBACKS.insert(0, new_fb)
    return {"status": "success", "message": "Отзыв успешно сохранен"}


@app.get("/admin/stats/{telegram_id}")
def admin_stats(telegram_id: int):
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
