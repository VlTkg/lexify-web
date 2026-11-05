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
        if line.startswith('# '):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(12)
            run = p.add_run(line.replace('# ', ''))
            run.bold = True
            run.font.name = "Times New Roman"
            run.font.size = Pt(13)
            run.font.color.rgb = RGBColor(0, 0, 0) # Четкий черный цвет
        elif line.startswith('1.') or line.startswith('2.') or line.startswith('3.') or line.startswith('4.') or line.startswith('5.') or line.startswith('6.') or line.startswith('7.'):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run(line)
            run.bold = True
            run.font.name = "Times New Roman"
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(0, 0, 0) # Четкий черный цвет
        else:
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run(line)
            run.font.name = "Times New Roman"
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(0, 0, 0) # Четкий черный цвет

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)

    filename = "Oferta.docx" if data.doc_type == "offer" else "Dogovor.docx"
    return Response(
        content=file_stream.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )
