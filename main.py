@app.post("/generate-protocol")
async def generate_protocol(data: ProtocolRequest):
    doc = docx.Document()

    # Оптимизированные поля страницы (по 1.5 см) для идеальной ширины таблицы без разъезжания в LibreOffice
    for section in doc.sections:
        section.top_margin = Inches(0.59)
        section.bottom_margin = Inches(0.59)
        section.left_margin = Inches(0.59)
        section.right_margin = Inches(0.59)

    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(10)
    normal_style.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run("ПРОТОКОЛ РАЗНОГЛАСИЙ")
    title_run.bold = True
    title_run.font.size = Pt(15)
    title_run.font.color.rgb = RGBColor(0x11, 0x18, 0x27)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = sub_p.add_run(f"к проекту договора ({clean_text(data.doc_type)})\n")
    sub_run.font.size = Pt(11)
    sub_run.italic = True
    sub_run.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)

    # Слово «законодательство» в именительном падеже
    law_p = doc.add_paragraph()
    law_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    law_run = law_p.add_run(f"Применимое право: законодательство ({clean_text(data.jurisdiction)})\n")
    law_run.font.size = Pt(10)
    law_run.font.color.rgb = RGBColor(0x6B, 0x72, 0x80)

    table = doc.add_table(rows=1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'

    col_widths = [Inches(1.2), Inches(2.3), Inches(2.5), Inches(2.0)]

    hdr_cells = table.rows[0].cells
    headers = ["Пункт", "Редакция Заказчика", "Предлагаемая редакция Исполнителя", "Обоснование"]

    for i, header_text in enumerate(headers):
        cell = hdr_cells[i]
        cell.width = col_widths[i]
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        set_cell_background(cell, "1F2937")
        set_cell_margins(cell, top=100, bottom=100, left=100, right=100)

        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(header_text)
        run.bold = True
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    if not data.risks:
        row_cells = table.add_row().cells
        for i in range(4):
            row_cells[i].width = col_widths[i]
        row_cells[0].merge(row_cells[3])
        row_cells[0].text = "Существенных рисков и замечаний к редакции договора не выявлено."
    else:
        for risk in data.risks:
            row_cells = table.add_row().cells

            for i in range(4):
                row_cells[i].width = col_widths[i]
                set_cell_margins(row_cells[i], top=80, bottom=80, left=80, right=80)

            row_cells[0].text = clean_text(risk.title) or "—"
            if row_cells[0].paragraphs[0].runs:
                row_cells[0].paragraphs[0].runs[0].bold = True
                row_cells[0].paragraphs[0].runs[0].font.size = Pt(9.5)

            row_cells[1].text = clean_text(risk.quote) or "—"
            if row_cells[1].paragraphs[0].runs:
                row_cells[1].paragraphs[0].runs[0].font.size = Pt(9.5)

            row_cells[2].text = clean_text(risk.recommendation) or "—"
            if row_cells[2].paragraphs[0].runs:
                row_cells[2].paragraphs[0].runs[0].font.size = Pt(9.5)

            row_cells[3].text = clean_text(risk.explanation) or "Защита интересов исполнителя"
            if row_cells[3].paragraphs[0].runs:
                row_cells[3].paragraphs[0].runs[0].font.size = Pt(9.5)

    doc.add_paragraph("\n")
    doc.add_paragraph("Настоящий Протокол разногласий является неотъемлемой частью Договора.\n")

    sig_table = doc.add_table(rows=1, cols=2)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    cell_left = sig_table.rows[0].cells[0]
    cell_right = sig_table.rows[0].cells[1]

    cell_left.width = Inches(3.8)
    cell_right.width = Inches(3.8)

    cell_left.text = "ПОДРЯДЧИК / ИСПОЛНИТЕЛЬ:\n\n\n_______________________ / ____________ /\nМ.П."
    cell_right.text = "ЗАКАЗЧИК:\n\n\n_______________________ / ____________ /\nМ.П."

    file_path = os.path.join(tempfile.gettempdir(), "Protocol_Raznogasiy.docx")
    doc.save(file_path)

    return FileResponse(
        file_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename="Protocol_Raznogasiy.docx"
    )
