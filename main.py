import io
from fastapi import FastAPI, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, StreamingResponse
from pydantic import BaseModel
from docx import Document

app = FastAPI()

# Монтируем статику правильно
app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/")
def read_root():
    return FileResponse("static/index.html")

class GenerateRequest(BaseModel):
    doc_type: str
    parameters: str
    init_data: str = ""
    jurisdiction: str = "Россия (РФ)"

@app.post("/api/generate-document")
def generate_document(req: GenerateRequest):
    try:
        doc = Document()
        
        titles = {
            "contract": "ОСНОВНОЙ ДОГОВОР",
            "addendum": "ДОПОЛНИТЕЛЬНОЕ СОГЛАШЕНИЕ",
            "claim": "ЮРИДИЧЕСКАЯ ПРЕТЕНЗИЯ / ПИСЬМО"
        }
        
        doc.add_heading(titles.get(req.doc_type, "ЮРИДИЧЕСКИЙ ДОКУМЕНТ"), level=1)
        doc.add_paragraph("Настоящий документ сформирован сервисом Lexify на основе предоставленных параметров.")
        doc.add_heading("Параметры и условия:", level=2)
        doc.add_paragraph(req.parameters if req.parameters else "Параметры не указаны пользователем.")
        
        doc.add_heading("Заключительные положения:", level=2)
        doc.add_paragraph("Документ носит рекомендательный характер и подлежит правовой экспертизе.")

        file_stream = io.BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)

        filename_map = {
            "contract": "Contract.docx",
            "addendum": "Addendum.docx",
            "claim": "Claim.docx"
        }
        filename = filename_map.get(req.doc_type, "Result.docx")

        return StreamingResponse(
            file_stream,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/generate-report")
def generate_report(req: GenerateRequest):
    try:
        doc = Document()
        if req.doc_type == "protocol":
            title = "ПРОТОКОЛ РАЗНОГЛАСИЙ"
            filename = "Protocol_Raznoglasiy.docx"
        else:
            title = "ЭКСПЕРТНОЕ ЗАКЛЮЧЕНИЕ ПО РЕЗУЛЬТАТАМ АНАЛИЗА"
            filename = "Lexify_Result_Analysis.docx"

        doc.add_heading(title, level=1)
        doc.add_paragraph(f"Юрисдикция: {req.jurisdiction}")
        doc.add_paragraph(req.parameters if req.parameters else "Результаты правовой проверки не предоставлены.")
        doc.add_paragraph("Документ сформирован сервисом Lexify и подлежит правовой экспертизе.")

        file_stream = io.BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)

        return StreamingResponse(
            file_stream,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/user-info/{telegram_id}")
def user_info(telegram_id: int):
    return {"telegram_id": telegram_id, "free_checks_left": 3, "is_admin": telegram_id == 0}

@app.get("/admin/stats/{telegram_id}")
def admin_stats(telegram_id: int):
    return {
        "total_users": 1,
        "total_analyses": 12,
        "total_generations": 5,
        "total_payments": 0
    }
