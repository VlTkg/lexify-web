import os
import httpx
from fastapi import FastAPI, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pydantic import BaseModel

app = FastAPI()

# Монтируем статические файлы (HTML, CSS, JS)
app.mount("/static", StaticFiles(directory="static"), name="static")

# URL вебхука Google Таблицы из переменных окружения Render
FEEDBACK_WEBHOOK_URL = os.environ.get("GOOGLE_SHEET_WEBHOOK_URL", "")

class FeedbackSchema(BaseModel):
    type: str = ""          # "Оценка анализа" или "Баг / Ошибка"
    rating: int = 0         # 1-5 звезд
    sentiment: str = ""     # "Полезно" или "Замечания"
    feedback: str = ""      # Текст отзыва
    doc_type: str = ""      # Тип договора
    jurisdiction: str = ""  # Юрисдикция

@app.get("/")
async def serve_index():
    return FileResponse("static/index.html")

@app.post("/feedback")
async def receive_feedback(data: FeedbackSchema):
    if not FEEDBACK_WEBHOOK_URL:
        print(f"Feedback received (no webhook configured): {data}")
        return {"status": "ok", "message": "Logged locally"}

    try:
        async with httpx.AsyncClient() as client:
            response = await client.post(
                FEEDBACK_WEBHOOK_URL,
                json=data.dict(),
                timeout=10.0
            )
            if response.status_code == 200:
                return {"status": "ok", "message": "Feedback saved to Google Sheet"}
            else:
                raise HTTPException(status_code=500, detail="Failed to post to Google Sheet")
    except Exception as e:
        print(f"Error sending feedback: {e}")
        return {"status": "error", "detail": str(e)}
