def run_gemini_analysis(contract_text: str) -> Dict[str, Any]:
    if not GEMINI_API_KEY:
        raise HTTPException(status_code=500, detail="Переменная GEMINI_API_KEY не задана на сервере.")
    
    prompt = f"Проанализируй текст договора:\n\n{contract_text}"
    
    # Используем модель gemini-3.8-flash, указанную Google API
    models_to_try = ["gemini-3.8-flash", "gemini-2.5-flash"]
    last_exception = None

    for model_name in models_to_try:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=prompt,
                config=genai_types.GenerateContentConfig(
                    system_instruction=SYSTEM_PROMPT,
                    response_mime_type="application/json",
                    temperature=0.2,
                ),
            )
            return json.loads(response.text)
        except Exception as e:
            last_exception = e
            continue

    raise HTTPException(status_code=500, detail=f"Ошибка AI API: {str(last_exception)}")
