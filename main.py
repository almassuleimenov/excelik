import io
import pandas as pd
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse

app = FastAPI(
    title="Excel Symmetric Difference API",
    description="API для двустороннего сравнения Excel файлов по ключевой колонке"
)

# Настройка CORS для интеграции с Vercel
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # В продакшене замени ["*"] на конкретный URL твоего фронтенда на Vercel
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/api/v1/compare")
async def compare_excel_files(
    file1: UploadFile = File(...),
    file2: UploadFile = File(...),
    id_column: str = Form("ID")
):
    # 1. Валидация форматов файлов
    for file in [file1, file2]:
        if not file.filename.endswith(('.xlsx', '.xls')):
            raise HTTPException(
                status_code=400, 
                detail=f"Неверный формат файла {file.filename}. Допускаются только .xlsx или .xls"
            )

    try:
        # 2. Чтение файлов напрямую из потока байтов в память
        file1_content = await file1.read()
        file2_content = await file2.read()
        
        df1 = pd.read_excel(io.BytesIO(file1_content))
        df2 = pd.read_excel(io.BytesIO(file2_content))
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"Ошибка парсинга Excel файлов: {str(e)}")

    # 3. Проверка наличия целевой колонки
    if id_column not in df1.columns or id_column not in df2.columns:
        raise HTTPException(
            status_code=400,
            detail=f"Колонка '{id_column}' не найдена в одном из файлов. Проверьте заголовки."
        )

    # 4. Нормализация данных для точного сравнения
    df1['Normalized_ID'] = df1[id_column].dropna().astype(str).str.strip()
    df2['Normalized_ID'] = df2[id_column].dropna().astype(str).str.strip()

    # 5. Реализация алгоритма симметрической разности через Hash Set
    set_a = set(df1['Normalized_ID'])
    set_b = set(df2['Normalized_ID'])

    only_in_df1 = df1[~df1['Normalized_ID'].isin(set_b)].copy()
    only_in_df1['Discrepancy_Source'] = 'Только в Файле 1'

    only_in_df2 = df2[~df2['Normalized_ID'].isin(set_a)].copy()
    only_in_df2['Discrepancy_Source'] = 'Только в Файле 2'

    # Объединяем расхождения
    final_df = pd.concat([only_in_df1, only_in_df2], ignore_index=True)
    final_df.drop(columns=['Normalized_ID'], inplace=True)

    # 6. Запись результата в байтовый буфер памяти (без сохранения на диск)
    output_buffer = io.BytesIO()
    try:
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False)
        output_buffer.seek(0)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка генерации итогового Excel: {str(e)}")

    # 7. Стриминг файла обратно клиенту
    return StreamingResponse(
        output_buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=discrepancies_report.xlsx"}
    )