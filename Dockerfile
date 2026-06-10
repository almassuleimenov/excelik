FROM python:3.11-slim

WORKDIR /app

# Установка системных зависимостей
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# Копирование и установка зависимостей Python
COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir -r requirements.txt

# Копирование исходного кода
COPY main.py .

# Render автоматически пробрасывает переменную PORT (обычно 10000).
# Используем shell-вызов, чтобы uvicorn слушал именно тот порт, который требует платформа.
# Если переменная PORT не задана, дефолтом останется 8000.
CMD uvicorn main:app --host 0.0.0.0 --port ${PORT:-8000}