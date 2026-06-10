# ==========================================
# STAGE 1: Сборка статического бинарника Go
# ==========================================
FROM golang:1.22-alpine AS builder

# Устанавливаем необходимые системные инструменты для сборки (если потребуются)
RUN apk add --no-cache git ca-certificates && update-ca-certificates

WORKDIR /build

# Сначала копируем файлы зависимостей для эффективного кэширования слоев Docker
COPY go.mod go.sum* ./
RUN go mod download

# Копируем исходный код приложения
COPY main.go .

# Компилируем оптимизированный статический бинарник под Linux
# -ldsflags="-s -w" убирает отладочную информацию для минимизации размера
RUN CGO_ENABLED=0 GOOS=linux GOARCH=amd64 go build \
    -ldflags="-s -w" \
    -o excel-comparator main.go

# ==========================================
# STAGE 2: Минимальный образ для рантайма
# ==========================================
FROM alpine:3.19

# Настройка часового пояса и сертификатов безопасности
RUN apk add --no-cache ca-certificates tzdata

# Создаем безопасного non-root пользователя
RUN adduser --disabled-password --gecos "" --home "/nonexistent" --shell "/sbin/nologin" --no-create-home --uid 10001 appuser

WORKDIR /app

# Копируем скомпилированный бинарник из предыдущей стадии сборщика
COPY --from=builder /build/excel-comparator .

# Переключаемся на безопасного пользователя
USER appuser

# Открываем порт для документации (Render сам назначит динамический порт через переменную)
EXPOSE 8080

# Запускаем бинарный файл напрямую (Exec-форма позволяет корректно пробрасывать сигналы ОС)
CMD ["./excel-comparator"]