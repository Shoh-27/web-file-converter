FROM node:18-slim

# Устанавливаем LibreOffice и необходимые зависимости
RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-writer \
    libreoffice-calc \
    libreoffice-impress \
    fonts-liberation \
    fonts-dejavu \
    fonts-noto \
    fonts-freefont-ttf \
    fontconfig \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Создаем рабочую директорию
WORKDIR /app

# Копируем файлы package
COPY package*.json ./

# Устанавливаем зависимости
RUN npm ci --only=production

# Создаем non-root пользователя ДО копирования файлов
RUN useradd -m -u 1001 appuser

# Создаем директории с правильными правами
RUN mkdir -p /tmp/conversions && \
    chmod 777 /tmp/conversions && \
    mkdir -p /app/public && \
    mkdir -p /app/js && \
    mkdir -p /app/style && \
    chmod -R 755 /app

# Копируем все файлы приложения
COPY --chown=appuser:appuser . .

# Устанавливаем права на директории
RUN chown -R appuser:appuser /app && \
    chown -R appuser:appuser /tmp/conversions

# Переключаемся на non-root пользователя
USER appuser

# Открываем порт
EXPOSE 3000

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD node -e "require('http').get('http://localhost:3000/health', (r) => { process.exit(r.statusCode === 200 ? 0 : 1); })"

# Запускаем приложение
CMD ["node", "js/server.js"]