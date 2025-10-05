#!/bin/bash

# Скрипт для быстрого запуска приложения

echo "🚀 Starting Excel Transformer..."

# Создание необходимых директорий
mkdir -p uploads output

# Проверка наличия Docker
if ! command -v docker &> /dev/null; then
    echo "❌ Docker не установлен. Пожалуйста, установите Docker."
    exit 1
fi

if ! command -v docker-compose &> /dev/null; then
    echo "❌ Docker Compose не установлен. Пожалуйста, установите Docker Compose."
    exit 1
fi

# Запуск через Docker Compose
echo "📦 Building and starting containers..."
docker-compose up -d --build

# Проверка статуса
if [ $? -eq 0 ]; then
    echo "✅ Application started successfully!"
    echo "🌐 Open http://localhost:8080 in your browser"
    echo ""
    echo "📋 Useful commands:"
    echo "  - View logs: docker-compose logs -f"
    echo "  - Stop app: docker-compose down"
    echo "  - Restart: docker-compose restart"
else
    echo "❌ Failed to start application"
    exit 1
fi
