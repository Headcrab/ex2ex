# PowerShell скрипт для быстрого запуска приложения

Write-Host "🚀 Starting Excel Transformer..." -ForegroundColor Green

# Создание необходимых директорий
if (-not (Test-Path "uploads")) {
    New-Item -ItemType Directory -Path "uploads" | Out-Null
}

if (-not (Test-Path "output")) {
    New-Item -ItemType Directory -Path "output" | Out-Null
}

# Проверка наличия Docker
try {
    $dockerVersion = docker --version
    Write-Host "✓ Docker found: $dockerVersion" -ForegroundColor Gray
} catch {
    Write-Host "❌ Docker не установлен. Пожалуйста, установите Docker Desktop." -ForegroundColor Red
    exit 1
}

# Проверка наличия Docker Compose
try {
    $composeVersion = docker-compose --version
    Write-Host "✓ Docker Compose found: $composeVersion" -ForegroundColor Gray
} catch {
    Write-Host "❌ Docker Compose не установлен." -ForegroundColor Red
    exit 1
}

# Запуск через Docker Compose
Write-Host ""
Write-Host "📦 Building and starting containers..." -ForegroundColor Cyan
docker-compose up -d --build

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "✅ Application started successfully!" -ForegroundColor Green
    Write-Host "🌐 Open http://localhost:8080 in your browser" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "📋 Useful commands:" -ForegroundColor Cyan
    Write-Host "  - View logs: docker-compose logs -f"
    Write-Host "  - Stop app: docker-compose down"
    Write-Host "  - Restart: docker-compose restart"
    Write-Host ""
    
    # Опционально: открыть браузер
    $openBrowser = Read-Host "Открыть браузер? (Y/N)"
    if ($openBrowser -eq "Y" -or $openBrowser -eq "y") {
        Start-Process "http://localhost:8080"
    }
} else {
    Write-Host ""
    Write-Host "❌ Failed to start application" -ForegroundColor Red
    Write-Host "Check logs with: docker-compose logs" -ForegroundColor Yellow
    exit 1
}
