# Скрипт для тестирования API

Write-Host "=== Тестирование API конфигурации ===" -ForegroundColor Cyan
Write-Host ""

# Проверка, что сервер запущен
Write-Host "1. Проверка доступности сервера..." -ForegroundColor Yellow
try {
    $response = Invoke-WebRequest -Uri "http://localhost:8080" -Method GET -TimeoutSec 5 -UseBasicParsing
    Write-Host "   ✓ Сервер доступен (код: $($response.StatusCode))" -ForegroundColor Green
} catch {
    Write-Host "   ✗ Сервер недоступен: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "Запустите сервер командой: go run main.go" -ForegroundColor Yellow
    exit 1
}

Write-Host ""

# Тест GET /api/config
Write-Host "2. Тестирование GET /api/config..." -ForegroundColor Yellow
try {
    $response = Invoke-RestMethod -Uri "http://localhost:8080/api/config" -Method GET -ContentType "application/json"
    Write-Host "   ✓ Конфигурация получена успешно" -ForegroundColor Green
    Write-Host "   - Имя файла: $($response.output_filename)" -ForegroundColor Gray
    Write-Host "   - Правил маппинга: $($response.mappings.Count)" -ForegroundColor Gray
    Write-Host "   - Выходных листов: $($response.output_sheets.Count)" -ForegroundColor Gray
    Write-Host ""
    Write-Host "   Полный ответ:" -ForegroundColor Gray
    $response | ConvertTo-Json -Depth 10 | Write-Host
} catch {
    Write-Host "   ✗ Ошибка: $_" -ForegroundColor Red
    Write-Host "   Детали: $($_.Exception.Response.StatusCode)" -ForegroundColor Red
}

Write-Host ""

# Тест POST /api/config
Write-Host "3. Тестирование POST /api/config (сохранение)..." -ForegroundColor Yellow
$testConfigJson = @"
{
    "output_filename": "test_result.xlsx",
    "mappings": [
        {
            "source": "TestSheet!A1",
            "destination": "ResultSheet!B2"
        }
    ],
    "output_sheets": [
        {
            "name": "ResultSheet",
            "create_if_not_exists": true
        }
    ]
}
"@

try {
    $response = Invoke-RestMethod -Uri "http://localhost:8080/api/config" -Method POST -Body $testConfigJson -ContentType "application/json"
    Write-Host "   ✓ Конфигурация сохранена успешно" -ForegroundColor Green
    Write-Host "   Ответ: $($response | ConvertTo-Json)" -ForegroundColor Gray
} catch {
    Write-Host "   ✗ Ошибка: $_" -ForegroundColor Red
    if ($_.ErrorDetails) {
        Write-Host "   Детали: $($_.ErrorDetails.Message)" -ForegroundColor Red
    }
}

Write-Host ""

# Проверка файла config.yaml
Write-Host "4. Проверка файла config.yaml..." -ForegroundColor Yellow
if (Test-Path "config.yaml") {
    $configContent = Get-Content "config.yaml" -Raw
    Write-Host "   ✓ Файл существует" -ForegroundColor Green
    Write-Host "   Содержимое:" -ForegroundColor Gray
    Write-Host $configContent -ForegroundColor DarkGray
} else {
    Write-Host "   ✗ Файл config.yaml не найден" -ForegroundColor Red
}

Write-Host ""
Write-Host "=== Тестирование завершено ===" -ForegroundColor Cyan
