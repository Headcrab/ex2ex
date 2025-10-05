# Руководство по настройке конфигурации

## Основы

Конфигурация приложения находится в файле `config.yaml`. Этот файл определяет:
1. Имя выходного файла
2. Правила маппинга данных
3. Настройки листов результирующего файла

## Структура конфигурации

### 1. Имя выходного файла

```yaml
output_filename: "result.xlsx"
```

Это имя будет использоваться для результирующего файла. К нему автоматически добавляется timestamp.

**Пример:** `20231005_120000_result.xlsx`

### 2. Правила маппинга (mappings)

Каждое правило маппинга определяет, откуда брать данные и куда их помещать.

#### Формат ссылки на ячейку/диапазон

```
ИмяЛиста!Ячейка
```

**Примеры:**
- `Sheet1!A1` - ячейка A1 на листе Sheet1
- `Data!B5:D10` - диапазон B5:D10 на листе Data
- `Summary!Z100` - ячейка Z100 на листе Summary

#### Типы маппинга

##### Одна ячейка → Одна ячейка

```yaml
mappings:
  - source: "Sheet1!A1"
    destination: "Result!B2"
```

**Что происходит:** Значение из ячейки A1 листа Sheet1 копируется в ячейку B2 листа Result.

##### Диапазон → Диапазон

```yaml
mappings:
  - source: "Sheet1!A1:C10"
    destination: "Result!D1"
```

**Что происходит:** Диапазон ячеек A1:C10 (3 колонки × 10 строк) копируется начиная с ячейки D1 листа Result.

### 3. Настройки выходных листов

```yaml
output_sheets:
  - name: "Result"
    create_if_not_exists: true
```

**Параметры:**
- `name` - имя листа в результирующем файле
- `create_if_not_exists` - создать лист, если его нет (true/false)

## Примеры сценариев использования

### Сценарий 1: Копирование отчета

**Задача:** Скопировать весь отчет с одного листа на другой.

```yaml
output_filename: "report_copy.xlsx"

mappings:
  - source: "Report!A1:Z100"
    destination: "Copy!A1"

output_sheets:
  - name: "Copy"
    create_if_not_exists: true
```

### Сценарий 2: Объединение данных из разных листов

**Задача:** Собрать данные из трех разных листов в один.

```yaml
output_filename: "combined_data.xlsx"

mappings:
  # Заголовки (берем из первого листа)
  - source: "January!A1:E1"
    destination: "Combined!A1"
  
  # Данные за январь
  - source: "January!A2:E50"
    destination: "Combined!A2"
  
  # Данные за февраль (начинаем с строки 52)
  - source: "February!A2:E50"
    destination: "Combined!A52"
  
  # Данные за март (начинаем с строки 102)
  - source: "March!A2:E50"
    destination: "Combined!A102"

output_sheets:
  - name: "Combined"
    create_if_not_exists: true
```

### Сценарий 3: Изменение структуры таблицы

**Задача:** Переставить колонки местами.

```yaml
output_filename: "restructured.xlsx"

mappings:
  # Колонка A → Колонка C
  - source: "Source!A1:A100"
    destination: "Target!C1"
  
  # Колонка B → Колонка A
  - source: "Source!B1:B100"
    destination: "Target!A1"
  
  # Колонка C → Колонка B
  - source: "Source!C1:C100"
    destination: "Target!B1"

output_sheets:
  - name: "Target"
    create_if_not_exists: true
```

### Сценарий 4: Создание сводной таблицы

**Задача:** Извлечь ключевые показатели из разных мест исходного файла.

```yaml
output_filename: "summary.xlsx"

mappings:
  # Заголовок отчета
  - source: "Dashboard!A1"
    destination: "Summary!A1"
  
  # Финансовые показатели
  - source: "Finance!B5"     # Выручка
    destination: "Summary!B3"
  
  - source: "Finance!B6"     # Прибыль
    destination: "Summary!B4"
  
  - source: "Finance!B7"     # ROI
    destination: "Summary!B5"
  
  # Операционные показатели
  - source: "Operations!C10:C15"
    destination: "Summary!B8"
  
  # Детальная таблица
  - source: "Details!A1:F50"
    destination: "Summary!A15"

output_sheets:
  - name: "Summary"
    create_if_not_exists: true
```

### Сценарий 5: Создание нескольких выходных листов

**Задача:** Разделить данные из одного листа на несколько.

```yaml
output_filename: "separated_data.xlsx"

mappings:
  # В первый лист - первая половина данных
  - source: "AllData!A1:Z50"
    destination: "Part1!A1"
  
  # Во второй лист - вторая половина данных
  - source: "AllData!A51:Z100"
    destination: "Part2!A1"
  
  # В третий лист - заголовки и итоги
  - source: "AllData!A1:Z1"
    destination: "Summary!A1"
  
  - source: "AllData!A101:Z105"
    destination: "Summary!A3"

output_sheets:
  - name: "Part1"
    create_if_not_exists: true
  - name: "Part2"
    create_if_not_exists: true
  - name: "Summary"
    create_if_not_exists: true
```

## Советы и рекомендации

### 1. Именование листов

- Используйте понятные имена листов
- Избегайте специальных символов: `\ / ? * [ ]`
- Имена листов чувствительны к регистру

### 2. Координаты ячеек

- Всегда используйте заглавные буквы для столбцов: `A1`, не `a1`
- Проверяйте правильность диапазонов
- Формат диапазона: `A1:Z100`, не `A1-Z100`

### 3. Порядок маппингов

Маппинги выполняются последовательно, в том порядке, в котором они указаны в конфигурации.

### 4. Обработка ошибок

Если маппинг не может быть выполнен (например, исходный лист не существует), приложение:
- Записывает предупреждение в лог
- Продолжает выполнение остальных маппингов
- Создает результирующий файл с данными, которые удалось обработать

### 5. Копирование форматирования

Приложение пытается скопировать не только значения, но и форматирование ячеек (цвет, шрифт, границы).

### 6. Тестирование конфигурации

Перед запуском на важных данных:
1. Создайте тестовый Excel файл
2. Настройте маппинги
3. Проверьте результат
4. Откорректируйте конфигурацию при необходимости

## Отладка

### Просмотр логов

**Docker:**
```bash
docker-compose logs -f
```

**Локально:**
Логи выводятся в консоль

### Типичные ошибки

#### Ошибка: "Sheet not found"
**Причина:** Указанный лист не существует в исходном файле  
**Решение:** Проверьте имя листа (с учетом регистра)

#### Ошибка: "Invalid cell reference"
**Причина:** Неправильный формат ссылки на ячейку  
**Решение:** Используйте формат `Sheet!A1` или `Sheet!A1:B10`

#### Ошибка: "Failed to parse range"
**Причина:** Неправильный формат диапазона  
**Решение:** Убедитесь, что используете формат `A1:B10`

## Дополнительные возможности

### Переменные окружения

Путь к конфигурации можно изменить через переменную окружения:

```bash
CONFIG_FILE=/path/to/custom/config.yaml
```

В Docker Compose:
```yaml
environment:
  - CONFIG_FILE=/app/custom-config.yaml
volumes:
  - ./custom-config.yaml:/app/custom-config.yaml
```

### Множественные конфигурации

Вы можете создать несколько конфигурационных файлов для разных сценариев:
- `config.yaml` - основная конфигурация
- `config-monthly.yaml` - месячные отчеты
- `config-quarterly.yaml` - квартальные отчеты

Переключайте их через переменную окружения или монтирование в Docker.

## Примеры готовых конфигураций

Смотрите файл `config.examples.yaml` для готовых примеров конфигураций под различные сценарии.
