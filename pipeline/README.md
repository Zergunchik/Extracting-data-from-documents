# Pipeline обработки спецификаций

Модуль `pipeline` разделяет процесс обработки PDF файлов на 3 независимых этапа для удобства поддержки и расширения кода.

## Структура модуля

```
pipeline/
├── __init__.py                    # Инициализация модуля, экспорт функций
├── stage1_extract_pdf.py          # Этап 1: Распознавание PDF
├── stage2_extract_breakers.py     # Этап 2: Извлечение автоматов
├── stage3_extract_relays.py       # Этап 3: Извлечение реле
└── run_pipeline.py                # Оркестратор полного конвейера
```

## Конвейер обработки

```
PDF файл 
   ↓
[Этап 1: stage1_extract_pdf.py] → {имя}_extracted_temp.xlsx
   ↓
[Этап 2: stage2_extract_breakers.py] → {имя}_Автоматические_выключатели.xlsx
   ↓
[Этап 3: stage3_extract_relays.py] → {имя}_Реле.xlsx
```

## Использование

### Через командную строку

Запуск полного конвейера:
```bash
python pipeline/run_pipeline.py specification.pdf
python pipeline/run_pipeline.py specification.pdf --output-dir=./output
python pipeline/run_pipeline.py specification.pdf --no-cache --no-gui
```

Запуск отдельных этапов:
```bash
# Этап 1: Извлечение данных из PDF
python pipeline/stage1_extract_pdf.py specification.pdf

# Этап 2: Извлечение автоматов (требует результат этапа 1)
python pipeline/stage2_extract_breakers.py specification_extracted_temp.xlsx

# Этап 3: Извлечение реле (требует результат этапа 1)
python pipeline/stage3_extract_relays.py specification_extracted_temp.xlsx
```

### Через Python API

```python
from pipeline.run_pipeline import run_full_pipeline

# Запуск полного конвейера
result = run_full_pipeline(
    pdf_path="specification.pdf",
    output_dir="./output",
    use_cache=True,
    skip_voltage_dialog=False
)

if result['success']:
    print(f"Автоматы: {result['stages']['stage2']['count']} шт.")
    print(f"Реле: {result['stages']['stage3']['count']} шт.")
    print(f"Файл автоматов: {result['stages']['stage2']['output']}")
    print(f"Файл реле: {result['stages']['stage3']['output']}")
```

Запуск отдельных этапов:
```python
from pipeline.stage1_extract_pdf import extract_pdf_to_excel
from pipeline.stage2_extract_breakers import extract_breakers_from_excel
from pipeline.stage3_extract_relays import extract_relays_from_excel

# Этап 1
temp_excel = extract_pdf_to_excel("specification.pdf", cache_manager=cache_manager)

# Этап 2
breakers, breakers_file = extract_breakers_from_excel(temp_excel, output_dir="./output")

# Этап 3
relays, relays_file = extract_relays_from_excel(temp_excel, output_dir="./output")
```

## Описание этапов

### Этап 1: Распознавание PDF (`stage1_extract_pdf.py`)

**Функции:**
- `extract_text_from_pdf(pdf_path)` - извлекает текст из PDF
- `extract_tables_from_pdf(pdf_path)` - извлекает таблицы из PDF
- `extract_pdf_to_excel(pdf_path, output_dir, cache_manager)` - главная функция этапа

**Входные данные:** Путь к PDF файлу
**Выходные данные:** Путь к Excel файлу с извлеченными данными

**Особенности:**
- Проверка кэша перед обработкой
- Сохранение текста и таблиц в отдельные листы Excel
- Поддержка альтернативного метода извлечения таблиц

### Этап 2: Извлечение автоматов (`stage2_extract_breakers.py`)

**Функции:**
- `is_valid_position(pos)` - проверка корректности позиции
- `expand_positions_breaker(text)` - раскрытие диапазонов позиций
- `extract_device_type(text)` - определение типа автомата
- `get_current_from_type(type_str)` - извлечение номинального тока
- `process_worksheet_breakers(ws)` - обработка листа Excel
- `load_nominals_library(library_path)` - загрузка библиотеки номиналов
- `match_device_type_with_library(device_type, nominals_library)` - сопоставление с библиотекой
- `extract_breakers_from_excel(input_excel_path, ...)` - главная функция этапа

**Входные данные:** Путь к Excel файлу с данными из PDF
**Выходные данные:** Список автоматов и путь к итоговому файлу

**Особенности:**
- Автоматическое применение библиотеки номинальных токов
- Поддержка различных форматов позиций (SF, QFD, SFD)
- Поиск типа устройства с заглядыванием вперед

### Этап 3: Извлечение реле (`stage3_extract_relays.py`)

**Функции:**
- `normalize_text(text)` - нормализация текста
- `has_negative_keywords(text)` - проверка на исключающие слова
- `find_columns(data)` - поиск колонок с позициями и наименованиями
- `expand_positions_relay(text)` - раскрытие диапазонов позиций реле
- `extract_relay_type(text)` - определение типа реле
- `process_worksheet_relays(ws, shield_name)` - обработка листа Excel
- `extract_relays_from_excel(input_excel_path, ...)` - главная функция этапа

**Входные данные:** Путь к Excel файлу с данными из PDF
**Выходные данные:** Список реле и путь к итоговому файлу

**Особенности:**
- Поддержка различных форматов реле (KL, K, KLB, KCC, KLP)
- Использование файла памяти напряжений (Relay_voltage.xlsx)
- Исключение реле напряжения (KV, KVZ)

## Задействованные файлы

- `cache_manager.py` - управление кэшем промежуточных данных
- `Library_nominals.xlsx` - библиотека номинальных токов автоматов
- `Relay_voltage.xlsx` - память напряжений реле

## Преимущества разделения

1. **Модульность** - каждый этап независим и может тестироваться отдельно
2. **Повторное использование** - можно запускать только нужные этапы
3. **Кэширование** - промежуточные результаты сохраняются для ускорения повторной обработки
4. **Расширяемость** - легко добавить новые этапы обработки
5. **Отладка** - проще найти и исправить ошибки в конкретном этапе

## Миграция со старого кода

Старый файл `extract_specification.py` продолжает работать как прежде. Новые скрипты используют тот же функционал, но разделены на логические модули.

Для постепенной миграции рекомендуется:
1. Протестировать новый конвейер на нескольких файлах
2. Сравнить результаты со старым скриптом
3. Обновить GUI и другие точки входа на использование нового API
