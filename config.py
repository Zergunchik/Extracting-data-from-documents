import os
from pathlib import Path

# Настройки внешнего вида
APPEARANCE_MODE = "dark"
COLOR_THEME = "green"

# Настройки окна
WINDOW_TITLE = "Извлечение данных для ТО"
WINDOW_GEOMETRY = "1200x800"
WINDOW_MIN_SIZE = (1000, 600)

# Имена папок и файлов
DEFAULT_RESULTS_FOLDER_NAME = "Результаты"
TEMPLATE_FILENAME = "Шаблон для автоматов.xlsx"
MERGED_TEMPLATE_FILENAME = "Объединенный_шаблон_автоматов.xlsx"

# Шаблоны для поиска файла КРУС-КТС
KTC_FILE_PATTERNS = ["*КТС*.pdf", "*ктс*.pdf", "*Ктс*.pdf"]

# Имена файлов
LIBRARY_FILENAME = "Library_nominals.xlsx"

# Расширения файлов
ALLOWED_XLSX_EXT = ".xlsx"
ALLOWED_PDF_EXT = ".pdf"
ALLOWED_EXTENSIONS = [ALLOWED_XLSX_EXT, ALLOWED_PDF_EXT]

# Имена скриптов
SCRIPT_NAMES = {
    "breakers": "extract_circuit_breaker.py",
    "transformers": "extract_current_transformer.py",
    "relays_contactors": "extract_contactors_with_relays.py",
    "basket_breakers": "extract_circuit_breaker_from_baskets.py",
    "secondary_breakers": "run_pipeline.py",
    "relays_pdf": "run_pipeline.py",
    "merge_template": "merge_circuit_breaker_QF.py",
}

# Операции и их скрипты
OPERATION_SCRIPT_MAP = {
    "breakers": SCRIPT_NAMES["breakers"],
    "transformers": SCRIPT_NAMES["transformers"],
    "relays": SCRIPT_NAMES["relays_contactors"],
    "basket_breakers": SCRIPT_NAMES["basket_breakers"],
    "secondary_breakers": SCRIPT_NAMES["secondary_breakers"],
    "relays_pdf": SCRIPT_NAMES["relays_pdf"],
}