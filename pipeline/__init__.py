"""
Pipeline для обработки спецификаций из PDF файлов.

Модуль содержит функции для разделения процесса обработки на 3 этапа:
1. Распознавание PDF и извлечение данных
2. Извлечение автоматических выключателей
3. Извлечение реле
"""

from .stage1_extract_pdf import extract_pdf_to_excel
from .stage2_extract_breakers import extract_breakers_from_excel
from .stage3_extract_relays import extract_relays_from_excel

__all__ = [
    'extract_pdf_to_excel',
    'extract_breakers_from_excel',
    'extract_relays_from_excel'
]
