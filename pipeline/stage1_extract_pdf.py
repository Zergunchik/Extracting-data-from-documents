# stage1_extract_pdf.py - Этап 1: Распознавание PDF и извлечение данных
"""
Этап 1 конвейера обработки спецификаций.
Извлекает текст и таблицы из PDF файла и сохраняет в промежуточный Excel файл.

Входные данные: путь к PDF файлу
Выходные данные: путь к Excel файлу с извлеченными данными
"""

import pdfplumber
import pandas as pd
import os
from pathlib import Path
from typing import Optional, List, Dict, Any


def extract_text_from_pdf(pdf_path: str) -> Optional[List[Dict[str, Any]]]:
    """
    Извлекает текст из PDF файла.
    
    Args:
        pdf_path: Путь к PDF файлу
        
    Returns:
        Список словарей с данными текста или None при ошибке
    """
    extracted_data = []
    print(f"   📄 Чтение текста из: {os.path.basename(pdf_path)}")
    try:
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            print(f"   📄 Всего страниц: {total_pages}")

            if total_pages == 0:
                print(f"   ⚠️ PDF файл не содержит страниц!")
                return None

            for page_num, page in enumerate(pdf.pages, 1):
                if page_num % 5 == 0 or page_num == total_pages:
                    print(f"   ⏳ Обработка текста: страница {page_num}/{total_pages}")

                text = page.extract_text()
                if text:
                    lines = text.split('\n')
                    for line_num, line in enumerate(lines, 1):
                        if line.strip():
                            extracted_data.append({
                                'Страница': page_num,
                                'Строка': line_num,
                                'Текст': line.strip()
                            })

            print(f"   ✅ Извлечено {len(extracted_data)} строк текста")
            return extracted_data
    except Exception as e:
        print(f"❌ Ошибка при чтении текста PDF: {e}")
        import traceback
        traceback.print_exc()
        return None


def extract_tables_from_pdf(pdf_path: str) -> List[Dict[str, Any]]:
    """
    Извлекает таблицы из PDF файла.
    
    Args:
        pdf_path: Путь к PDF файлу
        
    Returns:
        Список словарей с данными таблиц
    """
    all_tables = []
    print(f"   📊 Чтение таблиц из: {os.path.basename(pdf_path)}")
    try:
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            print(f"   📊 Всего страниц: {total_pages}")

            for page_num, page in enumerate(pdf.pages, 1):
                if page_num % 5 == 0 or page_num == total_pages:
                    print(f"   ⏳ Обработка таблиц: страница {page_num}/{total_pages}")

                try:
                    tables = page.extract_tables()
                    if tables:
                        print(f"   📊 Страница {page_num}: найдено {len(tables)} таблиц")
                    else:
                        tables = page.extract_tables({
                            "vertical_strategy": "lines",
                            "horizontal_strategy": "lines"
                        })
                        if tables:
                            print(f"   📊 Страница {page_num}: найдено {len(tables)} таблиц (альтернативный метод)")

                    for table_num, table in enumerate(tables, 1):
                        if table:
                            clean_table = []
                            for row in table:
                                clean_row = [str(cell).strip() if cell is not None else '' for cell in row]
                                if any(clean_row):
                                    clean_table.append(clean_row)
                            if clean_table:
                                all_tables.append({
                                    'Страница': page_num,
                                    'Таблица': table_num,
                                    'Данные': clean_table
                                })
                                print(f"   📊 Таблица {table_num} на странице {page_num}: {len(clean_table)} строк, {len(clean_table[0]) if clean_table else 0} колонок")
                except Exception as page_e:
                    print(f"   ⚠️ Ошибка на странице {page_num}: {page_e}")
                    continue
    except Exception as e:
        print(f"❌ Ошибка при извлечении таблиц: {e}")
        import traceback
        traceback.print_exc()
        return all_tables

    print(f"   ✅ Всего извлечено {len(all_tables)} таблиц")
    return all_tables


def extract_pdf_to_excel(
    pdf_path: str,
    output_dir: Optional[str] = None,
    cache_manager=None,
    progress_callback=None
) -> Optional[str]:
    """
    Главная функция этапа 1: извлечение данных из PDF и сохранение в Excel.
    
    Args:
        pdf_path: Путь к PDF файлу
        output_dir: Директория для сохранения выходного файла (опционально)
        cache_manager: Менеджер кэша для проверки и сохранения результатов
        progress_callback: Функция обратного вызова для обновления прогресса (percent, message)
        
    Returns:
        Путь к созданному Excel файлу или None при ошибке
    """
    pdf_file = Path(pdf_path)
    
    if not pdf_file.exists():
        print(f"❌ Ошибка: Файл '{pdf_path}' не найден!")
        return None
    
    # Проверяем кэш
    if cache_manager:
        cached_excel = cache_manager.get_cached_excel(pdf_file)
        if cached_excel and cached_excel.exists():
            print(f"✅ Найдены закэшированные данные: {cached_excel.name}")
            if progress_callback:
                progress_callback(15, "Распознавание PDF: данные загружены из кэша")
            return str(cached_excel)
    
    print(f"\n📁 Обработка файла: {pdf_file.name}")
    print("=" * 50)
    print("Шаг 1: Извлечение текста и таблиц из PDF...")
    
    if progress_callback:
        progress_callback(8, "Распознавание PDF: чтение текста...")
    print("PROGRESS:8:Распознавание PDF: чтение текста...")
    
    # Извлекаем данные
    text_data = extract_text_from_pdf(str(pdf_path))
    
    if progress_callback:
        progress_callback(12, "Распознавание PDF: извлечение таблиц...")
    print("PROGRESS:12:Распознавание PDF: извлечение таблиц...")
    
    tables_data = extract_tables_from_pdf(str(pdf_path))
    
    if not text_data and not tables_data:
        print("❌ Не удалось извлечь данные из PDF файла.")
        if progress_callback:
            progress_callback(15, "Ошибка: не удалось извлечь данные из PDF")
        print("PROGRESS:15:Ошибка: не удалось извлечь данные из PDF")
        return None
    
    # Определяем путь для временного файла
    if output_dir:
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
    else:
        output_path = pdf_file.parent
    
    temp_excel_path = output_path / f"{pdf_file.stem}_extracted_temp.xlsx"
    
    # Сохраняем в Excel
    if progress_callback:
        progress_callback(18, "Распознавание PDF: сохранение промежуточных данных...")
    print("PROGRESS:18:Распознавание PDF: сохранение промежуточных данных...")
    
    print(f"   Сохранение промежуточных данных: {temp_excel_path.name}")
    with pd.ExcelWriter(temp_excel_path, engine='openpyxl') as writer:
        if text_data:
            df_text = pd.DataFrame(text_data)
            df_text.to_excel(writer, sheet_name='Текст_из_PDF', index=False)
        if tables_data:
            for i, table_info in enumerate(tables_data):
                sheet_name = f"Таблица_{table_info['Страница']}_{table_info['Таблица']}"
                sheet_name = sheet_name[:31]  # Ограничение Excel на длину имени листа
                df_table = pd.DataFrame(table_info['Данные'])
                df_table.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
    
    # Сохраняем в кэш
    if cache_manager:
        temp_excel_path = cache_manager.save_to_cache(pdf_file, temp_excel_path)
    
    if progress_callback:
        progress_callback(20, "Распознавание PDF: данные сохранены")
    print("PROGRESS:20:Распознавание PDF: данные сохранены")
    
    print(f"✅ Этап 1 завершен. Временный файл: {temp_excel_path.name}")
    return str(temp_excel_path)


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("Использование: python stage1_extract_pdf.py <путь_к_PDF> [выходная_директория]")
        sys.exit(1)
    
    pdf_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None
    
    result = extract_pdf_to_excel(pdf_path, output_dir)
    if result:
        print(f"\n✅ Результат сохранен: {result}")
    else:
        print("\n❌ Ошибка выполнения этапа 1")
        sys.exit(1)
