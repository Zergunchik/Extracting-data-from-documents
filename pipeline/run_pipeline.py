#!/usr/bin/env python3
# run_pipeline.py - Оркестратор конвейера обработки спецификаций
"""
Скрипт для запуска полного конвейера обработки PDF файлов.
Объединяет все 3 этапа в единую последовательность:
1. Распознавание PDF и извлечение данных
2. Извлечение автоматических выключателей
3. Извлечение реле

Использование:
    python run_pipeline.py <путь_к_PDF> [выходная_директория]
    
Или через импорт:
    from pipeline.run_pipeline import run_full_pipeline
    result = run_full_pipeline("specification.pdf", output_dir="./output")
"""

import sys
import os
from pathlib import Path
from typing import Dict, Optional, Any

# Добавляем корневую директорию в путь для импорта cache_manager
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pipeline.stage1_extract_pdf import extract_pdf_to_excel
from pipeline.stage2_extract_breakers import extract_breakers_from_excel
from pipeline.stage3_extract_relays import extract_relays_from_excel


def run_full_pipeline(
    pdf_path: str,
    output_dir: Optional[str] = None,
    use_cache: bool = True,
    skip_voltage_dialog: bool = False
) -> Dict[str, Any]:
    """
    Запускает полный конвейер обработки PDF файла.
    
    Args:
        pdf_path: Путь к PDF файлу
        output_dir: Директория для сохранения результатов (опционально)
        use_cache: Использовать кэширование промежуточных данных
        skip_voltage_dialog: Пропустить GUI диалог выбора напряжения
        
    Returns:
        Словарь с результатами:
        {
            'success': bool,
            'stages': {
                'stage1': {'success': bool, 'output': str или None},
                'stage2': {'success': bool, 'output': str или None, 'count': int},
                'stage3': {'success': bool, 'output': str или None, 'count': int}
            },
            'errors': list
        }
    """
    result = {
        'success': False,
        'stages': {
            'stage1': {'success': False, 'output': None},
            'stage2': {'success': False, 'output': None, 'count': 0},
            'stage3': {'success': False, 'output': None, 'count': 0}
        },
        'errors': []
    }
    
    pdf_file = Path(pdf_path)
    
    if not pdf_file.exists():
        error_msg = f"Файл '{pdf_path}' не найден!"
        print(f"❌ {error_msg}")
        result['errors'].append(error_msg)
        return result
    
    # Инициализируем менеджер кэша
    cache_manager = None
    if use_cache:
        try:
            from cache_manager import init_cache_manager
            cache_manager = init_cache_manager()
            print("✅ Менеджер кэша инициализирован")
        except Exception as e:
            print(f"⚠️ Не удалось инициализировать менеджер кэша: {e}")
            cache_manager = None
    
    # ==========================================
    # ЭТАП 1: Распознавание PDF и извлечение данных
    # ==========================================
    print("\n" + "=" * 60)
    print("ЭТАП 1/3: Распознавание PDF и извлечение данных")
    print("=" * 60)
    
    temp_excel = extract_pdf_to_excel(
        pdf_path=str(pdf_path),
        output_dir=output_dir,
        cache_manager=cache_manager
    )
    
    if not temp_excel:
        error_msg = "Ошибка выполнения этапа 1: не удалось извлечь данные из PDF"
        result['errors'].append(error_msg)
        return result
    
    result['stages']['stage1']['success'] = True
    result['stages']['stage1']['output'] = temp_excel
    print(f"✅ Этап 1 завершен: {temp_excel}")
    
    # ==========================================
    # ЭТАП 2: Извлечение автоматических выключателей
    # ==========================================
    print("\n" + "=" * 60)
    print("ЭТАП 2/3: Извлечение автоматических выключателей")
    print("=" * 60)
    
    breakers, breakers_file = extract_breakers_from_excel(
        input_excel_path=temp_excel,
        output_dir=output_dir,
        apply_nominals=True,
        pdf_name=pdf_file.stem
    )
    
    result['stages']['stage2']['success'] = True
    result['stages']['stage2']['output'] = breakers_file
    result['stages']['stage2']['count'] = len(breakers)
    print(f"✅ Этап 2 завершен: найдено {len(breakers)} автоматов")
    
    # ==========================================
    # ЭТАП 3: Извлечение реле
    # ==========================================
    print("\n" + "=" * 60)
    print("ЭТАП 3/3: Извлечение реле")
    print("=" * 60)
    
    relays, relays_file = extract_relays_from_excel(
        input_excel_path=temp_excel,
        output_dir=output_dir,
        pdf_name=pdf_file.stem,
        skip_voltage_dialog=skip_voltage_dialog
    )
    
    result['stages']['stage3']['success'] = True
    result['stages']['stage3']['output'] = relays_file
    result['stages']['stage3']['count'] = len(relays)
    print(f"✅ Этап 3 завершен: найдено {len(relays)} реле")
    
    # ==========================================
    # ИТОГОВЫЙ ОТЧЕТ
    # ==========================================
    print("\n" + "=" * 60)
    print("ИТОГОВЫЙ ОТЧЕТ")
    print("=" * 60)
    print(f"PDF файл: {pdf_file.name}")
    print(f"Автоматы: {len(breakers)} шт." + (f" → {breakers_file}" if breakers_file else ""))
    print(f"Реле: {len(relays)} шт." + (f" → {relays_file}" if relays_file else ""))
    
    result['success'] = True
    return result


def main():
    """Главная функция для запуска из командной строки"""
    if len(sys.argv) < 2:
        print("Использование: python run_pipeline.py <путь_к_PDF> [выходная_директория]")
        print("\nОпции:")
        print("  --no-cache     Не использовать кэширование")
        print("  --no-gui       Пропустить GUI диалог выбора напряжения")
        print("  --help         Показать эту справку")
        sys.exit(0)
    
    pdf_path = sys.argv[1]
    
    # Проверка на --help или флаги вместо пути к файлу
    if pdf_path in ('--help', '-h', '--from-gui', '--no-cache', '--no-gui'):
        if pdf_path in ('--help', '-h'):
            print("Использование: python run_pipeline.py <путь_к_PDF> [выходная_директория]")
            print("\nОпции:")
            print("  --no-cache     Не использовать кэширование")
            print("  --no-gui       Пропустить GUI диалог выбора напряжения")
            print("  --help         Показать эту справку")
        else:
            print("Ошибка: требуется путь к PDF файлу")
            print("Использование: python run_pipeline.py <путь_к_PDF> [выходная_директория]")
        sys.exit(1 if pdf_path != '--help' else 0)
    
    output_dir = None
    use_cache = True
    skip_voltage_dialog = False
    
    # Парсим аргументы
    i = 2
    while i < len(sys.argv):
        arg = sys.argv[i]
        if arg == '--no-cache':
            use_cache = False
        elif arg == '--no-gui':
            skip_voltage_dialog = True
        elif arg.startswith('--output-dir='):
            output_dir = arg.split('=', 1)[1]
        elif arg == '--output-dir' and i + 1 < len(sys.argv):
            output_dir = sys.argv[i + 1]
            i += 1
        elif arg == '--from-gui':
            # Этот флаг игнорируется, но принимается для совместимости с GUI
            pass
        i += 1
    
    # Запускаем конвейер
    result = run_full_pipeline(
        pdf_path=pdf_path,
        output_dir=output_dir,
        use_cache=use_cache,
        skip_voltage_dialog=skip_voltage_dialog
    )
    
    if result['success']:
        print("\n✅ Конвейер успешно завершен!")
        sys.exit(0)
    else:
        print("\n❌ Конвейер завершился с ошибками:")
        for error in result['errors']:
            print(f"   - {error}")
        sys.exit(1)


if __name__ == "__main__":
    main()
