import re
from pathlib import Path
from typing import List, Tuple, Set, Optional

import config

def check_ktc_file(folder: Path) -> Tuple[bool, Optional[Path]]:
    """Проверяет наличие файла КРУС-КТС (PDF с KTC в имени) в папке"""
    pdf_files = []
    for pattern in config.KTC_FILE_PATTERNS:
        pdf_files.extend(folder.glob(pattern))
    pdf_files = list(set(pdf_files))
    return (True, pdf_files[0]) if pdf_files else (False, None)

def parse_dnd_data(raw_data: str) -> List[Tuple[str, Path]]:
    """
    Парсит данные из DND события.
    Возвращает список кортежей (тип_файла, путь).
    """
    result = []
    # Регулярка для путей в фигурных скобках (Windows) или без
    matches = re.findall(r'\{([^}]+)\}|(\S+)', raw_data)
    for match in matches:
        path_str = match[0] if match[0] else match[1]
        path_str = path_str.strip()
        if not path_str:
            continue
        p = Path(path_str)
        if p.suffix.lower() == config.ALLOWED_XLSX_EXT:
            result.append(('xlsx', p))
        elif p.suffix.lower() == config.ALLOWED_PDF_EXT:
            result.append(('pdf', p))
    return result

def extract_missing_types(output_text: str) -> Set[str]:
    """Извлекает из вывода скрипта типы автоматов, отсутствующие в библиотеке"""
    missing = set()
    # Примеры шаблонов из оригинального кода
    pattern1 = r"\[тип '([^']+)' не найден в библиотеке\]"
    pattern2 = r"-\s*\[тип '([^']+)' "
    for match in re.findall(pattern1, output_text):
        missing.add(match.strip())
    for match in re.findall(pattern2, output_text):
        clean = re.sub(r'\s*-\s*номинал\s+\d+A\s+найден\s+в\s+PDF', '', match)
        missing.add(clean.strip())
    return missing

def read_process_output(stdout: bytes, stderr: bytes) -> Tuple[str, str]:
    """Декодирует stdout/stderr с попыткой нескольких кодировок"""
    for encoding in ('utf-8', 'windows-1251'):
        try:
            out = stdout.decode(encoding, errors='replace')
            err = stderr.decode(encoding, errors='replace')
            return out, err
        except UnicodeDecodeError:
            continue
    return stdout.decode('utf-8', errors='replace'), stderr.decode('utf-8', errors='replace')

def get_script_path(script_name: str, current_folder: Path) -> Optional[Path]:
    """Возвращает путь к скрипту в текущей папке, папке scripts или pipeline"""
    candidates = [
        current_folder / script_name,
        current_folder / "scripts" / script_name,
        current_folder / "pipeline" / script_name,
    ]
    for p in candidates:
        if p.exists():
            return p
    return None

def copy_to_clipboard(root, text: str):
    """Копирует текст в буфер обмена через корневое окно"""
    try:
        root.clipboard_clear()
        root.clipboard_append(text)
        root.update()
    except Exception:
        pass