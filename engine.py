import subprocess
import sys
import shutil
import runpy
import os
import io
import time
import glob
import builtins
from pathlib import Path
from typing import List, Tuple, Set, Optional, Callable
from datetime import datetime
import pandas as pd
import config
from utils import get_script_path, read_process_output, extract_missing_types


def _safe_path(path: Path) -> Path:
    """Генерирует уникальный путь, добавляя суффикс (1), (2)... если файл существует."""
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    counter = 1
    while True:
        new_path = parent / f"{stem}_{counter}{suffix}"
        if not new_path.exists():
            return new_path
        counter += 1


def _run_script_as_module(script_name: str, input_file: Path, output_folder: Path,
                          current_folder: Path, merge_mode: bool = False,
                          stop_check: Optional[Callable[[], bool]] = None) -> Tuple[str, str, int]:
    """Выполняет скрипт как модуль (для .exe и разработки).
    Пробрасывает stop_check через builtins для доступности внутри скрипта."""
    original_argv = sys.argv
    original_stdout = sys.stdout
    original_stderr = sys.stderr
    stdout_capture = io.StringIO()
    stderr_capture = io.StringIO()
    module_name = script_name.replace('.py', '')

    try:
        sys.stdout = stdout_capture
        sys.stderr = stderr_capture

        # Пробрасываем stop_check в модуль через builtins
        if stop_check is not None:
            builtins.__dict__['_gui_stop_check'] = stop_check

        args = [script_name, "--from-gui", "--output-dir", str(output_folder)]
        if merge_mode:
            args.append("--merge")
        args.append(str(input_file))
        sys.argv = args

        if getattr(sys, 'frozen', False):
            if module_name in sys.modules:
                del sys.modules[module_name]
            module = __import__(module_name)
            if hasattr(module, 'main'):
                module.main()
            else:
                runpy.run_module(module_name, run_name="__main__")
        else:
            script_path = get_script_path(script_name, current_folder)
            if script_path and script_path.exists():
                runpy.run_path(str(script_path), run_name="__main__")
            else:
                stderr_capture.write(f"\nСкрипт не найден: {script_name}\n")
                return stdout_capture.getvalue(), stderr_capture.getvalue(), 1

    except SystemExit as e:
        return stdout_capture.getvalue(), stderr_capture.getvalue() + f"\nSystemExit: {e}", e.code or 0
    except Exception as e:
        import traceback
        stderr_capture.write(f"\nОшибка: {e}\n{traceback.format_exc()}")
        return stdout_capture.getvalue(), stderr_capture.getvalue(), 1
    finally:
        # Гарантированно очищаем stop_check из builtins
        builtins.__dict__.pop('_gui_stop_check', None)
        sys.stdout = original_stdout
        sys.stderr = original_stderr
        sys.argv = original_argv

    return stdout_capture.getvalue(), stderr_capture.getvalue(), 0


def _run_script_subprocess(script_name: str, input_file: Path, output_folder: Path,
                           current_folder: Path, merge_mode: bool = False,
                           process_holder=None) -> Tuple[int, str, str]:
    """Запускает скрипт через subprocess (для режима разработки)."""
    script_path = get_script_path(script_name, current_folder)
    if not script_path or not script_path.exists():
        return 1, "", f"Скрипт не найден: {script_name}"

    cmd = [sys.executable, str(script_path), "--from-gui", "--output-dir", str(output_folder)]
    if merge_mode:
        cmd.append("--merge")
    cmd.append(str(input_file))

    env = os.environ.copy()
    env['PYTHONIOENCODING'] = 'utf-8'

    process = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        cwd=str(current_folder),
        env=env
    )

    if process_holder and hasattr(process_holder, 'set_current_process'):
        process_holder.set_current_process(process)

    try:
        stdout_bytes, stderr_bytes = process.communicate(timeout=3600)
    except subprocess.TimeoutExpired:
        process.kill()
        stdout_bytes, stderr_bytes = process.communicate()
        stderr_bytes += b"\n[TIMEOUT] Process exceed time limit (3600 s)\n"
    except Exception:
        process.kill()
        process.wait()
        raise
    finally:
        if process_holder and hasattr(process_holder, 'set_current_process'):
            process_holder.set_current_process(None)

    stdout, stderr = read_process_output(stdout_bytes, stderr_bytes)
    return process.returncode, stdout, stderr


def run_script(script_name: str, input_file: Path, output_folder: Path,
               current_folder: Path, use_frozen: bool = False,
               merge_mode: bool = False, process_holder=None,
               stop_check: Optional[Callable[[], bool]] = None) -> Tuple[bool, str, str, List[Path]]:
    """Запускает скрипт, возвращает (success, message, full_output, created_files)."""
    start_time = time.time()

    if use_frozen:
        stdout, stderr, retcode = _run_script_as_module(
            script_name, input_file, output_folder, current_folder, merge_mode,
            stop_check=stop_check
        )
    else:
        retcode, stdout, stderr = _run_script_subprocess(
            script_name, input_file, output_folder, current_folder, merge_mode, process_holder
        )

    full_output = stdout + "\n" + stderr

    if retcode != 0:
        return False, f"Ошибка выполнения (код {retcode})", full_output, []

    created = []
    stem = input_file.stem

    if script_name == "extract_relays.py":
        for f in output_folder.glob("*.xlsx"):
            if "Реле" in f.stem or "Общий_Отчет_Реле" in f.stem:
                if f.stat().st_mtime >= start_time - 0.5:
                    created.append(f)
    elif script_name == "extract_specification.py":
        # Для extract_specification.py ищем оба типа файлов: реле и автоматы
        relay_candidate = output_folder / f"{stem}_Реле.xlsx"
        breaker_candidate = output_folder / f"{stem}_Автоматические_выключатели_вторичных_цепей.xlsx"
        if relay_candidate.exists():
            created.append(relay_candidate)
        if breaker_candidate.exists():
            created.append(breaker_candidate)
    else:
        expected_names = {
            "extract_circuit_breaker.py": f"{stem}_Автоматические_выключатели.xlsx",
            "extract_current_transformer.py": f"{stem}_Трансформаторы тока.xlsx",
            "extract_contactors_with_relays.py": f"{stem}_Контакторы_и_реле.xlsx",
            "extract_circuit_breaker_from_baskets.py": f"{stem}_Автоматические_выключатели_в_корзинах.xlsx",
            "extract_secondary_circuit_breaker.py": f"{stem}_Автоматические_выключатели_вторичных_цепей.xlsx",
        }
        if script_name in expected_names:
            candidate = output_folder / expected_names[script_name]
            if candidate.exists():
                created.append(candidate)
        else:
            for f in output_folder.glob(f"{stem}*.xlsx"):
                created.append(f)

    return True, "Обработка завершена", full_output, created


def process_files_for_operation(
    script_name: str,
    files: List[Path],
    output_folder: Path,
    operation_name: str,
    merge_mode: bool,
    current_folder: Path,
    use_frozen: bool,
    progress_callback: Optional[Callable[[int, str], None]] = None,
    stop_check: Optional[Callable[[], bool]] = None,
    process_holder: Optional[object] = None,
) -> Tuple[bool, List[str], str, List[Path], Set[str]]:
    """Обрабатывает несколько файлов для одной операции."""
    temp_output = output_folder / "temp_объединение" if merge_mode else output_folder
    temp_output.mkdir(exist_ok=True)

    all_messages = []
    all_outputs = []
    all_created = []
    all_missing = set()

    for i, file in enumerate(files):
        if stop_check and stop_check():
            all_messages.append("⏹️ Остановка по запросу пользователя")
            break

        if progress_callback:
            progress_callback(int((i + 1) / len(files) * 100), f"{operation_name}... ({i + 1}/{len(files)})")

        try:
            success, msg, output, created = run_script(
                script_name, file, temp_output, current_folder, use_frozen,
                merge_mode=False, process_holder=process_holder,
                stop_check=stop_check
            )
            all_messages.append(f"📄 {file.name}: {msg}")
            all_outputs.append(f"--- Файл: {file.name} ---\n{output}\n")
            all_created.extend(created)

            if script_name in ("extract_circuit_breaker.py", "extract_circuit_breaker_from_baskets.py",
                               "extract_secondary_circuit_breaker.py"):
                all_missing.update(extract_missing_types(output))
        except Exception as e:
            error_msg = f"❌ Неожиданная ошибка при обработке {file.name}: {e}"
            all_messages.append(error_msg)
            all_outputs.append(f"--- Файл: {file.name} ---\n{error_msg}\n")
            continue

    # Объединение результатов
    if merge_mode:
        if len(all_created) > 1:
            # Группируем файлы по типу для extract_specification.py
            if script_name == "extract_specification.py":
                relay_files = [f for f in all_created if "Реле" in f.name]
                breaker_files = [f for f in all_created if "Автоматические_выключатели_вторичных_цепей" in f.name]
                
                merged_files = []
                
                # Объединяем файлы с реле
                if relay_files:
                    relay_suffix = "Реле (объединенные)"
                    relay_merged_file = _safe_path(output_folder / f"{relay_suffix}.xlsx")
                    merge_success, merge_msg = merge_excel_files(relay_files, relay_merged_file)
                    if merge_success:
                        all_messages.append(f"📦 {len(relay_files)} файлов с реле объединены в: {relay_merged_file.name}")
                        merged_files.append(relay_merged_file)
                    else:
                        all_messages.append(f"❌ Ошибка объединения реле: {merge_msg}")
                
                # Объединяем файлы с автоматами
                if breaker_files:
                    breaker_suffix = "Автоматические выключатели вторичных цепей (объединенные)"
                    breaker_merged_file = _safe_path(output_folder / f"{breaker_suffix}.xlsx")
                    merge_success, merge_msg = merge_excel_files(breaker_files, breaker_merged_file)
                    if merge_success:
                        all_messages.append(f"📦 {len(breaker_files)} файлов с автоматами объединены в: {breaker_merged_file.name}")
                        merged_files.append(breaker_merged_file)
                    else:
                        all_messages.append(f"❌ Ошибка объединения автоматов: {merge_msg}")
                
                # Удаляем временные файлы
                for f in all_created:
                    try:
                        f.unlink()
                    except Exception:
                        pass
                all_created = merged_files
            else:
                # Стандартная логика для других скриптов
                suffix_map = {
                    "extract_circuit_breaker.py": "Автоматические выключатели (объединенные)",
                    "extract_current_transformer.py": "Трансформаторы тока (объединенные)",
                    "extract_contactors_with_relays.py": "Контакторы и реле (объединенные)",
                    "extract_circuit_breaker_from_baskets.py": "Автоматические выключатели в корзинах (объединенные)",
                    "extract_secondary_circuit_breaker.py": "Автоматические выключатели вторичных цепей (объединенные)",
                }
                suffix = suffix_map.get(script_name, "Объединенные результаты")
                merged_file = _safe_path(output_folder / f"{suffix}.xlsx")

                merge_success, merge_msg = merge_excel_files(all_created, merged_file)
                if merge_success:
                    all_messages.append(f"📦 {len(all_created)} файлов объединены в: {merged_file.name}")
                    for f in all_created:
                        try:
                            f.unlink()
                        except Exception:
                            pass
                    all_created = [merged_file]
                else:
                    all_messages.append(f"❌ Ошибка объединения: {merge_msg}")
        elif len(all_created) == 1:
            single_file = all_created[0]
            if single_file.parent != output_folder:
                new_path = _safe_path(output_folder / single_file.name)
                shutil.move(str(single_file), str(new_path))
                all_created = [new_path]

    if merge_mode and temp_output.exists() and temp_output != output_folder:
        shutil.rmtree(temp_output, ignore_errors=True)

    return True, all_messages, "\n".join(all_outputs), all_created, all_missing


def merge_excel_files(file_paths: List[Path], output_path: Path, source_column: str = "Исходный файл") -> Tuple[bool, str]:
    """Объединяет несколько Excel файлов в один. Возвращает (success, message)."""
    dataframes = []
    errors = []
    for f in file_paths:
        try:
            df = pd.read_excel(f, sheet_name=0)
            df[source_column] = f.stem
            dataframes.append(df)
        except Exception as e:
            errors.append(f"Ошибка чтения {f.name}: {e}")

    if errors:
        return False, "; ".join(errors)

    if not dataframes:
        return False, "Нет данных для объединения"

    try:
        merged = pd.concat(dataframes, ignore_index=True)
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            merged.to_excel(writer, sheet_name='Объединенные данные', index=False)
        return True, ""
    except Exception as e:
        return False, f"Ошибка записи {output_path.name}: {e}"


def run_merge_to_template(output_folder: Path, current_folder: Path, use_frozen: bool) -> Tuple[bool, str]:
    script_name = config.SCRIPT_NAMES["merge_template"]
    template_path = current_folder / config.TEMPLATE_FILENAME
    if not template_path.exists():
        return False, f"Файл шаблона не найден: {config.TEMPLATE_FILENAME}"

    output_msg = ""

    if use_frozen:
        original_argv = sys.argv
        original_stdout = sys.stdout
        original_stderr = sys.stderr
        stdout_capture = io.StringIO()
        stderr_capture = io.StringIO()
        try:
            sys.stdout = stdout_capture
            sys.stderr = stderr_capture
            sys.argv = [script_name, str(output_folder)]
            runpy.run_module("merge_circuit_breaker_QF", run_name="__main__")
        except Exception as e:
            return False, f"Ошибка выполнения: {e}\n{stderr_capture.getvalue()}"
        finally:
            sys.stdout = original_stdout
            sys.stderr = original_stderr
            sys.argv = original_argv
        output_msg = stdout_capture.getvalue() + stderr_capture.getvalue()
    else:
        script_path = get_script_path(script_name, current_folder)
        if not script_path or not script_path.exists():
            return False, f"Скрипт не найден: {script_name}"
        try:
            result = subprocess.run(
                [sys.executable, str(script_path), str(output_folder)],
                capture_output=True,
                cwd=str(current_folder)
            )
            stdout_decoded = result.stdout.decode('cp1251', errors='replace')
            stderr_decoded = result.stderr.decode('cp1251', errors='replace')
            output_msg = stdout_decoded + stderr_decoded
            if result.returncode != 0:
                return False, f"Ошибка (код {result.returncode}):\n{output_msg}"
        except Exception as e:
            return False, f"Ошибка запуска subprocess: {e}"

    pattern = "Таблица_автоматов_общая_*.xlsx"
    candidates = list(output_folder.glob(pattern))
    if candidates:
        newest = max(candidates, key=lambda p: p.stat().st_ctime)
        return True, f"Файл создан: {newest.name}"
    else:
        all_xlsx = list(output_folder.glob("*.xlsx"))
        dir_content = "\n".join([p.name for p in all_xlsx]) if all_xlsx else "(нет .xlsx файлов)"
        return False, (f"Не удалось найти выходной файл по маске '{pattern}'.\n"
                       f"Содержимое папки {output_folder}:\n{dir_content}\n\n"
                       f"Вывод скрипта:\n{output_msg}")


def process_relays_from_pdf(
    files: List[Path],
    output_folder: Path,
    current_folder: Path,
    merge_mode: bool,
    use_frozen: bool,
    progress_callback: Optional[Callable[[int, str], None]] = None,
    stop_check: Optional[Callable[[], bool]] = None,
    process_holder: Optional[object] = None,
) -> List[Path]:
    """Обрабатывает PDF файлы для извлечения реле."""
    script_name = config.SCRIPT_NAMES["relays_pdf"]
    all_created = []
    temp_output = output_folder / "temp_реле" if merge_mode else output_folder
    temp_output.mkdir(exist_ok=True)

    for i, file in enumerate(files):
        if stop_check and stop_check():
            break

        if progress_callback:
            progress_callback(int((i + 1) / len(files) * 50), f"Реле из PDF... ({i + 1}/{len(files)})")

        try:
            success, msg, output, created = run_script(
                script_name, file, temp_output, current_folder, use_frozen,
                merge_mode=False, process_holder=process_holder,
                stop_check=stop_check
            )
            all_created.extend(created)
        except Exception as e:
            # Логируем ошибку, но продолжаем обработку остальных файлов
            if progress_callback:
                progress_callback(int((i + 1) / len(files) * 50), f"Ошибка: {file.name}")
            continue

    if merge_mode:
        if len(all_created) > 1:
            merged_file = _safe_path(output_folder / "Реле_объединенные.xlsx")
            merge_success, merge_msg = merge_excel_files(all_created, merged_file)
            if merge_success:
                for f in all_created:
                    try:
                        f.unlink()
                    except Exception:
                        pass
                all_created = [merged_file]
        elif len(all_created) == 1:
            single_file = all_created[0]
            if single_file.parent != output_folder:
                new_path = _safe_path(output_folder / single_file.name)
                shutil.move(str(single_file), str(new_path))
                all_created = [new_path]

        if temp_output.exists() and temp_output != output_folder:
            shutil.rmtree(temp_output, ignore_errors=True)

    return all_created