import threading
import sys
from pathlib import Path
from typing import Set
import tkinter.messagebox as messagebox

import config
from engine import process_files_for_operation, run_merge_to_template, process_relays_from_pdf
from utils import check_ktc_file


class Controller:
    def __init__(self, gui, current_folder: Path):
        self.gui = gui
        self.current_folder = current_folder
        self.processing_active = False
        self.stop_requested = False
        self.current_process = None
        self.processing_thread = None

    def start_processing(self):
        """Запускает обработку в отдельном потоке."""
        # Чтение состояния GUI
        selected = self.gui.get_selected_operations()
        files = self.gui.get_files()
        save_mode = self.gui.get_save_mode()
        output_folder = self.gui.get_result_folder()
        use_frozen = getattr(sys, 'frozen', False)
        merge_mode = (save_mode == "merged")

        # Проверки
        has_xlsx_ops = any(selected[op] for op in ["breakers", "transformers", "relays", "basket_breakers"])
        has_pdf_ops = selected.get("specification", False)

        if has_xlsx_ops and not files["xlsx"]:
            self.gui.show_message("Предупреждение", "Для выбранных операций с .xlsx требуются Excel файлы!")
            return
        if has_pdf_ops and not files["pdf"]:
            self.gui.show_message("Предупреждение", "Для выбранных PDF операций требуются PDF файлы!")
            return
        if selected.get("merge_template", False):
            template = self.current_folder / config.TEMPLATE_FILENAME
            if not template.exists():
                self.gui.show_message(
                    "Отсутствует файл шаблона!",
                    f"Для операции 'Объединить автоматы в шаблон' требуется файл:\n{config.TEMPLATE_FILENAME}\n\n"
                    f"Поместите его в папку:\n{self.current_folder}"
                )
                return
            
        if selected.get("basket_breakers", False):
            found, _ = check_ktc_file(self.current_folder)
            if not found:
                self.gui.show_message(
                    "Отсутствует файл КРУС-КТС!",
                    f"Для работы функции 'Автоматы в корзинах' требуется файл КРУС-КТС (PDF).\n\n"
                    f"Поместите PDF файл с буквами 'KTC' в названии в папку:\n{self.current_folder}"
                )
                return
        
        # Проверка наличия файла библиотеки номиналов для операций с автоматами
        library_ops = ["breakers", "basket_breakers", "specification"]
        if any(selected[op] for op in library_ops):
            library_path = self.current_folder / config.LIBRARY_FILENAME
            if not library_path.exists():
                self.gui.show_message(
                    "Отсутствует файл библиотеки номиналов!",
                    f"Для выбранных операций с автоматами требуется файл библиотеки:\n{config.LIBRARY_FILENAME}\n\n"
                    f"Поместите его в папку:\n{self.current_folder}"
                )
                return

        # Сброс
        self.stop_requested = False
        self.processing_active = True
        self.gui.set_buttons_state(False)          # Блокируем кнопки загрузки/запуска
        self.gui.set_stop_button_state(True)       # Активируем кнопку Стоп
        self.gui.clear_output()
        self.gui.add_output("🚀 Начинаем обработку...", "info")

        self.processing_thread = threading.Thread(target=self._processing_thread)
        self.processing_thread.daemon = True
        self.processing_thread.start()

    def set_current_process(self, proc):
        """Сохраняет текущий процесс для возможности остановки."""
        self.current_process = proc

    def stop_processing(self):
        """Останавливает обработку."""
        if messagebox.askyesno("Подтверждение", "Вы действительно хотите остановить обработку?"):
            self.stop_requested = True
            self.gui.add_output("🛑 Запрошена остановка обработки...", "warning")
            if self.current_process and self.current_process.poll() is None:
                try:
                    self.current_process.kill()
                    self.gui.add_output("💀 Процесс принудительно завершен.", "warning")
                except:
                    pass

    def _processing_thread(self):
        """Основной поток обработки."""
        try:
            selected = self.gui.get_selected_operations()
            files = self.gui.get_files()
            save_mode = self.gui.get_save_mode()
            output_folder = self.gui.get_result_folder()
            use_frozen = getattr(sys, 'frozen', False)
            merge_mode = (save_mode == "merged")

            all_created = []
            all_missing = set()

            # Обработка операций
            ops = [
                ("breakers", "extract_circuit_breaker.py", "Обработка выключателей", files["xlsx"]),
                ("transformers", "extract_current_transformer.py", "Обработка трансформаторов", files["xlsx"]),
                ("relays", "extract_contactors_with_relays.py", "Обработка реле и контакторов", files["xlsx"]),
                ("basket_breakers", "extract_circuit_breaker_from_baskets.py", "Обработка автоматов в корзинах", files["xlsx"]),
            ]
            for op_key, script, label, file_list in ops:
                if selected.get(op_key, False) and not self.stop_requested:
                    self.gui.add_output(f"\n🔌 {label}...", "info")
                    success, msgs, _, created, missing = process_files_for_operation(
                        script, file_list, output_folder, label, merge_mode,
                        self.current_folder, use_frozen,
                        progress_callback=self._update_progress,
                        stop_check=lambda: self.stop_requested,
                        process_holder=self   # <-- передаём себя как process_holder
                    )
                    for msg in msgs:
                        self.gui.add_output(f"  {msg}")
                    all_created.extend(created)
                    all_missing.update(missing)

            # Извлечение данных из спецификации (реле и автоматы из PDF)
            if selected.get("specification", False) and not self.stop_requested:
                self.gui.add_output("\n📋 Извлечение данных из спецификации...", "info")
                success, msgs, _, created, missing = process_files_for_operation(
                    "extract_specification.py", files["pdf"], output_folder, "Извлечение данных из спецификации", merge_mode,
                    self.current_folder, use_frozen,
                    progress_callback=self._update_progress,
                    stop_check=lambda: self.stop_requested,
                    process_holder=self
                )
                for msg in msgs:
                    self.gui.add_output(f"  {msg}")
                all_created.extend(created)
                all_missing.update(missing)

            # Объединение в шаблон
            if selected.get("merge_template", False) and not self.stop_requested:
                self.gui.add_output("\n📋 Запуск объединения в шаблон...", "info")
                success, msg = run_merge_to_template(output_folder, self.current_folder, use_frozen)
                if success:
                    self.gui.add_output(f"✅ Объединение в шаблон завершено!\n{msg}", "success")
                else:
                    self.gui.add_output(f"❌ Объединение в шаблон завершилось с ошибкой:\n{msg}", "error")

            # Итог
            self.gui.update_progress(100, "Обработка завершена!" if not self.stop_requested else "Обработка остановлена!")
            if self.stop_requested:
                self.gui.add_output("\n" + "=" * 60, "warning")
                self.gui.add_output("⏹️ Обработка была остановлена пользователем.", "warning")
            else:
                self.gui.add_output("\n" + "=" * 60, "info")
                self.gui.add_output("✅ Обработка завершена!", "success")
                if all_created:
                    self.gui.add_output(f"\n📁 Созданные файлы ({len(all_created)}):", "info")
                    for f in all_created:
                        self.gui.add_output(f"  - {f.name}")
                if all_missing:
                    self._show_missing_report(all_missing)

        except Exception as e:
            self.gui.add_output(f"❌ Критическая ошибка: {e}", "error")
            import traceback
            self.gui.add_output(traceback.format_exc(), "error")
        finally:
            self.processing_active = False
            self.current_process = None   # очищаем процесс
            self.gui.set_buttons_state(True)
            self.gui.set_run_button_state(True, "▶️ Запустить обработку")
            self.gui.set_stop_button_state(False)

    def _update_progress(self, value: int, label: str):
        self.gui.update_progress(value, label)

    def _show_missing_report(self, missing: Set[str]):
        self.gui.add_output(" ", "info")
        self.gui.add_output("=" * 60, "info")
        self.gui.add_output("📋 Отсутствующие типы автоматических выключателей:", "info")
        self.gui.add_output(f"📊 Всего уникальных типов: {len(missing)}", "info")
        self.gui.add_output("=" * 60, "info")
        for t in sorted(missing):
            self.gui.add_output(f"  - {t}", "warning")
        self.gui.add_output(" ", "info")
        self.gui.add_output("💡 Скопируйте эти типы для добавления в библиотеку Library_nominals.xlsx", "info")

    def on_closing(self) -> bool:
        """Обработчик закрытия окна. Возвращает True, если можно закрыть."""
        if self.processing_active:
            return messagebox.askyesno("Подтверждение", "Обработка еще выполняется. Вы уверены, что хотите выйти?")
        return True