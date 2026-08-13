import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.ttk as ttk
from tkinterdnd2 import TkinterDnD, DND_FILES
from pathlib import Path
import os
import sys
import queue
from typing import List, Dict, Optional
import config
from utils import parse_dnd_data, check_ktc_file, copy_to_clipboard
import logging

# Настраиваем логирование для GUI
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


class MainWindow(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode(config.APPEARANCE_MODE)
        ctk.set_default_color_theme(config.COLOR_THEME)

        self.title(config.WINDOW_TITLE)
        self.geometry(config.WINDOW_GEOMETRY)
        self.minsize(*config.WINDOW_MIN_SIZE)

        if hasattr(sys, 'frozen'):
            self.current_folder = Path(sys.executable).parent
        else:
            self.current_folder = Path(os.path.dirname(os.path.abspath(__file__)))

        self.result_folder = self.current_folder / config.DEFAULT_RESULTS_FOLDER_NAME
        self.result_folder.mkdir(parents=True, exist_ok=True)

        self.loaded_xlsx: List[Path] = []
        self.loaded_pdf: List[Path] = []

        # Переменные операций
        self.var_breakers = ctk.BooleanVar(value=False)
        self.var_transformers = ctk.BooleanVar(value=False)
        self.var_relays = ctk.BooleanVar(value=False)
        self.var_basket_breakers = ctk.BooleanVar(value=False)
        self.var_specification = ctk.BooleanVar(value=False)
        self.var_merge_template = ctk.BooleanVar(value=False)
        self.save_mode_var = ctk.StringVar(value="separate")
        self.result_path_var = ctk.StringVar(value=str(self.result_folder))

        self.controller = None

        # === ПОТОКОБЕЗОПАСНОСТЬ: очередь для обновлений GUI ===
        self._ui_queue = queue.Queue()

        self._create_widgets()
        self._setup_drag_drop()
        self.update_file_list()
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

        # Запускаем polling очереди
        self._poll_ui_queue()

    def _poll_ui_queue(self):
        """Периодически проверяет очередь и выполняет обновления в главном потоке."""
        try:
            while True:
                func, args, kwargs = self._ui_queue.get_nowait()
                try:
                    func(*args, **kwargs)
                except Exception:
                    pass  # Виджет мог быть уничтожен
        except queue.Empty:
            pass
        # Планируем следующую проверку через 50мс
        try:
            self.after(50, self._poll_ui_queue)
        except Exception:
            pass  # Окно закрыто

    def _safe_call(self, func, *args, **kwargs):
        """Ставит вызов функции в очередь для выполнения в главном потоке."""
        self._ui_queue.put((func, args, kwargs))

    def set_controller(self, controller):
        self.controller = controller

    # ---------- Создание виджетов (без изменений) ----------
    def _create_widgets(self):
        self.main_frame = ctk.CTkScrollableFrame(self, orientation="vertical")
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.header_frame = ctk.CTkFrame(self.main_frame)
        self.header_frame.pack(fill="x", padx=10, pady=(10, 5))
        self.title_label = ctk.CTkLabel(
            self.header_frame, text="⚡ Извлечение данных для ТО",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        self.title_label.pack(side="left", padx=10)

        self.sidebar_frame = ctk.CTkFrame(self.main_frame, width=280)
        self.sidebar_frame.pack(side="left", fill="y", padx=(10, 5), pady=5)
        self.sidebar_frame.pack_propagate(False)

        ctk.CTkLabel(self.sidebar_frame, text="⚙️ Настройки", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=(10, 5))
        ctk.CTkLabel(self.sidebar_frame, text="Режим сохранения:").pack(anchor="w", padx=10, pady=(10, 0))
        ctk.CTkRadioButton(self.sidebar_frame, text="Отдельные файлы", variable=self.save_mode_var, value="separate").pack(anchor="w", padx=20, pady=2)
        ctk.CTkRadioButton(self.sidebar_frame, text="Общий файл (объединенный)", variable=self.save_mode_var, value="merged").pack(anchor="w", padx=20, pady=2)
        ctk.CTkLabel(self.sidebar_frame, text=" ").pack(pady=10)

        ctk.CTkLabel(self.sidebar_frame, text="📁 Папка результатов:", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(10, 5))
        self.result_path_entry = ctk.CTkEntry(self.sidebar_frame, textvariable=self.result_path_var, state="readonly")
        self.result_path_entry.pack(fill="x", padx=10, pady=(0, 5))
        ctk.CTkButton(self.sidebar_frame, text="📂 Выбрать папку", command=self._выбрать_папку).pack(fill="x", padx=10, pady=(0, 5))
        self.folder_info_label = ctk.CTkLabel(self.sidebar_frame, text=f"Текущая папка:\n{self.result_folder}", wraplength=260, font=ctk.CTkFont(size=11))
        self.folder_info_label.pack(padx=10, pady=5)

        self.content_frame = ctk.CTkFrame(self.main_frame)
        self.content_frame.pack(side="right", fill="both", expand=True, padx=(5, 10), pady=5)
        ctk.CTkLabel(self.content_frame, text="📊 Обработка Excel/PDF файлов", font=ctk.CTkFont(size=18, weight="bold")).pack(anchor="w", padx=10, pady=(10, 5))
        ctk.CTkLabel(self.content_frame, text="Загрузите файлы и выберите нужные операции").pack(anchor="w", padx=10, pady=(0, 10))

        self.upload_frame = ctk.CTkFrame(self.content_frame)
        self.upload_frame.pack(fill="x", padx=10, pady=5)
        self.upload_btn = ctk.CTkButton(self.upload_frame, text="📎 Загрузить файлы", command=self._загрузить_файлы, height=40, font=ctk.CTkFont(size=14))
        self.upload_btn.pack(side="left", padx=5, pady=5)
        self.clear_btn = ctk.CTkButton(self.upload_frame, text="🗑️ Очистить", command=self._очистить_файлы, fg_color="gray", height=40)
        self.clear_btn.pack(side="left", padx=5, pady=5)

        ctk.CTkLabel(self.content_frame, text="📂 Загруженные файлы (перетащите сюда или используйте кнопку выше):", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(10, 5))
        self.drag_drop_frame = tk.Frame(self.content_frame, bg="#2b2b2b", bd=2, relief="solid", highlightbackground="#4a90e2", highlightthickness=2, height=120)
        self.drag_drop_frame.pack(fill="x", padx=10, pady=5)
        self.drag_drop_frame.pack_propagate(False)
        self.files_textbox = tk.Text(self.drag_drop_frame, bg="#2b2b2b", fg="#ffffff", insertbackground="#ffffff",
                                     font=("Segoe UI", 10), wrap="word", state="normal", relief="flat",
                                     highlightthickness=0, bd=0, height=6)
        self.files_textbox.pack(fill="both", expand=True, padx=8, pady=8)

        ctk.CTkLabel(self.content_frame, text="Выберите операции для выполнения:", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(10, 5))
        self.operations_tabview = ctk.CTkTabview(self.content_frame, width=400, height=180)
        self.operations_tabview.pack(fill="x", padx=10, pady=5)
        tab_xlsx = self.operations_tabview.add("📊 .xlsx")
        tab_pdf = self.operations_tabview.add("📄 .pdf")

        ctk.CTkCheckBox(tab_xlsx, text="🔌 Силовые автоматы из таблицы фидеров", variable=self.var_breakers).grid(row=0, column=0, padx=10, pady=5, sticky="w")
        ctk.CTkCheckBox(tab_xlsx, text="⚡ Трансформаторы тока", variable=self.var_transformers).grid(row=0, column=1, padx=10, pady=5, sticky="w")
        ctk.CTkCheckBox(tab_xlsx, text="🔧 Реле из корзин и силовые контакторы", variable=self.var_relays).grid(row=1, column=0, padx=10, pady=5, sticky="w")
        ctk.CTkCheckBox(tab_xlsx, text="🔌 Вторичные автоматы в корзинах (требуется КРУС-КТС)", variable=self.var_basket_breakers, command=self._on_basket_breakers_toggle).grid(row=1, column=1, padx=10, pady=5, sticky="w")
        ctk.CTkCheckBox(tab_xlsx, text="📋 Объединить автоматы в шаблон", variable=self.var_merge_template).grid(row=2, column=0, padx=10, pady=5, sticky="w", columnspan=2)
        ctk.CTkLabel(tab_xlsx, text="ℹ️ Объединяет все файлы *Автоматические_выключатели*.xlsx в шаблон\n(требуется Шаблон для автоматов.xlsx в папке скрипта)",
                     font=ctk.CTkFont(size=10), text_color="#888888").grid(row=3, column=0, padx=10, pady=(0, 5), sticky="w", columnspan=2)

        ctk.CTkCheckBox(tab_pdf, text="📋 Извлечь данные из спецификации", variable=self.var_specification).grid(row=0, column=0, padx=10, pady=5, sticky="w")
        ctk.CTkLabel(tab_pdf, text="ℹ️ Извлекает реле (KL, K, KLB) и вторичные автоматы (SF/QFD/SFD)\nиз PDF файлов схем",
                     font=ctk.CTkFont(size=12), text_color="#888888").grid(row=1, column=0, padx=10, pady=5, sticky="w")

        self.buttons_frame = ctk.CTkFrame(self.content_frame)
        self.buttons_frame.pack(fill="x", padx=10, pady=10)
        self.run_btn = ctk.CTkButton(self.buttons_frame, text="▶️ Запустить обработку", command=self._запустить_обработку,
                                     height=50, font=ctk.CTkFont(size=16, weight="bold"))
        self.run_btn.pack(side="left", fill="both", expand=True, padx=(0, 5))
        self.stop_btn = ctk.CTkButton(self.buttons_frame, text="⏹️ Стоп", command=self._остановить_обработку,
                                      height=50, font=ctk.CTkFont(size=16, weight="bold"), fg_color="#c62828", state="disabled")
        self.stop_btn.pack(side="left", fill="both", expand=True, padx=(5, 0))

        self.progress_frame = ctk.CTkFrame(self.content_frame)
        self.progress_frame.pack(fill="x", padx=10, pady=5)
        self.progress_label = ctk.CTkLabel(self.progress_frame, text=" ")
        self.progress_label.pack()
        self.progress_bar = ttk.Progressbar(self.progress_frame, mode='determinate')
        self.progress_bar.pack(fill="x", padx=5, pady=5)

        ctk.CTkLabel(self.content_frame, text="📝 Результаты обработки:", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(10, 0))
        output_container = ctk.CTkFrame(self.content_frame)
        output_container.pack(fill="both", expand=True, padx=10, pady=5)
        self.output_textbox = ctk.CTkTextbox(output_container, height=200)
        self.output_textbox.pack(fill="both", expand=True, side="left")
        self.copy_btn = ctk.CTkButton(output_container, text="📋 Копировать", command=self._копировать_вывод,
                                      width=100, height=30, font=ctk.CTkFont(size=12))
        self.copy_btn.pack(side="right", padx=(5, 0), pady=(0, 10), anchor="ne")

    # ---------- Потокобезопасные методы обновления GUI ----------
    def add_output(self, text: str, type: str = "info"):
        """Потокобезопасная версия — ставит задачу в очередь."""
        self._safe_call(self._add_output_impl, text, type)

    def _add_output_impl(self, text: str, type: str):
        if not self.winfo_exists():
            return
        try:
            self.output_textbox.insert("end", f"{text}\n")
            self.output_textbox.see("end")
        except Exception:
            pass

    def clear_output(self):
        self._safe_call(self._clear_output_impl)

    def _clear_output_impl(self):
        if not self.winfo_exists():
            return
        try:
            self.output_textbox.delete("1.0", "end")
        except Exception:
            pass

    def update_progress(self, value: int, label: str):
        self._safe_call(self._update_progress_impl, value, label)

    def _update_progress_impl(self, value: int, label: str):
        if not self.winfo_exists():
            return
        try:
            self.progress_bar['value'] = value
            self.progress_label.configure(text=label)
        except Exception:
            pass

    def update_file_list(self):
        self._safe_call(self._update_file_list_impl)

    def _update_file_list_impl(self):
        if not self.winfo_exists():
            return
        try:
            self.files_textbox.config(state="normal")
            self.files_textbox.delete("1.0", "end")
            total = len(self.loaded_xlsx) + len(self.loaded_pdf)
            if total > 0:
                self.files_textbox.insert("1.0", f"📁 Выбрано файлов: {total}\n\n", "header")
                if self.loaded_xlsx:
                    self.files_textbox.insert("end", "📊 .xlsx файлы:\n", "category")
                    for i, f in enumerate(self.loaded_xlsx, 1):
                        self.files_textbox.insert("end", f"  {i}. {f.name}\n", "file")
                    self.files_textbox.insert("end", "\n")
                if self.loaded_pdf:
                    self.files_textbox.insert("end", "📄 .pdf файлы:\n", "category")
                    for i, f in enumerate(self.loaded_pdf, 1):
                        self.files_textbox.insert("end", f"  {i}. {f.name}\n", "file")
            else:
                self.files_textbox.insert("1.0", "📂 Перетащите файлы Excel или PDF сюда\nили нажмите кнопку 'Загрузить файлы' выше", "placeholder")
            self.files_textbox.tag_config("header", foreground="#4a90e2", font=("Segoe UI", 12, "bold"))
            self.files_textbox.tag_config("category", foreground="#ffa500", font=("Segoe UI", 11, "bold"))
            self.files_textbox.tag_config("file", foreground="#ffffff", font=("Segoe UI", 11))
            self.files_textbox.tag_config("placeholder", foreground="#888888", font=("Segoe UI", 11, "italic"))
            self.files_textbox.config(state="disabled")
        except Exception:
            pass

    def set_buttons_state(self, enabled: bool):
        self._safe_call(self._set_buttons_state_impl, enabled)

    def _set_buttons_state_impl(self, enabled: bool):
        if not self.winfo_exists():
            return
        state = "normal" if enabled else "disabled"
        try:
            self.upload_btn.configure(state=state)
            self.clear_btn.configure(state=state)
            self.run_btn.configure(state=state)
        except Exception:
            pass

    def set_run_button_state(self, enabled: bool, text: str = ""):
        self._safe_call(self._set_run_button_state_impl, enabled, text)

    def _set_run_button_state_impl(self, enabled: bool, text: str):
        if not self.winfo_exists():
            return
        try:
            self.run_btn.configure(state="normal" if enabled else "disabled")
            if text:
                self.run_btn.configure(text=text)
        except Exception:
            pass

    def set_stop_button_state(self, enabled: bool):
        self._safe_call(self._set_stop_button_state_impl, enabled)

    def _set_stop_button_state_impl(self, enabled: bool):
        if not self.winfo_exists():
            return
        try:
            self.stop_btn.configure(state="normal" if enabled else "disabled")
        except Exception:
            pass

    def show_message(self, title: str, message: str, type: str = "warning"):
        self._safe_call(self._show_message_impl, title, message, type)

    def _show_message_impl(self, title: str, message: str, type: str):
        if not self.winfo_exists():
            return
        try:
            if type == "warning":
                messagebox.showwarning(title, message)
            else:
                messagebox.showinfo(title, message)
        except Exception:
            pass

    # ---------- Получение состояния (вызываются из главного потока) ----------
    def get_selected_operations(self) -> Dict[str, bool]:
        return {
            "breakers": self.var_breakers.get(),
            "transformers": self.var_transformers.get(),
            "relays": self.var_relays.get(),
            "basket_breakers": self.var_basket_breakers.get(),
            "specification": self.var_specification.get(),
            "merge_template": self.var_merge_template.get(),
        }

    def get_files(self) -> Dict[str, List[Path]]:
        return {"xlsx": self.loaded_xlsx, "pdf": self.loaded_pdf}

    def get_save_mode(self) -> str:
        return self.save_mode_var.get()

    def get_result_folder(self) -> Path:
        return Path(self.result_path_var.get())

    # ---------- События GUI ----------
    def _on_basket_breakers_toggle(self):
        if self.var_basket_breakers.get():
            found, file = check_ktc_file(self.current_folder)
            if not found:
                messagebox.showwarning(
                    "Отсутствует файл КРУС-КТС!",
                    f"Для работы функции 'Автоматы в корзинах' требуется файл КРУС-КТС (PDF).\n\n"
                    f"Поместите PDF файл с буквами 'KTC' в названии в папку:\n{self.current_folder}\n\n"
                    "Операция будет отключена."
                )
                self.var_basket_breakers.set(False)
            else:
                self.add_output(f"✅ Файл КРУС-КТС найден: {file.name}", "info")

    def _запустить_обработку(self):
        if self.controller:
            self.controller.start_processing()

    def _остановить_обработку(self):
        if self.controller:
            self.controller.stop_processing()

    def _загрузить_файлы(self):
        files = filedialog.askopenfilenames(
            title="Выберите файлы",
            filetypes=[("Все поддерживаемые", "*.xlsx *.pdf"), ("Excel файлы", "*.xlsx"), ("PDF файлы", "*.pdf")]
        )
        for f in files:
            p = Path(f)
            if p.suffix.lower() == config.ALLOWED_XLSX_EXT:
                if p not in self.loaded_xlsx:
                    self.loaded_xlsx.append(p)
            elif p.suffix.lower() == config.ALLOWED_PDF_EXT:
                if p not in self.loaded_pdf:
                    self.loaded_pdf.append(p)
        self.update_file_list()

    def _очистить_файлы(self):
        """Очищает список загруженных файлов и кэш распознанных PDF."""
        logger.info("Пользователь нажал кнопку очистки файлов")
        self.loaded_xlsx = []
        self.loaded_pdf = []
        # Очищаем кэш при очистке файлов
        from cache_manager import init_cache_manager, clear_global_cache
        # Сначала инициализируем менеджер кэша, чтобы убедиться что он существует
        init_cache_manager()
        clear_global_cache()
        logger.info("Кэш распознанных PDF файлов очищен")
        self.add_output("🗑️ Кэш распознанных PDF файлов очищен", "info")
        self.update_file_list()

    def _выбрать_папку(self):
        folder = filedialog.askdirectory(title="Выберите папку для сохранения результатов", initialdir=self.result_path_var.get())
        if folder:
            self.result_folder = Path(folder)
            self.result_path_var.set(str(folder))
            self.folder_info_label.configure(text=f"Текущая папка:\n{folder}")

    def _копировать_вывод(self):
        text = self.output_textbox.get("1.0", "end-1c")
        if text.strip():
            copy_to_clipboard(self, text)
            self.add_output("📋 Текст скопирован в буфер обмена", "info")

    def _on_closing(self):
        """Обработчик закрытия окна. Очищает кэш и закрывает окно."""
        logger.info("Пользователь закрывает программу")
        # Очищаем кэш при закрытии программы
        from cache_manager import init_cache_manager, clear_global_cache
        # Сначала инициализируем менеджер кэша, чтобы убедиться что он существует
        init_cache_manager()
        clear_global_cache()
        logger.info("Кэш распознанных PDF файлов очищен при закрытии программы")
        
        if self.controller:
            if self.controller.on_closing():
                self.destroy()
        else:
            self.destroy()

    # ---------- Drag & Drop ----------
    def _setup_drag_drop(self):
        self.drag_drop_frame.drop_target_register(DND_FILES)
        self.drag_drop_frame.dnd_bind('<<Drop>>', self._on_drop)
        self.files_textbox.drop_target_register(DND_FILES)
        self.files_textbox.dnd_bind('<<Drop>>', self._on_drop)
        self.drag_drop_frame.bind("<Enter>", self._on_drag_enter)
        self.drag_drop_frame.bind("<Leave>", self._on_drag_leave)

    def _on_drop(self, event):
        raw = event.data
        parsed = parse_dnd_data(raw)
        added_xlsx = 0
        added_pdf = 0
        for typ, p in parsed:
            if typ == 'xlsx':
                if p not in self.loaded_xlsx:
                    self.loaded_xlsx.append(p)
                    added_xlsx += 1
            elif typ == 'pdf':
                if p not in self.loaded_pdf:
                    self.loaded_pdf.append(p)
                    added_pdf += 1
        if added_xlsx or added_pdf:
            self.update_file_list()
            self.add_output(f"✅ Добавлено файлов: {added_xlsx} .xlsx, {added_pdf} .pdf", "info")
        else:
            self.add_output("⚠️ Не добавлено ни одного файла. Поддерживаются .xlsx и .pdf файлы", "warning")
        self._on_drag_leave(None)

    def _on_drag_enter(self, event):
        try:
            self.drag_drop_frame.configure(bg="#1e3a1a", highlightbackground="#00ff00", highlightthickness=2)
        except Exception:
            pass

    def _on_drag_leave(self, event):
        try:
            self.drag_drop_frame.configure(bg="#2b2b2b", highlightbackground="#4a90e2", highlightthickness=2)
        except Exception:
            pass