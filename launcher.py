# launcher.py (исправленная версия с поддержкой pdfplumber и корректной проверкой пакетов)
import sys
import os
import subprocess
import importlib
import site
from pathlib import Path

# Отключаем цветной вывод для Windows, чтобы избежать проблем с кодировкой
USE_COLORS = False  # Принудительно отключаем цвета
try:
    import platform
    if platform.system() == 'Windows':
        USE_COLORS = False
except:
    pass

def safe_print(text, end='\n'):
    """Безопасная печать без эмодзи и спецсимволов"""
    try:
        # Удаляем эмодзи и заменяем на безопасные символы
        text = text.replace('🔍', '[ПРОВЕРКА]')
        text = text.replace('✅', '[OK]')
        text = text.replace('❌', '[ОШИБКА]')
        text = text.replace('⚠️', '[ПРЕДУПРЕЖДЕНИЕ]')
        text = text.replace('📦', '[УСТАНОВКА]')
        text = text.replace('🚀', '[ЗАПУСК]')
        text = text.replace('⏳', '[ОЖИДАНИЕ]')
        text = text.replace('📁', '[ФАЙЛ]')
        text = text.replace('⚡', '[ЭЛЕКТРИКА]')
        text = text.replace('📊', '[ДАННЫЕ]')
        text = text.replace('📄', '[ДОКУМЕНТ]')
        text = text.replace('🔧', '[НАСТРОЙКА]')
        text = text.replace('🗑️', '[УДАЛЕНИЕ]')
        text = text.replace('💾', '[СОХРАНЕНИЕ]')
        
        # Пытаемся закодировать в cp1251 для Windows
        try:
            encoded = text.encode('cp1251', errors='replace').decode('cp1251')
            print(encoded, end=end)
        except:
            print(text, end=end)
    except:
        print(text, end=end)

def is_gui_mode():
    """Определяет, запущено ли приложение в GUI-режиме"""
    try:
        if sys.stdin is None:
            return True
        return not sys.stdin.isatty()
    except:
        return True

def show_message_box(title, message, type='info'):
    """Показывает сообщение в GUI-окне"""
    try:
        import tkinter as tk
        from tkinter import messagebox
        
        root = tk.Tk()
        root.withdraw()
        
        if type == 'error':
            messagebox.showerror(title, message)
        elif type == 'warning':
            messagebox.showwarning(title, message)
        else:
            messagebox.showinfo(title, message)
        
        root.destroy()
    except:
        pass

def get_pdfplumber_dependencies():
    """
    Возвращает список всех зависимостей pdfplumber
    """
    return [
        'pdfplumber',
        'pdfminer.six',
        'Pillow',
        'cffi',
        'cryptography',
        'pypdfium2'
    ]

def check_dependency(package_name):
    """
    Проверяет, установлен ли пакет.
    Для .exe используем другой подход - пробуем импортировать.
    """
    try:
        # Для customtkinter и tkinterdnd2 специальная обработка
        if package_name == 'customtkinter':
            import customtkinter
            return True
        elif package_name == 'tkinterdnd2':
            import tkinterdnd2
            return True
        elif package_name == 'pdfplumber':
            import pdfplumber
            return True
        else:
            # Для остальных пакетов
            import_name = package_name.replace('-', '_')
            importlib.import_module(import_name)
            return True
    except ImportError as e:
        print(f"Import error for {package_name}: {e}")
        return False

def check_all_dependencies():
    """
    Проверяет все зависимости
    """
    required = [
        "customtkinter",
        "tkinterdnd2",
        "pandas",
        "openpyxl",
        "pdfplumber"
    ]
    
    missing = []
    
    safe_print("[ПРОВЕРКА] Проверка зависимостей...")
    safe_print("=" * 50)
    
    for package in required:
        if check_dependency(package):
            safe_print(f"[OK] {package} - установлен")
        else:
            safe_print(f"[ОШИБКА] {package} - ОТСУТСТВУЕТ")
            missing.append(package)
    
    return missing

def install_package_with_retry(package):
    """
    Устанавливает пакет с повторными попытками
    """
    max_retries = 3
    for attempt in range(max_retries):
        try:
            safe_print(f"[ОЖИДАНИЕ] Установка {package} (попытка {attempt + 1}/{max_retries})...")
            
            # Используем разные источники для установки
            if attempt == 0:
                # Основная попытка
                result = subprocess.run(
                    [sys.executable, "-m", "pip", "install", package],
                    capture_output=True,
                    text=True,
                    timeout=120
                )
            elif attempt == 1:
                # Попытка с флагом --no-cache-dir
                result = subprocess.run(
                    [sys.executable, "-m", "pip", "install", "--no-cache-dir", package],
                    capture_output=True,
                    text=True,
                    timeout=120
                )
            else:
                # Попытка с установкой из конкретного индекса
                result = subprocess.run(
                    [sys.executable, "-m", "pip", "install", "--index-url", "https://pypi.org/simple", package],
                    capture_output=True,
                    text=True,
                    timeout=120
                )
            
            if result.returncode == 0:
                safe_print(f"[OK] {package} установлен")
                return True
            else:
                safe_print(f"[ПРЕДУПРЕЖДЕНИЕ] Ошибка при установке {package}: {result.stderr[:200]}")
                
        except subprocess.TimeoutExpired:
            safe_print(f"[ПРЕДУПРЕЖДЕНИЕ] Таймаут при установке {package}")
        except Exception as e:
            safe_print(f"[ПРЕДУПРЕЖДЕНИЕ] Ошибка: {e}")
    
    return False

def install_packages(packages):
    """Устанавливает пакеты через pip"""
    safe_print("[УСТАНОВКА] Установка зависимостей...")
    safe_print("=" * 50)
    
    try:
        # Сначала обновляем pip
        safe_print("[ОЖИДАНИЕ] Обновление pip...")
        try:
            subprocess.run(
                [sys.executable, "-m", "pip", "install", "--upgrade", "pip"],
                capture_output=True,
                text=True,
                timeout=60
            )
        except:
            pass
        
        # Устанавливаем пакеты
        success = True
        
        for package in packages:
            if not install_package_with_retry(package):
                success = False
        
        # Дополнительная проверка: если pdfplumber не установился, пробуем установить его зависимости
        if 'pdfplumber' in packages and not check_dependency('pdfplumber'):
            safe_print("[ОЖИДАНИЕ] Установка зависимостей pdfplumber...")
            
            deps = get_pdfplumber_dependencies()
            for dep in deps:
                if dep != 'pdfplumber':
                    install_package_with_retry(dep)
            
            # Пробуем установить pdfplumber еще раз
            install_package_with_retry('pdfplumber')
        
        # Проверяем результат
        if check_dependency('pdfplumber'):
            safe_print("[OK] pdfplumber успешно установлен!")
            return True
        else:
            safe_print("[ОШИБКА] Не удалось установить pdfplumber")
            return False
            
    except Exception as e:
        safe_print(f"[ОШИБКА] Ошибка при установке: {e}")
        return False

def show_install_dialog(missing_packages):
    """Показывает GUI-диалог для подтверждения установки"""
    try:
        import tkinter as tk
        from tkinter import messagebox
        
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)  # Делаем окно поверх всех
        
        package_list = "\n".join([f"• {pkg}" for pkg in missing_packages])
        message = (
            f"Отсутствуют необходимые пакеты:\n\n{package_list}\n\n"
            "Установить их автоматически?"
        )
        
        response = messagebox.askyesno(
            "Установка зависимостей",
            message,
            icon='question'
        )
        
        root.destroy()
        return response
        
    except Exception as e:
        print(f"Ошибка при показе диалога: {e}")
        # Если не удалось показать GUI диалог, используем консоль
        try:
            response = input("Установить отсутствующие пакеты? (y/n): ").strip().lower()
            return response in ['y', 'yes']
        except:
            return False

def main():
    """Запускает приложение после проверки зависимостей"""
    
    try:
        # Проверяем зависимости
        missing = check_all_dependencies()
        
        if missing:
            safe_print(f"\n[ПРЕДУПРЕЖДЕНИЕ] Отсутствуют {len(missing)} пакетов:")
            for pkg in missing:
                safe_print(f"   - {pkg}")
            
            # Определяем режим
            if is_gui_mode():
                response = show_install_dialog(missing)
                if not response:
                    safe_print("[ОШИБКА] Установка отменена пользователем")
                    sys.exit(1)
            else:
                try:
                    response = input("\nУстановить отсутствующие пакеты? (y/n): ").strip().lower()
                    if response not in ['y', 'yes']:
                        safe_print("[ОШИБКА] Установка отменена")
                        sys.exit(1)
                except:
                    safe_print("[ОШИБКА] Невозможно получить ввод")
                    sys.exit(1)
            
            # Устанавливаем пакеты
            if not install_packages(missing):
                safe_print("[ОШИБКА] Не удалось установить зависимости")
                if is_gui_mode():
                    show_message_box(
                        "Ошибка установки",
                        "Не удалось установить необходимые пакеты.\n"
                        "Попробуйте установить их вручную:\n\n"
                        "pip install customtkinter tkinterdnd2 pandas openpyxl pdfplumber",
                        'error'
                    )
                sys.exit(1)
            
            safe_print("[OK] Все зависимости установлены!")
        
        # Запускаем основное приложение
        safe_print("\n[ЗАПУСК] Запуск приложения...")
        safe_print("=" * 50)
        
        try:
            # Проверяем, что customtkinter доступен
            import customtkinter as ctk
            import tkinterdnd2
            
            # Импортируем и запускаем app
            from app import App
            App().mainloop()
            
        except ImportError as e:
            error_msg = f"[ОШИБКА] Ошибка при запуске app.py: {e}"
            safe_print(error_msg)
            
            if is_gui_mode():
                show_message_box("Ошибка запуска", str(e), 'error')
            
            sys.exit(1)
            
    except Exception as e:
        error_msg = f"[ОШИБКА] Критическая ошибка: {e}"
        safe_print(error_msg)
        
        if is_gui_mode():
            show_message_box("Критическая ошибка", str(e), 'error')
        
        sys.exit(1)

if __name__ == "__main__":
    main()