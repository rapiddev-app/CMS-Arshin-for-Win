"""
GUI для ScriptManager на базе CustomTkinter
Итерация 8: Полная интеграция с новыми скриптами
"""

import logging
import sys
import os
from typing import Dict, Any
from datetime import datetime
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox

# Добавляем путь к src для импортов
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    import customtkinter as ctk
    from config_manager import ConfigManager
    from scheduler import TaskSchedulerManager
except ImportError as e:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Ошибка запуска", f"Не удалось загрузить компоненты:\n{e}\n\nВозможно, повреждена установка.")
    sys.exit(1)
except Exception as e:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Критическая ошибка", f"Ошибка при запуске:\n{e}")
    sys.exit(1)


class PasswordDialog(ctk.CTkToplevel):
    """Диалог ввода пароля для доступа к настройкам"""

    ADMIN_PASSWORD = "cms2025"  # Пароль администратора

    def __init__(self, parent):
        super().__init__(parent)

        self.title("Введите пароль")
        self.geometry("350x150")
        self.transient(parent)
        self.grab_set()

        self.result = False

        # Центрируем окно
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (350 // 2)
        y = (self.winfo_screenheight() // 2) - (150 // 2)
        self.geometry(f"+{x}+{y}")

        # Label
        ctk.CTkLabel(
            self,
            text="Для изменения настроек введите пароль:",
            font=("Arial", 12)
        ).pack(pady=20)

        # Entry
        self.password_entry = ctk.CTkEntry(
            self,
            width=250,
            show="*",
            placeholder_text="Пароль"
        )
        self.password_entry.pack(pady=10)
        self.password_entry.focus()
        self.password_entry.bind("<Return>", lambda e: self.check_password())

        # Кнопки
        buttons_frame = ctk.CTkFrame(self, fg_color="transparent")
        buttons_frame.pack(pady=10)

        ctk.CTkButton(
            buttons_frame,
            text="OK",
            command=self.check_password,
            width=100
        ).pack(side="left", padx=5)

        ctk.CTkButton(
            buttons_frame,
            text="Отмена",
            command=self.destroy,
            width=100
        ).pack(side="left", padx=5)

    def check_password(self):
        """Проверка введённого пароля"""
        if self.password_entry.get() == self.ADMIN_PASSWORD:
            self.result = True
            self.destroy()
        else:
            messagebox.showerror("Ошибка", "Неверный пароль!")
            self.password_entry.delete(0, "end")
            self.password_entry.focus()


class ScriptManagerGUI:
    """
    Главное окно приложения ScriptManager

    Отображает список из 6 скриптов с возможностью:
    - Включить/Отключить скрипт
    - Настроить параметры скрипта (с паролем)
    - Создать задачу в Task Scheduler
    - Тестовая отправка
    - Просмотреть статус (следующий запуск, последний результат)
    """

    def __init__(self):
        """Инициализация GUI"""
        # Настройка темы
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # Создаём главное окно
        self.root = ctk.CTk()
        self.root.title("ScriptManager v1.0.0")
        self.root.geometry("1000x750")

        # Установка иконки окна
        icon_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets", "ScriptM.ico")
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
            except Exception as e:
                logging.warning(f"Не удалось установить иконку: {e}")

        # Менеджеры
        self.config_manager = ConfigManager()
        self.scheduler_manager = TaskSchedulerManager()

        # Загружаем конфигурацию скриптов
        self.scripts_config = self.config_manager.load_scripts_config()

        # Словарь для хранения виджетов скриптов
        self.script_widgets: Dict[str, Dict[str, Any]] = {}

        # Создаём интерфейс
        self.create_main_layout()

        logging.info("GUI инициализирован")

    def create_main_layout(self):
        """Создание главного интерфейса"""
        # Заголовок
        header = ctk.CTkLabel(
            self.root,
            text="ScriptManager - Управление скриптами",
            font=("Arial", 20, "bold")
        )
        header.pack(pady=(20, 10))

        # Фрейм для кастомных вкладок (выравнивание по левому краю)
        self.tabs_container = ctk.CTkFrame(self.root, fg_color="transparent")
        self.tabs_container.pack(fill="x", padx=20)

        # Контейнер для содержимого вкладок
        self.content_frame = ctk.CTkFrame(self.root, corner_radius=0)
        self.content_frame.pack(pady=(0, 10), padx=20, fill="both", expand=True)

        # Фреймы
        self.frames = {
            "Скрипты": ctk.CTkScrollableFrame(self.content_frame, fg_color="transparent"),
            "Логи": ctk.CTkFrame(self.content_frame, fg_color="transparent")
        }
        self.scripts_frame = self.frames["Скрипты"]

        self.tab_buttons = {}
        self.tab_lines = {}
        
        for tab_name in ["Скрипты", "Логи"]:
            # Фрейм для кнопки и линии
            tab_wrapper = ctk.CTkFrame(self.tabs_container, fg_color="transparent")
            tab_wrapper.pack(side="left", padx=(0, 2))

            btn = ctk.CTkButton(
                tab_wrapper,
                text=tab_name,
                width=100,
                height=30,
                corner_radius=5,
                font=("Arial", 14),
                command=lambda name=tab_name: self.switch_tab(name)
            )
            btn.pack(pady=(0, 0))
            
            line = ctk.CTkFrame(tab_wrapper, height=3, width=100)
            line.pack(fill="x")

            self.tab_buttons[tab_name] = btn
            self.tab_lines[tab_name] = line

        self.switch_tab("Скрипты")

        # Создаём карточки для каждого скрипта
        for script_id, config in self.scripts_config.items():
            self.create_script_card(script_id, config)

        # Фрейм для нижних кнопок
        bottom_buttons_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        bottom_buttons_frame.pack(pady=10)

        # Кнопка обновления статусов
        refresh_btn = ctk.CTkButton(
            bottom_buttons_frame,
            text="Обновить статусы",
            command=self.refresh_all_statuses,
            width=200
        )
        refresh_btn.pack(side="left", padx=10)

        # Кнопка глобальных настроек (email/telegram)
        global_settings_btn = ctk.CTkButton(
            bottom_buttons_frame,
            text="Глобальные настройки",
            command=self.open_global_settings_dialog,
            width=200,
            fg_color="#333333"
        )
        global_settings_btn.pack(side="left", padx=10)

        # Создаём интерфейс вкладки Логи
        self.create_logs_tab()

    def switch_tab(self, tab_name):
        active_bg = "#FFFFFF" # Excel белый
        active_text = "#000000"
        active_line = "#107C41" # Excel зеленый
        
        inactive_bg = "#333333" 
        inactive_text = "#AAAAAA"
        inactive_line = "transparent"

        for name, btn in self.tab_buttons.items():
            if name == tab_name:
                btn.configure(fg_color=active_bg, text_color=active_text, hover_color=active_bg)
                self.tab_lines[name].configure(fg_color=active_line)
                self.frames[name].pack(fill="both", expand=True)
            else:
                btn.configure(fg_color=inactive_bg, text_color=inactive_text, hover_color="#444444")
                self.tab_lines[name].configure(fg_color=inactive_line)
                self.frames[name].pack_forget()

    def create_logs_tab(self):
        """Создание элементов управления для вкладки 'Логи'"""
        logs_tab = self.frames["Логи"]
        
        # Кнопки управления логами
        btns_frame = ctk.CTkFrame(logs_tab, fg_color="transparent")
        btns_frame.pack(fill="x", padx=10, pady=5)
        
        refresh_logs_btn = ctk.CTkButton(
            btns_frame,
            text="Обновить логи",
            command=self.load_logs,
            width=150
        )
        refresh_logs_btn.pack(side="left", padx=5)

        # Текстовое поле для логов
        self.logs_textbox = ctk.CTkTextbox(
            logs_tab, 
            wrap="word",
            font=("Consolas", 12)
        )
        self.logs_textbox.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Сразу пытаемся загрузить логи
        self.load_logs()

    def get_log_file_path(self):
        """Определяет путь к файлу логов app.log (в соответствии с utils.py)"""
        current_dir = os.path.dirname(os.path.abspath(__file__))
        project_root = os.path.dirname(current_dir)
        if sys.platform == 'win32':
            app_data = os.environ.get('APPDATA')
            if app_data:
                return os.path.join(app_data, 'ScriptManager', 'logs', 'app.log')
        return os.path.join(project_root, "logs", "app.log")

    def load_logs(self):
        """Чтение файла лога и вывод его в Textbox"""
        log_path = self.get_log_file_path()
        self.logs_textbox.configure(state="normal")
        self.logs_textbox.delete("1.0", "end")

        if os.path.exists(log_path):
            try:
                with open(log_path, 'r', encoding='utf-8') as f:
                    # Читаем последние 500 строк
                    lines = f.readlines()
                    last_lines = lines[-500:]
                    content = "".join(last_lines)
                    if len(lines) > 500:
                        content = f"...(показаны последние 500 строк из {len(lines)})...\n\n" + content
                    self.logs_textbox.insert("end", content)
                    # Прокручиваем в самый низ
                    self.logs_textbox.yview("end")
            except Exception as e:
                self.logs_textbox.insert("end", f"Ошибка чтения логов: {e}")
        else:
            self.logs_textbox.insert("end", f"Файл лога не найден:\n{log_path}")

        self.logs_textbox.configure(state="disabled")

    def create_script_card(self, script_id: str, config: Dict[str, Any]):
        """
        Создание карточки скрипта

        Args:
            script_id: идентификатор скрипта
            config: конфигурация скрипта
        """
        # Рамка для скрипта
        card_frame = ctk.CTkFrame(
            self.scripts_frame,
            corner_radius=8,
            border_width=1,
            border_color="#444444"
        )
        card_frame.pack(pady=5, padx=10, fill="x")

        # Настраиваем grid для главной рамки (выравнивание кнопок по столбцам)
        card_frame.grid_columnconfigure(0, weight=1)  # Левая часть тянется
        # Остальные столбцы 1-5 имеют минимальный нужный размер
        
        # --- ЛЕВАЯ ЧАСТЬ (Имя + Статус) РОЛЬ 0 ---
        title_row = ctk.CTkFrame(card_frame, fg_color="transparent")
        title_row.grid(row=0, column=0, sticky="w", padx=15, pady=10)

        # Название скрипта
        name_label = ctk.CTkLabel(
            title_row,
            text=config['name'],
            font=("Arial", 16, "bold")
        )
        name_label.pack(side="left")

        # Индикатор статуса (кружок) - используем символ и цвет текста
        status_text = "●"
        status_color = "#2ECC71" if config['enabled'] else "#E74C3C"
        status_label = ctk.CTkLabel(
            title_row,
            text=status_text,
            text_color=status_color,
            font=("Arial", 16)
        )
        status_label.pack(side="left", padx=10)

        # --- КНОПКИ (РЯД 0, КОЛОНКИ 1-5) ---
        # Кнопка Настроить
        settings_btn = ctk.CTkButton(
            card_frame,
            text="⚙ Настроить",
            command=lambda sid=script_id: self.open_settings_dialog(sid),
            width=110
        )
        settings_btn.grid(row=0, column=1, padx=5, pady=10)

        # Кнопка Создать задачу
        create_task_btn = ctk.CTkButton(
            card_frame,
            text="➕ Создать задачу",
            command=lambda sid=script_id: self.create_scheduler_task(sid),
            width=140,
            fg_color="purple",
            hover_color="#6A1B9A"
        )
        create_task_btn.grid(row=0, column=2, padx=5, pady=10)

        # Кнопка Запустить (Run)
        run_btn = ctk.CTkButton(
            card_frame,
            text="▶ Запуск",
            command=lambda sid=script_id: self.run_script_now(sid),
            width=90,
            fg_color="#0066CC",
            hover_color="#0052A3"
        )
        run_btn.grid(row=0, column=3, padx=5, pady=10)

        # Кнопка Тестовая отправка
        test_btn = ctk.CTkButton(
            card_frame,
            text="✉ Тест",
            command=lambda sid=script_id: self.test_send(sid),
            width=80,
            fg_color="#F39C12",
            hover_color="#D68910"
        )
        test_btn.grid(row=0, column=4, padx=5, pady=10)

        # Кнопка Включить/Отключить
        toggle_text = "Отключить" if config['enabled'] else "Включить"
        toggle_color = "#E74C3C" if config['enabled'] else "#2ECC71"
        hover_color = "#C0392B" if config['enabled'] else "#27AE60"
        
        toggle_btn = ctk.CTkButton(
            card_frame,
            text=toggle_text,
            command=lambda sid=script_id: self.toggle_script(sid),
            width=100,
            fg_color=toggle_color,
            hover_color=hover_color
        )
        toggle_btn.grid(row=0, column=5, padx=(5, 15), pady=10)

        # --- НИЖНЯЯ ЧАСТЬ (Информация) РЯД 1 ---
        task_status = self.scheduler_manager.get_task_status(script_id)

        if task_status and config['enabled']:
            next_run = task_status.get('next_run')
            last_run = task_status.get('last_run')
            last_result = task_status.get('last_result', 0)

            result_icon = "✓" if last_result == 0 else "✗"
            result_color = "green" if last_result == 0 else "red"

            info_text = f"Следующий запуск: {next_run.strftime('%Y-%m-%d %H:%M') if next_run else 'Не запланирован'}"
            if last_run:
                info_text += f" | Последний: {last_run.strftime('%Y-%m-%d %H:%M')} {result_icon}"
        else:
            info_text = "Задача не создана или отключена"
            result_color = "gray"

        info_label = ctk.CTkLabel(
            card_frame,
            text=info_text,
            font=("Arial", 12),
            text_color=result_color if task_status else "gray",
            anchor="w",
            justify="left"
        )
        info_label.grid(row=1, column=0, columnspan=6, sticky="ew", padx=15, pady=(0, 10))

        # Сохраняем ссылки на виджеты для обновления
        self.script_widgets[script_id] = {
            'status_label': status_label,
            'toggle_btn': toggle_btn,
            'info_label': info_label
        }

    def toggle_script(self, script_id: str):
        """
        Включить/отключить скрипт

        Args:
            script_id: идентификатор скрипта
        """
        config = self.config_manager.load_script_config(script_id)
        new_state = not config['enabled']

        # Обновляем конфиг
        self.config_manager.update_script_field(script_id, 'enabled', new_state)

        # Обновляем Task Scheduler
        if new_state:
            self.scheduler_manager.enable_task(script_id)
        else:
            self.scheduler_manager.disable_task(script_id)

        # Обновляем интерфейс
        self.refresh_script_status(script_id)

        logging.info(f"Скрипт {script_id} {'включен' if new_state else 'отключен'}")

    def get_python_executable(self):
        """
        Получение правильного пути к интерпретатору Python.
        Если используется Embedded Python, возвращает путь к нему.
        Иначе возвращает sys.executable.
        """
        # Если запущено через start.bat, sys.executable уже указывает на python_embedded/python.exe
        # Если запущено через IDE/venv, указывает на venv python.exe
        # В обоих случаях sys.executable - правильный выбор.
        return sys.executable

    def get_runner_script_path(self):
        """Получение пути к скрипту runner.py"""
        return os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "src", "scripts", "runner.py"
        )

    def create_scheduler_task(self, script_id: str):
        """
        Создать задачу в Windows Task Scheduler

        Args:
            script_id: идентификатор скрипта
        """
        config = self.config_manager.load_script_config(script_id)

        # Формируем путь к runner.py
        runner_path = self.get_runner_script_path()
        python_exe = self.get_python_executable()

        # Вычисляем следующий запуск
        schedule = config['schedule']
        send_time = schedule.get('send_time', '09:00')

        # Парсим время
        hour, minute = map(int, send_time.split(':'))

        # Определяем дату следующего запуска
        from scripts.base_script import BaseScript
        from datetime import datetime, timedelta

        # Вычисляем дату запуска без использования экземпляра BaseScript
        now = datetime.now()

        if schedule['type'] == 'daily':
            # Ежедневно - завтра в указанное время
            next_run = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
            if next_run <= now:
                next_run += timedelta(days=1)
        else:
            # Еженедельно - следующий указанный день недели
            send_day = schedule.get('send_day', 'Wednesday') # Default fallback
            days_map = {
                'Monday': 0, 'Tuesday': 1, 'Wednesday': 2, 'Thursday': 3,
                'Friday': 4, 'Saturday': 5, 'Sunday': 6
            }
            # Fallback if send_day is missing or invalid
            target_day = days_map.get(send_day, 3) 
            current_day = now.weekday()
            days_ahead = (target_day - current_day) % 7
            if days_ahead == 0:
                days_ahead = 7  # Следующая неделя
            next_run = now + timedelta(days=days_ahead)
            next_run = next_run.replace(hour=hour, minute=minute, second=0, microsecond=0)

        # Создаём задачу
        success = self.scheduler_manager.create_task(
            script_id=script_id,
            script_path=runner_path,
            script_args=script_id,
            next_run_time=next_run,
            description=f"Автоматический запуск скрипта {config['name']}",
            python_exe=python_exe,
            enabled=config['enabled']
        )

        if success:
            state_text = "активна" if config['enabled'] else "отключена (как в настройках)"
            messagebox.showinfo("Успех", f"Задача для {config['name']} создана!\nСтатус: {state_text}\nСледующий запуск: {next_run.strftime('%Y-%m-%d %H:%M')}")
            self.refresh_script_status(script_id)
        else:
            messagebox.showerror("Ошибка", f"Не удалось создать задачу для {config['name']}")

    def test_send(self, script_id: str):
        """
        Тестовая отправка - запускает скрипт вручную

        Args:
            script_id: идентификатор скрипта
        """
        config = self.config_manager.load_script_config(script_id)

        # Диалог для ввода тестового email
        dialog = ctk.CTkInputDialog(
            text=f"Введите Email для тестовой отправки {config['name']}:\n(Оставьте пустым для использования адреса из настроек)",
            title="Тестовая отправка"
        )
        test_email = dialog.get_input()
        
        # Если нажали Cancel (None) -> выход. Если пустая строка -> продолжаем со стандартным email
        if test_email is None:
            return

        # Формируем путь к runner.py
        runner_path = self.get_runner_script_path()
        python_exe = self.get_python_executable()

        # Подготовка окружения
        env = os.environ.copy()
        if test_email.strip():
            env['TEST_EMAIL_RECIPIENT'] = test_email.strip()

        # Запускаем скрипт в отдельном процессе
        try:
            # Показываем окно ожидания
            wait_dialog = ctk.CTkToplevel(self.root)
            wait_dialog.title("Выполнение...")
            wait_dialog.geometry("350x120")
            wait_dialog.transient(self.root)
            wait_dialog.grab_set()
            
            # Центрируем
            wait_dialog.update_idletasks()
            x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (350 // 2)
            y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (120 // 2)
            wait_dialog.geometry(f"+{x}+{y}")

            label_text = f"Выполняется тестовая отправка...\n{config['name']}"
            if test_email.strip():
                label_text += f"\n\nПолучатель: {test_email.strip()}"
            
            ctk.CTkLabel(
                wait_dialog,
                text=label_text,
                font=("Arial", 12)
            ).pack(expand=True)

            wait_dialog.update()

            # Запускаем, используя полный путь к python
            process = subprocess.run(
                [python_exe, runner_path, script_id],
                capture_output=True,
                text=True,
                timeout=300,  # 5 минут максимум
                env=env,      # Передаем окружение с TEST_EMAIL_RECIPIENT
                # Важно: для скрытия консоли при запуске из GUI (если нужно)
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
            )

            wait_dialog.destroy()

            # Проверяем результат
            if process.returncode == 0:
                msg = f"Тестовая отправка для {config['name']} выполнена успешно!"
                if test_email.strip():
                    msg += f"\n\nПисьмо отправлено на: {test_email.strip()}"
                messagebox.showinfo("Успех", msg)
            else:
                # Если stderr пустой, возможно ошибка в stdout
                error_msg = process.stderr if process.stderr else process.stdout
                if not error_msg:
                    error_msg = "Неизвестная ошибка (пустой вывод)"
                
                messagebox.showerror(
                    "Ошибка",
                    f"Ошибка при выполнении:\n{error_msg[:1000]}"
                )

        except subprocess.TimeoutExpired:
            if 'wait_dialog' in locals():
                wait_dialog.destroy()
            messagebox.showerror("Ошибка", "Превышено время ожидания (5 минут)")

        except Exception as e:
            if 'wait_dialog' in locals():
                wait_dialog.destroy()
            messagebox.showerror("Ошибка", f"Не удалось запустить скрипт:\n{str(e)}")

        finally:
            # Обновляем статус
            self.refresh_script_status(script_id)

    def run_script_now(self, script_id: str):
        """
        Ручной запуск скрипта (боевой режим)
        """
        config = self.config_manager.load_script_config(script_id)

        # Подтверждение
        if not messagebox.askyesno("Подтверждение", f"Вы уверены, что хотите запустить {config['name']} сейчас?\n\nБудет выполнена полная отправка данных на {config['email']['recipient']}."):
            return

        # Формируем путь к runner.py
        runner_path = self.get_runner_script_path()
        python_exe = self.get_python_executable()

        # Запускаем скрипт в отдельном процессе
        try:
            # Показываем окно ожидания
            wait_dialog = ctk.CTkToplevel(self.root)
            wait_dialog.title("Выполнение...")
            wait_dialog.geometry("350x120")
            wait_dialog.transient(self.root)
            wait_dialog.grab_set()
            
            # Центрируем
            wait_dialog.update_idletasks()
            x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (350 // 2)
            y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (120 // 2)
            wait_dialog.geometry(f"+{x}+{y}")

            ctk.CTkLabel(
                wait_dialog,
                text=f"Выполнение скрипта {config['name']}...",
                font=("Arial", 12)
            ).pack(expand=True)

            wait_dialog.update()

            # Запускаем
            process = subprocess.run(
                [python_exe, runner_path, script_id],
                capture_output=True,
                text=True,
                timeout=600,  # 10 минут
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
            )

            wait_dialog.destroy()

            # Проверяем результат
            if process.returncode == 0:
                messagebox.showinfo("Успех", f"Скрипт {config['name']} выполнен успешно!")
            else:
                error_msg = process.stderr if process.stderr else process.stdout
                if not error_msg:
                    error_msg = "Неизвестная ошибка (пустой вывод)"
                
                messagebox.showerror(
                    "Ошибка",
                    f"Ошибка при выполнении:\n{error_msg[:1000]}"
                )

        except subprocess.TimeoutExpired:
            if 'wait_dialog' in locals():
                wait_dialog.destroy()
            messagebox.showerror("Ошибка", "Превышено время ожидания")

        except Exception as e:
            if 'wait_dialog' in locals():
                wait_dialog.destroy()
            messagebox.showerror("Ошибка", f"Не удалось запустить скрипт:\n{str(e)}")

        finally:
            self.refresh_script_status(script_id)

    def refresh_script_status(self, script_id: str):
        """
        Обновить отображение статуса скрипта
        
        Args:
            script_id: идентификатор скрипта
        """
        config = self.config_manager.load_script_config(script_id)
        widgets = self.script_widgets.get(script_id)

        if not widgets:
            return

        # Обновляем статус
        status_text = "●"
        status_color = "#2ECC71" if config['enabled'] else "#E74C3C"
        
        try:
            widgets['status_label'].configure(text=status_text, text_color=status_color)
        except Exception:
            widgets['status_label'].configure(text=status_text, text_color=status_color)

        # Обновляем кнопку
        toggle_text = "Отключить" if config['enabled'] else "Включить"
        toggle_color = "#E74C3C" if config['enabled'] else "#2ECC71"
        hover_color = "#C0392B" if config['enabled'] else "#27AE60"
        widgets['toggle_btn'].configure(
            text=toggle_text,
            fg_color=toggle_color,
            hover_color=hover_color
        )

        # Обновляем информацию о запуске
        task_status = self.scheduler_manager.get_task_status(script_id)

        if task_status and config['enabled']:
            next_run = task_status.get('next_run')
            last_run = task_status.get('last_run')
            last_result = task_status.get('last_result', 0)

            result_icon = "✓" if last_result == 0 else "✗"
            result_color = "green" if last_result == 0 else "red"

            info_text = f"Следующий запуск: {next_run.strftime('%Y-%m-%d %H:%M') if next_run else 'Не запланирован'}"
            if last_run:
                info_text += f" | Последний: {last_run.strftime('%Y-%m-%d %H:%M')} {result_icon}"
        else:
            info_text = "Задача не создана или отключена"
            result_color = "gray"

        widgets['info_label'].configure(text=info_text, text_color=result_color)

    def refresh_all_statuses(self):
        """Обновить статусы всех скриптов"""
        for script_id in self.scripts_config.keys():
            self.refresh_script_status(script_id)
        logging.info("Статусы всех скриптов обновлены")

    def open_settings_dialog(self, script_id: str):
        """
        Открыть окно настроек скрипта (требует пароль)

        Args:
            script_id: идентификатор скрипта
        """
        # Запрашиваем пароль
        password_dialog = PasswordDialog(self.root)
        self.root.wait_window(password_dialog)

        if not password_dialog.result:
            return

        # Пароль верный - открываем настройки
        config = self.config_manager.load_script_config(script_id)
        day_combo = None

        # Создаём модальное окно
        dialog = ctk.CTkToplevel(self.root)
        dialog.title(f"Настройки: {config['name']}")
        dialog.geometry("700x650")
        dialog.transient(self.root)
        dialog.grab_set()

        # Кнопки (упаковываем снизу)
        buttons_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        buttons_frame.pack(side="bottom", pady=10)

        def save_settings():
            """Сохранение настроек"""
            # Обновляем пути
            if 'input_file' in paths:
                paths['input_file'] = source_entry.get() if 'source_entry' in locals() else paths.get('input_file', '')
            if 'input_folder' in paths:
                paths['input_folder'] = input_folder_entry.get() if 'input_folder_entry' in locals() else paths.get('input_folder', '')
            if 'summary_file' in paths:
                paths['summary_file'] = summary_entry.get() if 'summary_entry' in locals() else paths.get('summary_file', '')

            paths['output_folder'] = output_entry.get()

            if 'sheet_name' in paths:
                paths['sheet_name'] = sheet_entry.get() if 'sheet_entry' in locals() else paths.get('sheet_name', '')

            # Обновляем email
            email['recipient'] = email_entry.get()
            email['subject_template'] = subject_entry.get()
            email['body_template'] = body_text.get("1.0", "end-1c")

            # Обновляем расписание
            schedule['send_time'] = time_entry.get()
            schedule['type'] = type_combo.get()
            schedule['smart_reschedule'] = smart_var.get()

            if day_combo:
                schedule['send_day'] = day_combo.get()

            # Сохраняем
            self.config_manager.save_script_config(script_id, config)
            logging.info(f"Настройки скрипта {script_id} сохранены")

            # Обновляем интерфейс
            self.refresh_script_status(script_id)

            # Закрываем окно
            messagebox.showinfo("Успех", "Настройки сохранены!")
            dialog.destroy()

        ctk.CTkButton(buttons_frame, text="Сохранить", command=save_settings, width=120).pack(side="left", padx=5)
        ctk.CTkButton(buttons_frame, text="Отмена", command=dialog.destroy, width=120).pack(side="left", padx=5)

        # Контейнер с прокруткой
        scroll_frame = ctk.CTkScrollableFrame(dialog, width=650, height=550)
        scroll_frame.pack(pady=20, padx=20, fill="both", expand=True)

        # === Пути ===
        ctk.CTkLabel(scroll_frame, text="Пути к файлам", font=("Arial", 14, "bold")).pack(anchor="w", pady=(0, 10))

        paths = config.get('paths', {})

        # Исходный файл (если есть)
        if 'input_file' in paths:
            ctk.CTkLabel(scroll_frame, text="Исходный файл:").pack(anchor="w")
            source_frame = ctk.CTkFrame(scroll_frame, fg_color="transparent")
            source_frame.pack(fill="x", pady=(0, 10))

            source_entry = ctk.CTkEntry(source_frame, width=480)
            source_entry.insert(0, paths.get('input_file', ''))
            source_entry.pack(side="left", padx=(0, 5))

            ctk.CTkButton(
                source_frame,
                text="Browse...",
                width=100,
                command=lambda e=source_entry: self.browse_file(e, "Выберите исходный файл")
            ).pack(side="left")

        # Папка для исходных файлов (для CMS48)
        if 'input_folder' in paths:
            ctk.CTkLabel(scroll_frame, text="Папка исходных файлов:").pack(anchor="w")
            input_folder_frame = ctk.CTkFrame(scroll_frame, fg_color="transparent")
            input_folder_frame.pack(fill="x", pady=(0, 10))

            input_folder_entry = ctk.CTkEntry(input_folder_frame, width=480)
            input_folder_entry.insert(0, paths.get('input_folder', ''))
            input_folder_entry.pack(side="left", padx=(0, 5))

            ctk.CTkButton(
                input_folder_frame,
                text="Browse...",
                width=100,
                command=lambda e=input_folder_entry: self.browse_directory(e, "Выберите папку")
            ).pack(side="left")

        # Итоговая аршин ЦМС (summary_file)
        if 'summary_file' in paths:
            ctk.CTkLabel(scroll_frame, text="Файл 'Итоговая аршин ЦМС':").pack(anchor="w")
            summary_frame = ctk.CTkFrame(scroll_frame, fg_color="transparent")
            summary_frame.pack(fill="x", pady=(0, 10))

            summary_entry = ctk.CTkEntry(summary_frame, width=480)
            summary_entry.insert(0, paths.get('summary_file', ''))
            summary_entry.pack(side="left", padx=(0, 5))

            ctk.CTkButton(
                summary_frame,
                text="Browse...",
                width=100,
                command=lambda e=summary_entry: self.browse_file(e, "Выберите итоговый файл")
            ).pack(side="left")

        # Путь сохранения
        ctk.CTkLabel(scroll_frame, text="Папка для сохранения отчётов:").pack(anchor="w")
        output_frame = ctk.CTkFrame(scroll_frame, fg_color="transparent")
        output_frame.pack(fill="x", pady=(0, 10))

        output_entry = ctk.CTkEntry(output_frame, width=480)
        output_entry.insert(0, paths.get('output_folder', ''))
        output_entry.pack(side="left", padx=(0, 5))

        ctk.CTkButton(
            output_frame,
            text="Browse...",
            width=100,
            command=lambda e=output_entry: self.browse_directory(e, "Выберите папку для сохранения")
        ).pack(side="left")

        # Имя листа
        if 'sheet_name' in paths:
            ctk.CTkLabel(scroll_frame, text="Имя листа Excel:").pack(anchor="w")
            sheet_entry = ctk.CTkEntry(scroll_frame, width=300)
            sheet_entry.insert(0, paths.get('sheet_name', ''))
            sheet_entry.pack(anchor="w", pady=(0, 10))

        # === Email настройки ===
        ctk.CTkLabel(scroll_frame, text="Email настройки", font=("Arial", 14, "bold")).pack(anchor="w", pady=(10, 10))

        email = config.get('email', {})

        # Email получателя
        ctk.CTkLabel(scroll_frame, text="Email получателя:").pack(anchor="w")
        email_entry = ctk.CTkEntry(scroll_frame, width=500)
        email_entry.insert(0, email.get('recipient', ''))
        email_entry.pack(fill="x", pady=(0, 10))

        # Тема письма
        ctk.CTkLabel(scroll_frame, text="Шаблон темы письма:").pack(anchor="w")
        subject_entry = ctk.CTkEntry(scroll_frame, width=500)
        subject_entry.insert(0, email.get('subject_template', ''))
        subject_entry.pack(fill="x", pady=(0, 10))

        # Тело письма
        ctk.CTkLabel(scroll_frame, text="Текст письма:").pack(anchor="w")
        body_text = ctk.CTkTextbox(scroll_frame, width=600, height=100)
        body_text.insert("1.0", email.get('body_template', ''))
        body_text.pack(fill="x", pady=(0, 10))

        # === Настройки расписания ===
        ctk.CTkLabel(scroll_frame, text="Настройки расписания", font=("Arial", 14, "bold")).pack(anchor="w", pady=(10, 10))

        schedule = config.get('schedule', {})

        # Время запуска
        ctk.CTkLabel(scroll_frame, text="Время запуска (чч:мм):").pack(anchor="w")
        time_entry = ctk.CTkEntry(scroll_frame, width=100)
        time_entry.insert(0, schedule.get('send_time', '09:00'))
        time_entry.pack(anchor="w", pady=(0, 10))

        # Тип расписания
        ctk.CTkLabel(scroll_frame, text="Тип расписания:").pack(anchor="w")
        type_combo = ctk.CTkComboBox(scroll_frame, values=['daily', 'weekly'], state="readonly")
        type_combo.set(schedule.get('type', 'daily'))
        type_combo.pack(anchor="w", pady=(0, 10))

        # Контейнер для настроек weekly (чтобы сохранять порядок при скрытии/показе)
        weekly_frame = ctk.CTkFrame(scroll_frame, fg_color="transparent")
        weekly_frame.pack(fill="x", pady=0)

        # Элементы внутри weekly_frame
        week_day_label = ctk.CTkLabel(weekly_frame, text="День недели:")
        
        send_day = schedule.get('send_day', 'Thursday')
        days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
        day_combo = ctk.CTkComboBox(weekly_frame, values=days, state="readonly")
        day_combo.set(send_day)

        def update_visibility(choice):
            if type_combo.get() == 'weekly':
                week_day_label.pack(anchor="w")
                day_combo.pack(anchor="w", pady=(0, 10))
            else:
                week_day_label.pack_forget()
                day_combo.pack_forget()

        type_combo.configure(command=update_visibility)
        update_visibility(None) # Инициализация видимости

        # Умное перепланирование
        smart_var = ctk.BooleanVar(value=schedule.get('smart_reschedule', True))
        ctk.CTkCheckBox(
            scroll_frame,
            text="Умное перепланирование (учитывать календарь РФ)",
            variable=smart_var
        ).pack(anchor="w", pady=(0, 10))

        # Кнопки перемещены наверх

    def browse_file(self, entry_widget, title: str):
        """
        Открыть диалог выбора файла

        Args:
            entry_widget: поле для записи пути
            title: заголовок диалога
        """
        filename = filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            entry_widget.delete(0, "end")
            entry_widget.insert(0, filename)

    def browse_directory(self, entry_widget, title: str):
        """
        Открыть диалог выбора папки

        Args:
            entry_widget: поле для записи пути
            title: заголовок диалога
        """
        directory = filedialog.askdirectory(title=title)
        if directory:
            entry_widget.delete(0, "end")
            entry_widget.insert(0, directory)

    def open_global_settings_dialog(self):
        """
        Открыть окно глобальных настроек (email, telegram)
        """
        # Запрашиваем пароль
        password_dialog = PasswordDialog(self.root)
        self.root.wait_window(password_dialog)

        if not password_dialog.result:
            return

        dialog = ctk.CTkToplevel(self.root)
        dialog.title("Глобальные настройки (Отправители и Telegram)")
        dialog.geometry("700x700")
        dialog.transient(self.root)
        dialog.grab_set()

        # Загружаем конфиги
        gmail_config = self.config_manager.load_email_config("gmail")
        yandex_config = self.config_manager.load_email_config("yandex")
        telegram_config = self.config_manager.load_telegram_config()

        def save_global_settings():
            # Обновляем Gmail
            gmail_config["smtp"]["username"] = gmail_user.get()
            gmail_config["smtp"]["password"] = gmail_pass.get()
            
            # Обновляем Yandex
            yandex_config["smtp"]["username"] = yandex_user.get()
            yandex_config["smtp"]["password"] = yandex_pass.get()
            
            # Обновляем Telegram
            telegram_config["bot_token"] = tg_token.get()
            telegram_config["personal_chat_id"] = tg_personal.get()
            telegram_config["group_chat_id"] = tg_group.get()
            
            # Сохраняем
            self.config_manager.save_email_config(gmail_config, "gmail")
            self.config_manager.save_email_config(yandex_config, "yandex")
            self.config_manager.save_telegram_config(telegram_config)
            
            logging.info("Глобальные настройки успешно сохранены")
            messagebox.showinfo("Успех", "Глобальные настройки сохранены!")
            dialog.destroy()

        # Кнопки (внизу)
        buttons_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        buttons_frame.pack(side="bottom", pady=10)
        ctk.CTkButton(buttons_frame, text="Сохранить", command=save_global_settings, width=120).pack(side="left", padx=5)
        ctk.CTkButton(buttons_frame, text="Отмена", command=dialog.destroy, width=120).pack(side="left", padx=5)

        # Контейнер с прокруткой
        scroll_frame = ctk.CTkScrollableFrame(dialog, width=650, height=600)
        scroll_frame.pack(pady=20, padx=20, fill="both", expand=True)

        # Почта Gmail
        ctk.CTkLabel(scroll_frame, text="Настройки отправителя (Gmail)", font=("Arial", 14, "bold")).pack(anchor="w", pady=(0, 10))
        
        ctk.CTkLabel(scroll_frame, text="Пользователь (Email):").pack(anchor="w")
        gmail_user = ctk.CTkEntry(scroll_frame, width=400)
        gmail_user.insert(0, gmail_config.get("smtp", {}).get("username", ""))
        gmail_user.pack(anchor="w", pady=(0, 5))
        
        ctk.CTkLabel(scroll_frame, text="Пароль приложения:").pack(anchor="w")
        gmail_pass = ctk.CTkEntry(scroll_frame, width=400, show="*")
        gmail_pass.insert(0, gmail_config.get("smtp", {}).get("password", ""))
        gmail_pass.pack(anchor="w", pady=(0, 20))

        # Почта Yandex
        ctk.CTkLabel(scroll_frame, text="Настройки отправителя (Yandex)", font=("Arial", 14, "bold")).pack(anchor="w", pady=(0, 10))
        
        ctk.CTkLabel(scroll_frame, text="Пользователь (Email):").pack(anchor="w")
        yandex_user = ctk.CTkEntry(scroll_frame, width=400)
        yandex_user.insert(0, yandex_config.get("smtp", {}).get("username", ""))
        yandex_user.pack(anchor="w", pady=(0, 5))
        
        ctk.CTkLabel(scroll_frame, text="Пароль приложения:").pack(anchor="w")
        yandex_pass = ctk.CTkEntry(scroll_frame, width=400, show="*")
        yandex_pass.insert(0, yandex_config.get("smtp", {}).get("password", ""))
        yandex_pass.pack(anchor="w", pady=(0, 20))

        # Настройки Telegram
        ctk.CTkLabel(scroll_frame, text="Настройки Telegram Bot", font=("Arial", 14, "bold")).pack(anchor="w", pady=(0, 10))
        
        ctk.CTkLabel(scroll_frame, text="Bot Token:").pack(anchor="w")
        tg_token = ctk.CTkEntry(scroll_frame, width=500)
        tg_token.insert(0, telegram_config.get("bot_token", ""))
        tg_token.pack(anchor="w", pady=(0, 5))
        
        ctk.CTkLabel(scroll_frame, text="Personal Chat ID:").pack(anchor="w")
        tg_personal = ctk.CTkEntry(scroll_frame, width=300)
        tg_personal.insert(0, telegram_config.get("personal_chat_id", ""))
        tg_personal.pack(anchor="w", pady=(0, 5))

        ctk.CTkLabel(scroll_frame, text="Group Chat ID:").pack(anchor="w")
        tg_group = ctk.CTkEntry(scroll_frame, width=300)
        tg_group.insert(0, telegram_config.get("group_chat_id", ""))
        tg_group.pack(anchor="w", pady=(0, 10))

    def run(self):
        """Запуск GUI"""
        logging.info("Запуск GUI...")
        self.root.mainloop()


if __name__ == "__main__":
    from utils import setup_logging
    setup_logging()

    app = ScriptManagerGUI()
    app.run()
