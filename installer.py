"""
Инсталлятор ScriptManager

Автоматически устанавливает ScriptManager в систему:
- Создаёт папку в Program Files
- Копирует все необходимые файлы
- Создаёт ярлык на рабочем столе
- Настраивает автозапуск (опционально)
- Регистрирует в реестре Windows для деинсталляции
"""

import os
import sys
import shutil
import winreg
import ctypes
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, filedialog
import subprocess

class ScriptManagerInstaller:
    """Инсталлятор ScriptManager"""

    APP_NAME = "ScriptManager"
    APP_VERSION = "1.0.0"
    APP_PUBLISHER = "CMS Arshin"

    def __init__(self):
        """Инициализация инсталлятора"""
        # Проверка прав администратора
        if not self.is_admin():
            messagebox.showerror(
                "Требуются права администратора",
                "Для установки запустите инсталлятор от имени администратора"
            )
            sys.exit(1)

        # Пути установки
        self.install_dir = Path(os.environ['ProgramFiles']) / self.APP_NAME
        self.desktop = Path.home() / "Desktop"
        self.startup_folder = Path(os.environ['APPDATA']) / "Microsoft/Windows/Start Menu/Programs/Startup"

        # Текущая директория (откуда запущен инсталлятор)
        if getattr(sys, 'frozen', False):
            # Если запущено как exe, ресурсы лежат рядом с exe
            self.source_dir = Path(sys.executable).parent
        else:
            self.source_dir = Path(__file__).parent

        # Создаём GUI
        self.create_gui()

    def is_admin(self):
        """Проверка прав администратора"""
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    def create_gui(self):
        """Создание интерфейса инсталлятора"""
        self.root = tk.Tk()
        self.root.title(f"Установка {self.APP_NAME} {self.APP_VERSION}")
        self.root.geometry("600x600")
        self.root.resizable(True, True)
        
        # Установка иконки
        icon_path = self.source_dir / "assets" / "ScriptM.ico"
        if icon_path.exists():
            try:
                self.root.iconbitmap(str(icon_path))
            except Exception:
                pass

        # Кнопки (упаковываем снизу, чтобы они всегда были видны)
        buttons_frame = tk.Frame(self.root)
        buttons_frame.pack(side="bottom", pady=20)

        tk.Button(
            buttons_frame,
            text="Установить",
            command=self.install,
            width=15,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 10, "bold")
        ).pack(side="left", padx=5)

        tk.Button(
            buttons_frame,
            text="Отмена",
            command=self.root.quit,
            width=15,
            font=("Arial", 10)
        ).pack(side="left", padx=5)

        # Заголовок
        header = tk.Label(
            self.root,
            text=f"Установка {self.APP_NAME}",
            font=("Arial", 16, "bold")
        )
        header.pack(pady=20)

        # Описание
        description = tk.Label(
            self.root,
            text="Система автоматизации отправки отчётов\nс интеграцией Windows Task Scheduler",
            font=("Arial", 10),
            justify="center"
        )
        description.pack(pady=10)

        # Путь установки
        path_frame = tk.Frame(self.root)
        path_frame.pack(pady=20, padx=20, fill="x")

        tk.Label(path_frame, text="Путь установки:", font=("Arial", 10)).pack(anchor="w")

        path_entry_frame = tk.Frame(path_frame)
        path_entry_frame.pack(fill="x", pady=5)

        self.path_var = tk.StringVar(value=str(self.install_dir))
        path_entry = tk.Entry(path_entry_frame, textvariable=self.path_var, font=("Arial", 9))
        path_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))

        tk.Button(
            path_entry_frame,
            text="Обзор...",
            command=self.browse_install_dir
        ).pack(side="left")

        # Опции
        options_frame = tk.Frame(self.root)
        options_frame.pack(pady=10, padx=30, fill="x")

        self.desktop_shortcut_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            options_frame,
            text="Создать ярлык на рабочем столе",
            variable=self.desktop_shortcut_var,
            font=("Arial", 9)
        ).pack(anchor="w")

        self.autostart_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            options_frame,
            text="Добавить в автозагрузку Windows",
            variable=self.autostart_var,
            font=("Arial", 9)
        ).pack(anchor="w")

        # Информация
        info_text = (
            "Будет установлено:\n"
            "• Приложение ScriptManager\n"
            "• Конфигурационные файлы\n"
            "• Документация"
        )
        info = tk.Label(
            self.root,
            text=info_text,
            font=("Arial", 9),
            justify="left",
            fg="gray"
        )
        info.pack(pady=15)

        # Кнопки перемещены наверх (pack side=bottom)

        self.root.mainloop()

    def browse_install_dir(self):
        """Выбор директории установки"""
        directory = filedialog.askdirectory(
            title="Выберите папку для установки",
            initialdir=self.install_dir.parent
        )
        if directory:
            self.path_var.set(directory)
            self.install_dir = Path(directory) / self.APP_NAME

    def install(self):
        """Процесс установки"""
        try:
            # Обновляем путь установки
            self.install_dir = Path(self.path_var.get())

            # Показываем прогресс
            progress = tk.Toplevel(self.root)
            progress.title("Установка...")
            progress.geometry("400x150")
            progress.transient(self.root)
            progress.grab_set()

            progress_label = tk.Label(progress, text="Установка ScriptManager...", font=("Arial", 10))
            progress_label.pack(pady=20)

            status_label = tk.Label(progress, text="", font=("Arial", 9), fg="gray")
            status_label.pack(pady=10)

            def update_status(text):
                status_label.config(text=text)
                progress.update()

            # 1. Создание директории
            update_status("Создание директории установки...")
            self.install_dir.mkdir(parents=True, exist_ok=True)

            # 2. Копирование файлов
            update_status("Копирование файлов...")
            self.copy_files()

            # 3. Создание ярлыка на рабочем столе
            if self.desktop_shortcut_var.get():
                update_status("Создание ярлыка на рабочем столе...")
                self.create_desktop_shortcut()

            # 4. Добавление в автозагрузку
            if self.autostart_var.get():
                update_status("Добавление в автозагрузку...")
                self.add_to_startup()

            # 5. Регистрация в реестре
            update_status("Регистрация в системе...")
            self.register_in_registry()

            # 6. Создание деинсталлятора
            update_status("Создание деинсталлятора...")
            self.create_uninstaller()

            progress.destroy()

            # Успех!
            messagebox.showinfo(
                "Установка завершена",
                f"{self.APP_NAME} успешно установлен!\n\n"
                f"Путь: {self.install_dir}\n\n"
                "Запустите приложение с ярлыка на рабочем столе."
            )

            self.root.quit()

        except Exception as e:
            messagebox.showerror("Ошибка установки", f"Не удалось установить приложение:\n{e}")

    def copy_files(self):
        """Копирование файлов приложения"""
        # Копируем основные директории
        dirs_to_copy = ['src', 'config', 'doc', 'assets']

        for dir_name in dirs_to_copy:
            source = self.source_dir / dir_name
            dest = self.install_dir / dir_name

            if source.exists():
                if dest.exists():
                    shutil.rmtree(dest)
                shutil.copytree(source, dest)

        # Копируем Embedded Python (если есть)
        python_embedded_source = self.source_dir / 'python_embedded'
        python_embedded_dest = self.install_dir / 'python_embedded'

        if python_embedded_source.exists():
            if python_embedded_dest.exists():
                shutil.rmtree(python_embedded_dest)
            shutil.copytree(python_embedded_source, python_embedded_dest)

        # Копируем отдельные файлы
        files_to_copy = [
            'requirements.txt',
            'README.md',
            'QUICKSTART.md',
            'PROJECT_SUMMARY.md'
        ]

        for file_name in files_to_copy:
            source = self.source_dir / file_name
            dest = self.install_dir / file_name

            if source.exists():
                shutil.copy2(source, dest)

        # Папка logs теперь создается в APPDATA при первом запуске
        # (self.install_dir / 'logs').mkdir(exist_ok=True)

    def create_desktop_shortcut(self):
        """Создание ярлыка на рабочем столе"""
        shortcut_path = self.desktop / f"{self.APP_NAME}.lnk"
        
        # Определяем пути
        embedded_python = self.install_dir / "python_embedded" / "python.exe"
        embedded_pythonw = self.install_dir / "python_embedded" / "pythonw.exe"
        gui_path = self.install_dir / "src" / "gui.py"

        # Логика выбора цели ярлыка
        target_path = ""
        arguments = ""

        if embedded_pythonw.exists():
            # Используем pythonw.exe - нативное скрытие консоли
            target_path = str(embedded_pythonw)
            arguments = f'"{str(gui_path)}"'
        elif embedded_python.exists():
            # Fallback к VBS
            vbs_path = self.install_dir / "start_hidden.vbs"
            python_path = str(embedded_python)
            
            with open(vbs_path, 'w') as f:
                f.write('Set WshShell = CreateObject("WScript.Shell")\n')
                f.write(f'WshShell.CurrentDirectory = "{str(self.install_dir)}"\n')
                f.write(f'WshShell.Run """{python_path}"" ""{str(gui_path)}""", 0, False\n')
            
            target_path = str(vbs_path)
            arguments = ""
        else:
            # Системный python
            target_path = "pythonw.exe" 
            arguments = f'"{str(gui_path)}"'

        # Создаём .bat файл для запуска
        bat_path = self.install_dir / "start.bat"

        with open(bat_path, 'w') as f:
            f.write(f'@echo off\n')
            f.write(f'cd /d "{self.install_dir}"\n')

            if embedded_python.exists():
                # Используем встроенный Python
                f.write(f'python_embedded\\python.exe src\\gui.py\n')
            else:
                # Используем системный Python
                f.write(f'python src\\gui.py\n')
            f.write('pause\n')

        # Создаём ярлык через PowerShell
        ps_script = f'''
        $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut("{shortcut_path}")
        $Shortcut.TargetPath = "{target_path}"
        $Shortcut.Arguments = '{arguments}'
        $Shortcut.WorkingDirectory = "{self.install_dir}"
        $Shortcut.IconLocation = "{self.install_dir / 'assets' / 'ScriptM.ico'}"
        $Shortcut.Description = "ScriptManager - Управление скриптами"
        $Shortcut.Save()
        '''

        subprocess.run(['powershell', '-Command', ps_script], check=True)

    def add_to_startup(self):
        """Добавление в автозагрузку Windows"""
        startup_shortcut = self.startup_folder / f"{self.APP_NAME}.lnk"
        bat_path = self.install_dir / "start.bat"

        ps_script = f'''
        $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut("{startup_shortcut}")
        $Shortcut.TargetPath = "{bat_path}"
        $Shortcut.WorkingDirectory = "{self.install_dir}"
        $Shortcut.WindowStyle = 7
        $Shortcut.Save()
        '''

        subprocess.run(['powershell', '-Command', ps_script], check=True)

    def register_in_registry(self):
        """Регистрация приложения в реестре Windows"""
        uninstall_key = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        app_key = f"{uninstall_key}\\{self.APP_NAME}"

        try:
            # Создаём ключ
            key = winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, app_key)

            # Заполняем информацию
            winreg.SetValueEx(key, "DisplayName", 0, winreg.REG_SZ, self.APP_NAME)
            winreg.SetValueEx(key, "DisplayVersion", 0, winreg.REG_SZ, self.APP_VERSION)
            winreg.SetValueEx(key, "Publisher", 0, winreg.REG_SZ, self.APP_PUBLISHER)
            winreg.SetValueEx(key, "InstallLocation", 0, winreg.REG_SZ, str(self.install_dir))
            winreg.SetValueEx(key, "UninstallString", 0, winreg.REG_SZ, f'"{str(self.install_dir / "uninstall.bat")}"')
            winreg.SetValueEx(key, "DisplayIcon", 0, winreg.REG_SZ, str(self.install_dir / "assets" / "ScriptM.ico"))

            winreg.CloseKey(key)
        except Exception as e:
            print(f"Ошибка регистрации в реестре: {e}")

    def create_uninstaller(self):
        """Создание деинсталлятора (Batch версия)"""
        uninstall_bat = self.install_dir / "uninstall.bat"
        
        try:
            with open(uninstall_bat, 'w', encoding='cp866') as f: # cp866 для корректного отображения кириллицы в консоли
                f.write('@echo off\n')
                f.write('echo ================================================\n')
                f.write('echo   ScriptManager - Uninstall\n')
                f.write('echo ================================================\n')
                f.write('echo.\n')
                f.write('set /p CONFIRM=Are you sure you want to uninstall ScriptManager? (Y/n): \n')
                f.write('if /i "%CONFIRM%"=="n" (\n')
                f.write('    echo Uninstall canceled.\n')
                f.write('    pause\n')
                f.write('    exit /b 0\n')
                f.write(')\n\n')

                # Создаем временный скрипт удаления, чтобы можно было удалить саму папку
                f.write('echo.\n')
                f.write('echo Preparing to uninstall...\n')
                f.write('set TEMP_UNINSTALL=%TEMP%\\uninstall_scriptmanager.bat\n')
                
                # Генерируем содержимое временного скрипта
                f.write('echo @echo off > "%TEMP_UNINSTALL%"\n')
                f.write('echo echo [1/4] Removing program files... >> "%TEMP_UNINSTALL%"\n')
                # Удаляем директорию установки
                f.write(f'echo rd /s /q "{self.install_dir}" >> "%TEMP_UNINSTALL%"\n')
                f.write('echo echo      OK >> "%TEMP_UNINSTALL%"\n')
                
                f.write('echo echo [2/4] Removing shortcuts... >> "%TEMP_UNINSTALL%"\n')
                # Удаляем ярлыки
                f.write(f'echo del /f /q "{self.desktop}\\ScriptManager.lnk" >> "%TEMP_UNINSTALL%"\n')
                startup_lnk = self.startup_folder / f"{self.APP_NAME}.lnk"
                f.write(f'echo del /f /q "{startup_lnk}" >> "%TEMP_UNINSTALL%"\n')
                f.write('echo echo      OK >> "%TEMP_UNINSTALL%"\n')
                
                f.write('echo echo [3/4] Removing registry entries... >> "%TEMP_UNINSTALL%"\n')
                # Удаляем из реестра
                f.write('echo reg delete "HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\ScriptManager" /f ^>nul 2^>^&1 >> "%TEMP_UNINSTALL%"\n')
                f.write('echo echo      OK >> "%TEMP_UNINSTALL%"\n')
                
                f.write('echo echo [4/4] Cleanup finished. >> "%TEMP_UNINSTALL%"\n')
                f.write('echo echo. >> "%TEMP_UNINSTALL%"\n')
                f.write('echo echo ScriptManager has been successfully uninstalled. >> "%TEMP_UNINSTALL%"\n')
                f.write('echo echo. >> "%TEMP_UNINSTALL%"\n')
                f.write('echo pause >> "%TEMP_UNINSTALL%"\n')
                # Самоудаление временного скрипта
                f.write('echo del "%%~f0" >> "%TEMP_UNINSTALL%"\n')
                
                f.write('\n')
                f.write('start "" "%TEMP_UNINSTALL%"\n')
                f.write('exit\n')

        except Exception as e:
            print(f"Error creating uninstaller: {e}")


if __name__ == "__main__":
    installer = ScriptManagerInstaller()
