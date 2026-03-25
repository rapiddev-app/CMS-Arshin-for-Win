"""
Простой инсталлятор ScriptManager (консольный)
"""

import os
import sys
import shutil
import winreg
import ctypes
from pathlib import Path
import subprocess

class SimpleInstaller:
    """Упрощённый консольный инсталлятор"""

    APP_NAME = "ScriptManager"
    APP_VERSION = "1.0.0"
    APP_PUBLISHER = "CMS Arshin"

    def __init__(self):
        # Проверка прав администратора
        if not self.is_admin():
            print("\nОШИБКА: Требуются права администратора!")
            print("Запустите install.bat от имени администратора")
            input("\nНажмите Enter для выхода...")
            sys.exit(1)

        # Пути
        self.install_dir = Path(os.environ['ProgramFiles']) / self.APP_NAME
        self.desktop = Path.home() / "Desktop"
        if getattr(sys, 'frozen', False):
            self.source_dir = Path(sys.executable).parent
        else:
            self.source_dir = Path(__file__).parent

    def is_admin(self):
        """Проверка прав администратора"""
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    def kill_running_processes(self):
        """Принудительное завершение процессов ScriptManager"""
        try:
            # Ищем процессы python.exe и pythonw.exe, запущенные из папки установки
            # Используем PowerShell для фильтрации по пути
            
            # Экранируем путь для PowerShell
            install_path_str = str(self.install_dir).replace("'", "''")
            
            ps_script = f"""
            Get-WmiObject Win32_Process | Where-Object {{ $_.ExecutablePath -and $_.ExecutablePath.StartsWith('{install_path_str}') }} | ForEach-Object {{ Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }}
            """
            
            subprocess.run(["powershell", "-Command", ps_script], check=False, capture_output=True)
            
        except Exception as e:
            print(f"      Предупреждение при остановке процессов: {e}")

    def install(self):
        """Процесс установки"""
        print("="*60)
        print(f"  Установка {self.APP_NAME} v{self.APP_VERSION}")
        print("="*60)
        print()
        print(f"Путь установки: {self.install_dir}")
        print()

        # Подтверждение
        answer = input("Продолжить установку? (Y/n): ").strip().lower()
        if answer and answer != 'y' and answer != 'д':
            print("\nУстановка отменена.")
            return False

        try:
            print()
            print("[1/8] Остановка запущенных процессов...")
            self.kill_running_processes()
            print("      OK")

            print("[2/8] Создание директории...")
            self.install_dir.mkdir(parents=True, exist_ok=True)
            print("      OK")

            print("[2/7] Копирование файлов...")
            self.copy_files()
            print("      OK")

            print("[3/7] Создание ярлыка на рабочем столе...")
            self.create_desktop_shortcut()
            print("      OK")

            print("[4/7] Регистрация в системе...")
            self.register_in_registry()
            print("      OK")

            print("[5/7] Создание деинсталлятора...")
            self.create_uninstaller()
            print("      OK")

            # Папка logs теперь создается в APPDATA
            # print("[6/7] Создание папки для логов...")
            # (self.install_dir / 'logs').mkdir(exist_ok=True)
            # print("      OK")

            print("[7/7] Финализация...")
            print("      OK")

            print()
            print("="*60)
            print("  Установка завершена успешно!")
            print("="*60)
            print()
            print(f"ScriptManager установлен в: {self.install_dir}")
            print(f"Ярлык создан на рабочем столе")
            print()
            print("Следующие шаги:")
            print("1. Откройте ScriptManager с ярлыка")
            print("2. Настройте config/email.json и config/telegram.json")
            print("3. Настройте каждый скрипт через GUI (пароль: cms2025)")
            print()

            return True

        except Exception as e:
            print(f"\nОШИБКА УСТАНОВКИ: {e}")
            import traceback
            traceback.print_exc()
            return False

    def copy_files(self):
        """Копирование файлов"""
        # Основные директории
        dirs_to_copy = ['src', 'config', 'doc', 'assets']
        for dir_name in dirs_to_copy:
            source = self.source_dir / dir_name
            dest = self.install_dir / dir_name
            if source.exists():
                if dest.exists():
                    shutil.rmtree(dest)
                shutil.copytree(source, dest)

        # Embedded Python
        python_embedded_source = self.source_dir / 'python_embedded'
        if python_embedded_source.exists():
            python_embedded_dest = self.install_dir / 'python_embedded'
            if python_embedded_dest.exists():
                shutil.rmtree(python_embedded_dest)
            shutil.copytree(python_embedded_source, python_embedded_dest)

        # Отдельные файлы
        files_to_copy = ['requirements.txt', 'README.md', 'QUICKSTART.md', 'PROJECT_SUMMARY.md']
        for file_name in files_to_copy:
            source = self.source_dir / file_name
            if source.exists():
                shutil.copy2(source, self.install_dir / file_name)

    def create_desktop_shortcut(self):
        """Создание ярлыка"""
        shortcut_path = self.desktop / f"{self.APP_NAME}.lnk"
        
        # Определяем пути
        embedded_python = self.install_dir / "python_embedded" / "python.exe"
        embedded_pythonw = self.install_dir / "python_embedded" / "pythonw.exe"
        gui_path = self.install_dir / "src" / "gui.py"

        # Логика выбора цели ярлыка
        # Приоритет 1: pythonw.exe (Embedded, без консоли)
        # Приоритет 2: VBS скрипт (если pythonw нет)
        # Приоритет 3: python.exe (С консолью)
        
        target_path = ""
        arguments = ""

        if embedded_pythonw.exists():
            # Используем pythonw.exe - нативное скрытие консоли
            target_path = str(embedded_pythonw)
            arguments = f'"{str(gui_path)}"'
        elif embedded_python.exists():
            # Fallback к VBS, если pythonw нет
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

        # Создаём ярлык через PowerShell
        ps_script = f'''
$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("{shortcut_path}")
$Shortcut.TargetPath = "{target_path}"
$Shortcut.Arguments = '{arguments}'
$Shortcut.WorkingDirectory = "{self.install_dir}"
$Shortcut.IconLocation = "{self.install_dir / 'assets' / 'ScriptM.ico'}"
$Shortcut.Description = "{self.APP_NAME}"
$Shortcut.Save()
'''
        subprocess.run(["powershell", "-Command", ps_script], check=True, capture_output=True)

        # Также создаём обычный start.bat для ручного запуска (с консолью для отладки)
        bat_path = self.install_dir / "start.bat"
        with open(bat_path, 'w') as f:
            f.write('@echo off\n')
            f.write(f'cd /d "{self.install_dir}"\n')
            if embedded_python.exists():
                f.write('python_embedded\\python.exe src\\gui.py\n')
            else:
                f.write('python src\\gui.py\n')
            f.write('pause\n')

    def register_in_registry(self):
        """Регистрация в реестре Windows"""
        try:
            app_key = winreg.CreateKey(
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ScriptManager"
            )

            winreg.SetValueEx(app_key, "DisplayName", 0, winreg.REG_SZ, self.APP_NAME)
            winreg.SetValueEx(app_key, "DisplayVersion", 0, winreg.REG_SZ, self.APP_VERSION)
            winreg.SetValueEx(app_key, "Publisher", 0, winreg.REG_SZ, self.APP_PUBLISHER)
            winreg.SetValueEx(app_key, "InstallLocation", 0, winreg.REG_SZ, str(self.install_dir))
            winreg.SetValueEx(app_key, "UninstallString", 0, winreg.REG_SZ, f'"{str(self.install_dir / "uninstall.bat")}"')
            winreg.SetValueEx(app_key, "DisplayIcon", 0, winreg.REG_SZ, str(self.install_dir / "assets" / "ScriptM.ico"))

            winreg.CloseKey(app_key)
        except Exception as e:
            print(f"      Предупреждение: не удалось зарегистрировать в реестре: {e}")

    def create_uninstaller(self):
        """Создание деинсталлятора"""
        uninstall_bat = self.install_dir / "uninstall.bat"
        with open(uninstall_bat, 'w', encoding='cp866') as f:
            f.write('@echo off\n')
            f.write('echo ================================================\n')
            f.write('echo   ScriptManager - Udalyenie (Uninstall)\n')
            f.write('echo ================================================\n')
            f.write('echo.\n')
            f.write('set /p CONFIRM=Prodolzhitch udalenie? (Y/n): \n')
            f.write('if /i "%CONFIRM%"=="n" (\n')
            f.write('    echo Otmena.\n')
            f.write('    pause\n')
            f.write('    exit /b 0\n')
            f.write(')\n\n')

            # Копируем себя во временную папку и запускаем оттуда
            f.write('echo.\n')
            f.write('echo Podgotovka...\n')
            f.write('set TEMP_UNINSTALL=%TEMP%\\uninstall_scriptmanager.bat\n')
            f.write('echo @echo off > "%TEMP_UNINSTALL%"\n')
            # f.write('echo chcp 65001 ^>nul >> "%TEMP_UNINSTALL%"\n')
            f.write('echo echo [1/4] Удаление файлов программы... >> "%TEMP_UNINSTALL%"\n')
            f.write(f'echo rd /s /q "{self.install_dir}" >> "%TEMP_UNINSTALL%"\n')
            f.write('echo echo      OK >> "%TEMP_UNINSTALL%"\n')
            f.write('echo echo [2/4] Удаление ярлыка... >> "%TEMP_UNINSTALL%"\n')
            f.write(f'echo del /f /q "{self.desktop}\\ScriptManager.lnk" >> "%TEMP_UNINSTALL%"\n')
            f.write('echo echo      OK >> "%TEMP_UNINSTALL%"\n')
            f.write('echo echo [3/4] Удаление из реестра... >> "%TEMP_UNINSTALL%"\n')
            f.write('echo reg delete "HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\ScriptManager" /f ^>nul 2^>^&1 >> "%TEMP_UNINSTALL%"\n')
            f.write('echo echo      OK >> "%TEMP_UNINSTALL%"\n')
            f.write('echo echo [4/4] Очистка завершена >> "%TEMP_UNINSTALL%"\n')
            f.write('echo echo      OK >> "%TEMP_UNINSTALL%"\n')
            f.write('echo echo. >> "%TEMP_UNINSTALL%"\n')
            f.write('echo echo ================================================ >> "%TEMP_UNINSTALL%"\n')
            f.write('echo echo   ScriptManager успешно удалён! >> "%TEMP_UNINSTALL%"\n')
            f.write('echo echo ================================================ >> "%TEMP_UNINSTALL%"\n')
            f.write('echo echo. >> "%TEMP_UNINSTALL%"\n')
            f.write('echo pause >> "%TEMP_UNINSTALL%"\n')
            f.write('echo del "%%~f0" >> "%TEMP_UNINSTALL%"\n')
            f.write('\n')
            f.write('start "" "%TEMP_UNINSTALL%"\n')
            f.write('exit\n')


if __name__ == "__main__":
    installer = SimpleInstaller()
    success = installer.install()

    if success:
        input("\nНажмите Enter для завершения...")
    else:
        input("\nНажмите Enter для выхода...")

    sys.exit(0 if success else 1)
