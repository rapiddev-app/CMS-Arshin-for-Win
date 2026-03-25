"""
Скрипт сборки проекта ScriptManager (build.py)

Выполняет следующие задачи:
1. Собирает GUI приложение и инсталлятор в .exe с помощью PyInstaller (если требуется)
2. Подготавливает папку release с полной структурой для распространения
3. Включает Embedded Python (если подготовлен) для автономной работы
"""

import subprocess
import shutil
import os
from pathlib import Path

class Builder:
    """Класс для сборки дистрибутива ScriptManager"""

    def __init__(self):
        self.root_dir = Path(__file__).parent
        self.dist_dir = self.root_dir / "dist"
        self.build_dir = self.root_dir / "build"
        self.release_dir = self.root_dir / "release"
        self.python_embedded_dir = self.root_dir / "python_embedded"

    def clean(self):
        """Очистка временных директорий перед сборкой"""
        print("Очистка временных файлов...")
        dirs_to_clean = [self.dist_dir, self.build_dir, self.release_dir]
        for dir_path in dirs_to_clean:
            if dir_path.exists():
                try:
                    shutil.rmtree(dir_path)
                    print(f"   Удалено: {dir_path}")
                except Exception as e:
                    print(f"   Ошибка удаления {dir_path}: {e}")

    def build_installer(self):
        """Сборка инсталлятора в .exe"""
        print("\nСборка инсталлятора...")
        
        # Используем installer_simple.py как более надежный вариант
        installer_script = self.root_dir / "installer_simple.py"
        if not installer_script.exists():
             installer_script = self.root_dir / "installer.py"

        print(f"   Используем скрипт: {installer_script.name}")

        cmd = [
            "pyinstaller",
            "--onefile",
            "--windowed",
            "--name=ScriptManager_Setup",
            "--icon=NONE",
            "--clean",
            str(self.root_dir / "installer.py") 
        ]
        
        # Примечание: installer.py - это GUI на tkinter, ему нужен --windowed.
        # installer_simple.py - консольный, ему нужен --console.
        # В текущей реализации build.py собирался installer.py. Оставим так.

        try:
            subprocess.run(cmd, check=True, capture_output=False) # capture_output=False чтобы видеть прогресс
            print("   Инсталлятор собран успешно!")
            return True
        except subprocess.CalledProcessError as e:
            print(f"   Ошибка сборки инсталлятора!")
            return False

    def get_dir_size(self, path):
        """Рекурсивный подсчет размера директории"""
        total = 0
        try:
            for entry in os.scandir(path):
                if entry.is_file():
                    total += entry.stat().st_size
                elif entry.is_dir():
                    total += self.get_dir_size(entry.path)
        except OSError:
            pass
        return total

    def create_release_package(self):
        """Создание финальной папки release"""
        print("\nФормирование release пакета...")

        self.release_dir.mkdir(exist_ok=True)

        # 1. Копируем собранный инсталлятор
        installer_src = self.dist_dir / "ScriptManager_Setup.exe"
        installer_dest = self.release_dir / "ScriptManager_Setup.exe"

        if installer_src.exists():
            shutil.copy2(installer_src, installer_dest)
            print(f"   Инсталлятор скопирован: {installer_dest.name}")
        else:
            print("   ОШИБКА: Инсталлятор не найден в dist/")
            return False

        # 2. Копируем Embedded Python (если есть)
        python_dest = self.release_dir / "python_embedded"
        
        if self.python_embedded_dir.exists():
            if python_dest.exists():
                shutil.rmtree(python_dest)
            shutil.copytree(self.python_embedded_dir, python_dest)
            size_mb = self.get_dir_size(python_dest) / 1024 / 1024
            print(f"   Embedded Python скопирован ({size_mb:.1f} MB)")
        else:
            print("   ВНИМАНИЕ: Embedded Python не найден. Сборка будет требовать установленного Python у пользователя.")

        # 3. Копируем файлы проекта (исходники, конфиги, доки)
        # Это нужно для того, чтобы инсталлятор мог их взять и установить
        files_to_copy = [
            'src',
            'config',
            'doc',
            'requirements.txt',
            'README.md',
            'QUICKSTART.md',
            'PROJECT_SUMMARY.md',
            # Скрипты инсталлятора тоже кладем, вдруг пригодятся для отладки
            'installer.py', 
            'installer_simple.py',
            'install.bat',
            'assets' 
        ]

        for item in files_to_copy:
            src = self.root_dir / item
            dest = self.release_dir / item
            
            if not src.exists():
                continue

            if src.is_dir():
                # Игнорируем __pycache__
                shutil.copytree(src, dest, dirs_exist_ok=True, ignore=shutil.ignore_patterns('__pycache__', '*.pyc'))
            else:
                shutil.copy2(src, dest)
            
            print(f"   Скопировано: {item}")

        # 4. Создаем README для релиза
        self._create_release_readme()
        
        print(f"\nСБОРКА ЗАВЕРШЕНА УСПЕШНО!")
        print(f"Папка релиза: {self.release_dir}")
        print(f"Размер: {self.get_dir_size(self.release_dir) / 1024 / 1024:.2f} MB")
        
        return True

    def _create_release_readme(self):
        readme_path = self.release_dir / "УСТАНОВКА.txt"
        has_embedded = (self.release_dir / "python_embedded").exists()
        
        text = """
==============================================================================
                   ScriptManager - Инструкция по установке
==============================================================================

1. Запустите файл ScriptManager_Setup.exe
   (Требуются права администратора для установки в Program Files)

2. Следуйте инструкциям установщика.

"""
        if has_embedded:
            text += """
[ВАЖНО] 
В эту версию УЖЕ ВСТРОЕН Python. Ничего дополнительно скачивать не нужно.
Программа полностью автономна.
"""
        else:
            text += """
[ВАЖНО]
Для работы программы требуется установленный Python 3.10 или новее.
Если он не установлен, инсталлятор не сможет запустить программу.
"""

        text += """
3. После установки на рабочем столе появится ярлык ScriptManager.

------------------------------------------------------------------------------
Настройка после установки:
------------------------------------------------------------------------------
1. При первом запуске нажмите "Настроить" для любого скрипта.
2. Введите пароль администратора (по умолчанию: cms2025).
3. Настройте Email и Telegram уведомления.

==============================================================================
"""
        try:
            with open(readme_path, 'w', encoding='utf-8') as f:
                f.write(text)
            print("   Создан файл УСТАНОВКА.txt")
        except Exception as e:
            print(f"   Ошибка создания readme: {e}")

    def run(self):
        self.clean()
        if self.build_installer():
            self.create_release_package()

if __name__ == "__main__":
    builder = Builder()
    builder.run()
