"""
Скрипт для подготовки Embedded Python

Автоматически:
1. Скачивает Python 3.11.9 Embedded (64-bit)
2. Распаковывает в папку python/
3. Устанавливает pip
4. Устанавливает зависимости из requirements.txt
"""

import urllib.request
import zipfile
import subprocess
import sys
from pathlib import Path
import shutil

class EmbeddedPythonPreparer:
    """Подготовка Embedded Python для упаковки"""

    PYTHON_VERSION = "3.11.9"
    PYTHON_URL = "https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip"
    GET_PIP_URL = "https://bootstrap.pypa.io/get-pip.py"

    def __init__(self):
        self.root_dir = Path(__file__).parent
        self.python_dir = self.root_dir / "python_embedded"
        self.zip_path = self.root_dir / "python_embedded.zip"

    def download_python(self):
        """Скачивание Embedded Python"""
        print(f"\nСкачивание Python {self.PYTHON_VERSION} Embedded...")

        if self.zip_path.exists():
            print("   Архив уже скачан, пропускаю...")
            return True

        try:
            urllib.request.urlretrieve(
                self.PYTHON_URL,
                self.zip_path,
                reporthook=self._download_progress
            )
            print("\n   Python скачан успешно!")
            return True
        except Exception as e:
            print(f"\n   Ошибка скачивания: {e}")
            return False

    def _download_progress(self, block_num, block_size, total_size):
        """Прогресс скачивания"""
        downloaded = block_num * block_size
        percent = min(downloaded * 100 / total_size, 100)
        print(f"\r   Прогресс: {percent:.1f}%", end='')

    def extract_python(self):
        """Распаковка Python"""
        print("\n\nРаспаковка Python...")

        if self.python_dir.exists():
            print("   Удаляю старую версию...")
            shutil.rmtree(self.python_dir)

        self.python_dir.mkdir(exist_ok=True)

        try:
            with zipfile.ZipFile(self.zip_path, 'r') as zip_ref:
                zip_ref.extractall(self.python_dir)
            print("   Python распакован!")
            return True
        except Exception as e:
            print(f"   Ошибка распаковки: {e}")
            return False

    def configure_python(self):
        """Настройка Python для использования pip"""
        print("\nНастройка Python...")

        # Редактируем python311._pth для включения site-packages
        pth_file = self.python_dir / "python311._pth"

        if pth_file.exists():
            with open(pth_file, 'r') as f:
                content = f.read()

            # Раскомментируем import site
            content = content.replace('#import site', 'import site')

            # Добавляем пути
            if 'Lib' not in content:
                content += '\nLib\n'
            if 'Lib\\site-packages' not in content:
                content += 'Lib\\site-packages\n'

            with open(pth_file, 'w') as f:
                f.write(content)

            print("   python311._pth настроен!")

        return True

    def install_pip(self):
        """Установка pip в Embedded Python"""
        print("\nУстановка pip...")

        python_exe = self.python_dir / "python.exe"
        get_pip_path = self.python_dir / "get-pip.py"

        # Скачиваем get-pip.py
        try:
            urllib.request.urlretrieve(self.GET_PIP_URL, get_pip_path)
            print("   get-pip.py скачан!")
        except Exception as e:
            print(f"   Ошибка скачивания get-pip.py: {e}")
            return False

        # Устанавливаем pip
        try:
            result = subprocess.run(
                [str(python_exe), str(get_pip_path)],
                capture_output=True,
                text=True,
                check=True
            )
            print("   pip установлен!")
            return True
        except subprocess.CalledProcessError as e:
            print(f"   Ошибка установки pip:\n{e.stderr}")
            return False

    def copy_tkinter(self):
        """Копирование tkinter и pythonw.exe из системного Python"""
        print("\nКопирование компонентов из системного Python...")

        import sys
        system_python_dir = Path(sys.executable).parent
        
        # Проверка версий
        sys_ver = sys.version_info
        if sys_ver.major != 3 or sys_ver.minor != 11:
            print(f"   [ВНИМАНИЕ] Версия системного Python ({sys_ver.major}.{sys_ver.minor}) отличается от Embedded (3.11)!")
            print("   Копирование DLL и pythonw.exe может привести к несовместимости.")
            print("   Рекомендуется запускать этот скрипт из-под Python 3.11.")
            # Не прерываем, но предупреждаем
            
        # 0. Попытка скопировать pythonw.exe (для скрытого запуска)
        pythonw_src = system_python_dir / "pythonw.exe"
        pythonw_dest = self.python_dir / "pythonw.exe"
        
        if pythonw_src.exists():
            shutil.copy2(pythonw_src, pythonw_dest)
            print("   Скопирован pythonw.exe (поддержка запуска без консоли)")
        else:
            print("   pythonw.exe не найден в системе (будет использоваться VBS wrapper)")

        # 1. Копируем модуль tkinter

        # 1. Копируем модуль tkinter
        tkinter_src = system_python_dir / "Lib" / "tkinter"
        tkinter_dest = self.python_dir / "Lib" / "tkinter"

        if tkinter_src.exists():
            if tkinter_dest.exists():
                shutil.rmtree(tkinter_dest)
            shutil.copytree(tkinter_src, tkinter_dest)
            print("   Скопирован модуль tkinter")

        # 2. Копируем DLLs
        tkinter_dll = system_python_dir / "DLLs" / "tcl86t.dll"
        tk_dll = system_python_dir / "DLLs" / "tk86t.dll"
        _tkinter_pyd = system_python_dir / "DLLs" / "_tkinter.pyd"

        dest_dir = self.python_dir / "DLLs"
        dest_dir.mkdir(exist_ok=True)

        files_to_copy = [
            (tkinter_dll, "tcl86t.dll"),
            (tk_dll, "tk86t.dll"),
            (_tkinter_pyd, "_tkinter.pyd")
        ]

        for src_file, filename in files_to_copy:
            if src_file.exists():
                shutil.copy2(src_file, dest_dir / filename)
                # Также копируем все DLL в корень для правильной работы
                shutil.copy2(src_file, self.python_dir / filename)
                print(f"   Скопирован: {filename}")

        # 3. Копируем библиотеки tcl/tk
        tcl_dir = system_python_dir / "tcl"
        if tcl_dir.exists():
            dest_tcl = self.python_dir / "tcl"
            if dest_tcl.exists():
                shutil.rmtree(dest_tcl)
            shutil.copytree(tcl_dir, dest_tcl)
            print("   Скопирована библиотека tcl")

        return True

    def install_dependencies(self):
        """Установка зависимостей из requirements.txt"""
        print("\nУстановка зависимостей...")

        python_exe = self.python_dir / "python.exe"
        requirements = self.root_dir / "requirements.txt"

        if not requirements.exists():
            print("   requirements.txt не найден!")
            return False

        try:
            # Сначала устанавливаем совместимые версии numpy и pandas
            subprocess.run(
                [str(python_exe), "-m", "pip", "install", "numpy==1.24.3", "pandas==2.0.3"],
                capture_output=True,
                text=True,
                check=True
            )

            result = subprocess.run(
                [str(python_exe), "-m", "pip", "install", "-r", str(requirements)],
                capture_output=True,
                text=True,
                check=True
            )
            print("   Зависимости установлены!")
            print(f"\n   Установленные пакеты:")

            # Показываем список установленных пакетов
            result = subprocess.run(
                [str(python_exe), "-m", "pip", "list"],
                capture_output=True,
                text=True,
                check=True
            )

            for line in result.stdout.split('\n')[2:]:  # Пропускаем заголовок
                if line.strip():
                    print(f"      {line}")

            return True
        except subprocess.CalledProcessError as e:
            print(f"   Ошибка установки зависимостей:\n{e.stderr}")
            return False

    def cleanup(self):
        """Очистка временных файлов"""
        print("\nОчистка временных файлов...")

        if self.zip_path.exists():
            self.zip_path.unlink()
            print("   Архив удалён")

        get_pip = self.python_dir / "get-pip.py"
        if get_pip.exists():
            get_pip.unlink()

        return True

    def verify_installation(self):
        """Проверка установки"""
        print("\nПроверка установки...")

        python_exe = self.python_dir / "python.exe"

        # Проверяем версию Python
        result = subprocess.run(
            [str(python_exe), "--version"],
            capture_output=True,
            text=True
        )
        print(f"   Python: {result.stdout.strip()}")

        # Проверяем pip
        result = subprocess.run(
            [str(python_exe), "-m", "pip", "--version"],
            capture_output=True,
            text=True
        )
        print(f"   pip: {result.stdout.strip()}")

        # Проверяем импорт ключевых модулей
        test_imports = ['tkinter', 'customtkinter', 'pandas', 'openpyxl', 'requests']

        print("\n   Проверка модулей:")
        all_ok = True
        for module in test_imports:
            result = subprocess.run(
                [str(python_exe), "-c", f"import {module}; print('{module}: OK')"],
                capture_output=True,
                text=True
            )
            if result.returncode == 0:
                print(f"      OK {module}")
            else:
                print(f"      FAIL {module}")
                all_ok = False

        return all_ok

    def prepare_all(self):
        """Полная подготовка Embedded Python"""
        print("="*50)
        print("  Подготовка Embedded Python для упаковки")
        print("="*50)

        steps = [
            ("Скачивание Python", self.download_python),
            ("Распаковка Python", self.extract_python),
            ("Настройка Python", self.configure_python),
            ("Установка pip", self.install_pip),
            ("Копирование tkinter", self.copy_tkinter),
            ("Установка зависимостей", self.install_dependencies),
            ("Проверка установки", self.verify_installation),
            ("Очистка", self.cleanup),
        ]

        for step_name, step_func in steps:
            if not step_func():
                print(f"\nОшибка на этапе: {step_name}")
                return False

        print("\n" + "="*50)
        print("EMBEDDED PYTHON ГОТОВ К УПАКОВКЕ!")
        print("="*50)
        print(f"\nПапка: {self.python_dir}")
        print(f"Размер: {self._get_dir_size(self.python_dir) / 1024 / 1024:.1f} MB")
        print("\nСледующий шаг:")
        print("   python build.py")

        return True

    def _get_dir_size(self, path):
        """Получение размера директории"""
        import os
        total = 0
        for entry in os.scandir(path):
            if entry.is_file():
                total += entry.stat().st_size
            elif entry.is_dir():
                total += self._get_dir_size(entry.path)
        return total


if __name__ == "__main__":
    preparer = EmbeddedPythonPreparer()
    success = preparer.prepare_all()
    sys.exit(0 if success else 1)
