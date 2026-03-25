# 📦 Реализация автономной установки с Embedded Python

## 🎯 Цель

Создать полностью автономный инсталлятор ScriptManager, который:
- ✅ **НЕ требует установки Python** на целевом компьютере
- ✅ Все зависимости уже встроены
- ✅ Один клик - и приложение работает
- ✅ Простая установка и удаление через Windows

---

## 📋 Архитектура решения

### Компоненты системы:

```
┌─────────────────────────────────────────────────────────┐
│  prepare_embedded_python.py                             │
│  • Скачивает Python 3.11.9 Embedded                     │
│  • Устанавливает pip                                    │
│  • Устанавливает все зависимости                        │
└─────────────────────────────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────┐
│  python_embedded/                                        │
│  ├── python.exe                                          │
│  ├── python311.dll                                       │
│  └── Lib/site-packages/                                 │
│      ├── customtkinter/                                  │
│      ├── pandas/                                         │
│      ├── openpyxl/                                       │
│      └── ...                                             │
└─────────────────────────────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────┐
│  build.py                                                │
│  • Собирает ScriptManager_Setup.exe                     │
│  • Копирует python_embedded/ в release/                 │
│  • Создаёт финальный пакет                              │
└─────────────────────────────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────┐
│  release/                                                │
│  ├── ScriptManager_Setup.exe                            │
│  ├── python_embedded/                                    │
│  ├── src/                                                │
│  └── config/                                             │
└─────────────────────────────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────┐
│  installer.py                                            │
│  • Копирует все файлы в C:\Program Files\               │
│  • Создаёт start.bat с путём к python_embedded/         │
│  • Регистрирует в Windows                               │
└─────────────────────────────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────┐
│  C:\Program Files\ScriptManager\                        │
│  ├── python_embedded/python.exe                         │
│  ├── start.bat → запускает python_embedded/python.exe  │
│  └── src/gui.py                                          │
└─────────────────────────────────────────────────────────┘
```

---

## 🔧 Реализованные компоненты

### 1. prepare_embedded_python.py

**Назначение:** Автоматическая подготовка Embedded Python со всеми зависимостями

**Функционал:**
```python
class EmbeddedPythonPreparer:
    def download_python(self):
        """Скачивание Python 3.11.9 Embedded с python.org"""
        URL = "https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip"

    def extract_python(self):
        """Распаковка в python_embedded/"""

    def configure_python(self):
        """Настройка python311._pth для включения site-packages"""
        # Раскомментируем: import site
        # Добавляем: Lib\site-packages

    def install_pip(self):
        """Установка pip в Embedded Python"""
        # Скачиваем get-pip.py
        # Запускаем: python.exe get-pip.py

    def install_dependencies(self):
        """Установка всех зависимостей из requirements.txt"""
        # python.exe -m pip install -r requirements.txt

    def verify_installation(self):
        """Проверка работоспособности"""
        # Тестируем импорт всех модулей
```

**Результат:**
- Папка `python_embedded/` (~120 MB)
- Все зависимости установлены
- Готово к упаковке

---

### 2. build.py (модифицирован)

**Изменения:**

```python
def create_release_package(self):
    """Создание финального пакета"""

    # НОВОЕ: Копирование Embedded Python
    python_embedded_src = self.root_dir / "python_embedded"
    python_embedded_dest = self.release_dir / "python_embedded"

    if python_embedded_src.exists():
        shutil.copytree(python_embedded_src, python_embedded_dest)
        print(f"📦 Embedded Python скопирован")
    else:
        print(f"⚠️ Embedded Python не найден")
        print(f"💡 Запустите: python prepare_embedded_python.py")

    # НОВОЕ: Разные инструкции для разных вариантов
    has_embedded_python = (self.root_dir / "python_embedded").exists()

    if has_embedded_python:
        # Создаём README для автономной установки
        f.write("✅ Python НЕ требуется! Всё встроено!")
    else:
        # Создаём README с требованием Python
        f.write("⚠️ Требуется Python 3.11.9+")
```

---

### 3. installer.py (модифицирован)

**Изменения:**

```python
def copy_files(self):
    """Копирование файлов приложения"""

    # Копируем основные директории
    dirs_to_copy = ['src', 'config', 'doc']

    # НОВОЕ: Копирование Embedded Python
    python_embedded_source = self.source_dir / 'python_embedded'
    python_embedded_dest = self.install_dir / 'python_embedded'

    if python_embedded_source.exists():
        shutil.copytree(python_embedded_source, python_embedded_dest)

def create_desktop_shortcut(self):
    """Создание ярлыка на рабочем столе"""

    bat_path = self.install_dir / "start.bat"

    # НОВОЕ: Определение типа установки
    embedded_python = self.install_dir / "python_embedded" / "python.exe"

    with open(bat_path, 'w') as f:
        if embedded_python.exists():
            # Используем встроенный Python
            f.write('python_embedded\\python.exe src\\gui.py\n')
        else:
            # Используем системный Python
            f.write('python src\\gui.py\n')
```

**Логика работы:**
1. Проверяет наличие `python_embedded/` в установочном пакете
2. Если есть - копирует и использует встроенный Python
3. Если нет - использует системный Python

---

## 📦 Структура файлов

### До установки (release пакет):

```
release/
├── ScriptManager_Setup.exe           # 10 MB (инсталлятор)
├── python_embedded/                  # 120 MB (встроенный Python)
│   ├── python.exe
│   ├── python311.dll
│   ├── Lib/
│   │   └── site-packages/
│   │       ├── customtkinter/
│   │       ├── pandas/
│   │       ├── openpyxl/
│   │       ├── requests/
│   │       └── win32com/
│   └── python311._pth               # Настроен для site-packages
├── src/                              # 50 KB (исходники)
├── config/                           # 10 KB (конфигурация)
└── doc/                              # 100 KB (документация)

ИТОГО: ~150 MB
```

### После установки (на целевом компьютере):

```
C:\Program Files\ScriptManager\
├── python_embedded/                  # Встроенный Python
│   └── ...
├── src/                              # Исходники
│   ├── gui.py
│   └── scripts/
├── config/                           # Конфигурация
├── logs/                             # Логи (создаётся при первом запуске)
├── start.bat                         # Запускает python_embedded/python.exe
└── uninstall.bat                     # Деинсталлятор
```

---

## 🚀 Процесс сборки

### Этап 1: Подготовка Embedded Python

```bash
python prepare_embedded_python.py
```

**Шаги:**
1. Скачивается `python-3.11.9-embed-amd64.zip` (~30 MB)
2. Распаковывается в `python_embedded/`
3. Редактируется `python311._pth`:
   ```
   python311.zip
   .

   # Раскомментировано:
   import site

   # Добавлено:
   Lib\site-packages
   ```
4. Скачивается `get-pip.py`
5. Запускается `python.exe get-pip.py`
6. Устанавливаются зависимости:
   ```bash
   python.exe -m pip install -r requirements.txt
   ```
7. Проверяется импорт всех модулей

**Время:** 3-5 минут
**Результат:** Папка `python_embedded/` (~120 MB)

---

### Этап 2: Сборка release пакета

```bash
python build.py
```

**Шаги:**
1. Очистка старых сборок:
   ```python
   shutil.rmtree('dist/')
   shutil.rmtree('build/')
   shutil.rmtree('release/')
   ```

2. Сборка инсталлятора:
   ```bash
   pyinstaller --onefile --windowed --name=ScriptManager_Setup installer.py
   ```

3. Создание release пакета:
   ```python
   release/
   ├── ScriptManager_Setup.exe  ← Из dist/
   ├── python_embedded/         ← Копируется из корня
   ├── src/                     ← Копируется из корня
   ├── config/                  ← Копируется из корня
   └── ...
   ```

4. Создание `УСТАНОВКА.txt`:
   - Если `python_embedded/` есть → "Python НЕ требуется!"
   - Если `python_embedded/` нет → "Требуется Python 3.11.9+"

**Время:** 2-3 минуты
**Результат:** Папка `release/` (~150 MB)

---

### Этап 3: Распространение

**Вариант А: ZIP архив**
```bash
cd release
7z a ScriptManager_v1.0.0_Standalone.zip *
```

**Вариант Б: Копирование на флешку/сетевой диск**
```bash
xcopy release "D:\ScriptManager" /E /I
```

---

## 🖥️ Процесс установки на целевом компьютере

### Шаг 1: Запуск инсталлятора

```
ScriptManager_Setup.exe
```

**Требования:**
- ✅ Права администратора
- ❌ Python НЕ требуется!

---

### Шаг 2: Копирование файлов

Инсталлятор копирует:
```python
C:\Program Files\ScriptManager\
├── python_embedded/        ← 120 MB
├── src/                    ← 50 KB
├── config/                 ← 10 KB
├── doc/                    ← 100 KB
├── requirements.txt
└── ...
```

---

### Шаг 3: Создание start.bat

```batch
@echo off
cd /d "C:\Program Files\ScriptManager"
python_embedded\python.exe src\gui.py
```

**Важно:** Используется `python_embedded\python.exe`, а НЕ системный `python`!

---

### Шаг 4: Создание ярлыка

Ярлык на рабочем столе → запускает `start.bat`

---

### Шаг 5: Регистрация в Windows

```python
winreg.CreateKey(HKEY_LOCAL_MACHINE,
    r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ScriptManager")
```

Приложение появляется в "Программы и компоненты"

---

## ✅ Преимущества решения

### 1. Полная автономность
- ❌ Не нужно устанавливать Python
- ❌ Не нужно настраивать PATH
- ❌ Не нужно устанавливать зависимости
- ✅ Всё уже встроено!

### 2. Изолированность
- Не конфликтует с системным Python
- Не влияет на другие Python-приложения
- Собственная версия всех библиотек

### 3. Простота использования
- Один файл для запуска установки
- Один клик - приложение установлено
- Работает "из коробки"

### 4. Предсказуемость
- Фиксированная версия Python (3.11.9)
- Фиксированные версии зависимостей
- Одинаковая работа на всех компьютерах

---

## ⚠️ Ограничения и компромиссы

### 1. Размер дистрибутива

**Вариант 1 (системный Python):** ~50 MB
**Вариант 2 (embedded Python):** ~150 MB

**Компромисс:** +100 MB за полную автономность

---

### 2. Обновление Python

**Вариант 1:** Обновляется отдельно от приложения
**Вариант 2:** Встроен в приложение, обновляется вместе с ним

**Компромисс:** Меньше гибкости, но проще поддержка

---

### 3. Время сборки

**Вариант 1:** 2-3 минуты
**Вариант 2:** 5-8 минут (+ скачивание и настройка Python)

**Компромисс:** +5 минут один раз при сборке

---

## 🎯 Рекомендации по использованию

### Используйте Вариант 2 (Embedded Python), если:

✅ На целевых компьютерах НЕТ Python
✅ Пользователи не технические специалисты
✅ Нужна максимально простая установка
✅ Важна стабильность и предсказуемость
✅ Есть достаточно места (~200 MB)

**РЕКОМЕНДУЕТСЯ для большинства случаев!**

---

### Используйте Вариант 1 (системный Python), если:

✅ Python уже установлен на всех компьютерах
✅ Есть централизованное управление ПО
✅ Критично минимизировать размер дистрибутива
✅ Пользователи знакомы с Python

---

## 📊 Сравнительная таблица

| Параметр | Вариант 1 | Вариант 2 |
|----------|-----------|-----------|
| Размер дистрибутива | ~50 MB | ~150 MB |
| Требует Python | ✅ Да | ❌ Нет |
| Требует pip install | ✅ Да | ❌ Нет |
| Время установки | 5-10 мин | 2-3 мин |
| Сложность установки | Средняя | Низкая |
| Конфликты версий | Возможны | Невозможны |
| Время сборки | 2-3 мин | 5-8 мин |
| Обновление Python | Независимо | Вместе с приложением |
| **Рекомендация** | 🟡 Если Python есть | 🟢 **Рекомендуется** |

---

## 🔍 Технические детали

### Настройка python311._pth

**До:**
```
python311.zip
.

# Uncomment to run site.main() automatically
#import site
```

**После:**
```
python311.zip
.

# Uncomment to run site.main() automatically
import site

# Custom paths
Lib\site-packages
```

**Зачем:** Включить поддержку site-packages для pip-пакетов

---

### Проверка работоспособности

```bash
# Проверка версии
python_embedded\python.exe --version
# Python 3.11.9

# Проверка pip
python_embedded\python.exe -m pip --version
# pip 24.0

# Проверка модулей
python_embedded\python.exe -c "import customtkinter, pandas, openpyxl, requests; print('OK')"
# OK
```

---

## 🐛 Устранение проблем

### Проблема: "No module named 'pip'"

**Причина:** pip не установлен в Embedded Python

**Решение:**
```bash
rmdir /s python_embedded
python prepare_embedded_python.py
```

---

### Проблема: "ImportError: DLL load failed"

**Причина:** Отсутствуют системные библиотеки

**Решение:**
Установить Visual C++ Redistributable:
https://aka.ms/vs/17/release/vc_redist.x64.exe

---

### Проблема: "ModuleNotFoundError: No module named 'tkinter'"

**Причина:** Embedded Python не включает tkinter по умолчанию

**Решение:**
tkinter включён в стандартную версию Embedded Python 3.11.9, проверить версию

---

## 📝 Заключение

Реализация автономной установки с Embedded Python обеспечивает:

✅ **Простоту:** Один клик - приложение работает
✅ **Надёжность:** Нет зависимостей от системы
✅ **Предсказуемость:** Одинаковая работа везде
✅ **Изолированность:** Не конфликтует с другими приложениями

**Компромисс:** +100 MB размера за полную автономность

**Вывод:** Рекомендуется для большинства сценариев использования!

---

**Готово к использованию! 🚀**
