# 🚀 Быстрая инструкция по сборке релиза

## Вариант 2: Автономный (рекомендуется) ⭐

### Требования для сборки:
- Windows 10/11
- Python 3.11.9+
- Интернет (для скачивания Embedded Python)

### Шаг 1: Подготовка Embedded Python

```bash
python prepare_embedded_python.py
```

**Что происходит:**
1. Скачивается Python 3.11.9 Embedded (~30 MB)
2. Распаковывается в `python_embedded/`
3. Устанавливается pip
4. Устанавливаются все зависимости из `requirements.txt`
5. Проверяется работоспособность

**Время:** 3-5 минут
**Результат:** Папка `python_embedded/` (~120 MB)

### Шаг 2: Сборка release пакета

```bash
python build.py
```

**Что происходит:**
1. Очищаются старые сборки
2. Собирается `ScriptManager_Setup.exe` через PyInstaller
3. Копируются все файлы в `release/`:
   - `ScriptManager_Setup.exe` - инсталлятор
   - `python_embedded/` - встроенный Python
   - `src/` - исходники
   - `config/` - конфигурация
   - `doc/` - документация
   - `README.md`, `QUICKSTART.md`, `requirements.txt`
4. Создаётся `УСТАНОВКА.txt` с инструкцией

**Время:** 2-3 минуты
**Результат:** Папка `release/` (~150 MB)

### Шаг 3: Распространение

```bash
# Вариант А: ZIP архив
cd release
7z a ScriptManager_v1.0.0_Standalone.zip *

# Вариант Б: Скопировать на флешку/сетевой диск
xcopy release "Z:\ScriptManager\" /E /I
```

### Шаг 4: Установка на целевом компьютере

1. Скопировать папку `release/` на целевой компьютер
2. Запустить `ScriptManager_Setup.exe` от имени администратора
3. Выбрать путь установки (по умолчанию: `C:\Program Files\ScriptManager`)
4. Нажать "Установить"
5. Дождаться завершения (2-3 минуты)
6. Запустить с ярлыка на рабочем столе

**Python НЕ требуется!** Всё уже встроено.

---

## Вариант 1: С системным Python (компактный)

### Если Python уже установлен на целевых компьютерах:

```bash
python build.py
```

**Результат:** Папка `release/` (~50 MB) без `python_embedded/`

**На целевом компьютере потребуется:**
1. Установить Python 3.11.9+
2. Запустить `pip install -r requirements.txt`

---

## Проверка сборки

### После `prepare_embedded_python.py`:

```bash
# Проверить, что Python работает
python_embedded\python.exe --version

# Проверить установленные пакеты
python_embedded\python.exe -m pip list

# Проверить импорт модулей
python_embedded\python.exe -c "import customtkinter, pandas, openpyxl, requests; print('OK')"
```

### После `build.py`:

```bash
# Проверить размер release
dir release

# Проверить содержимое
tree release /F

# Проверить инсталлятор
release\ScriptManager_Setup.exe
```

---

## Структура финального пакета

```
release/
│
├── ScriptManager_Setup.exe           # Инсталлятор (запустить!)
│
├── python_embedded/                  # Встроенный Python
│   ├── python.exe                    # ~30 MB
│   ├── python311.dll
│   ├── Lib/
│   │   └── site-packages/            # Все зависимости
│   │       ├── customtkinter/
│   │       ├── pandas/
│   │       ├── openpyxl/
│   │       ├── requests/
│   │       └── ...
│   └── ...
│
├── src/                              # Исходники
│   ├── gui.py
│   ├── config_manager.py
│   ├── scheduler.py
│   └── scripts/
│
├── config/                           # Конфигурация
│   ├── scripts.json
│   ├── email.json.example
│   ├── telegram.json.example
│   └── calendar.json
│
├── doc/                              # Документация
│   ├── packaging_guide.md
│   ├── calendar_auto_update.md
│   └── ...
│
├── README.md
├── QUICKSTART.md
├── PROJECT_SUMMARY.md
├── requirements.txt
└── УСТАНОВКА.txt                     # Инструкция для пользователя
```

**Размер:** ~150 MB (Вариант 2) или ~50 MB (Вариант 1)

---

## Устранение проблем

### Ошибка: "ModuleNotFoundError: No module named 'pip'"

**Решение:**
```bash
# Удалить python_embedded/
rmdir /s python_embedded

# Запустить prepare_embedded_python.py заново
python prepare_embedded_python.py
```

### Ошибка: "PyInstaller not found"

**Решение:**
```bash
pip install pyinstaller
```

### Ошибка: "python_embedded/ не найден" при сборке

**Решение:**
```bash
# Сначала запустить prepare_embedded_python.py
python prepare_embedded_python.py

# Потом build.py
python build.py
```

### Инсталлятор не запускается

**Решение:**
1. Запускать от имени администратора
2. Проверить антивирус (может блокировать .exe)
3. Проверить логи сборки в `build/`

---

## Checklist перед распространением

- [ ] `python prepare_embedded_python.py` выполнен успешно
- [ ] `python_embedded/` создана (~120 MB)
- [ ] `python build.py` выполнен успешно
- [ ] `release/ScriptManager_Setup.exe` существует
- [ ] `release/python_embedded/` скопирована
- [ ] Тестовая установка на чистой Windows 10/11
- [ ] Приложение запускается с ярлыка
- [ ] Все скрипты настраиваются через GUI
- [ ] Тестовая отправка работает
- [ ] Деинсталлятор работает

---

## Дополнительная информация

**Полная документация:** [doc/packaging_guide.md](doc/packaging_guide.md)

**Быстрый старт для пользователей:** [QUICKSTART.md](QUICKSTART.md)

**Техническая информация:** [README.md](README.md)

---

**Готово к сборке и распространению! 📦**
