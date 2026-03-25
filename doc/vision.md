# Техническое видение проекта: Менеджер автоматизации скриптов для Windows

## Технологии

### Язык программирования
**Python 3.10+**
- Простота разработки и поддержки
- Богатая экосистема библиотек
- Легкая упаковка в .exe через PyInstaller
- Встроенная поддержка работы с JSON, файлами, процессами

### GUI
**CustomTkinter**
- Современный внешний вид поверх стандартного tkinter
- Простота использования
- Встроенные виджеты (кнопки, переключатели, поля ввода)
- Не требует сложных зависимостей
- Достаточно для создания понятного интерфейса для непродвинутых пользователей

### Запуск скриптов
**subprocess (встроенный модуль Python)**
- Универсальность - поддержка запуска .py, .bat, .ps1, .exe файлов
- Скрипты запускаются как внешние процессы без модификации их кода
- Перехват stdout/stderr для логирования результатов выполнения
- Изоляция - ошибка в скрипте не роняет приложение

### Планирование задач
**Windows Task Scheduler** (через библиотеки `pythoncom` и `win32com.client`)
- Надежность и стабильность встроенного планировщика Windows
- Работа независимо от GUI-приложения (скрипты выполняются даже когда приложение закрыто)
- Встроенная в ОС - не требует дополнительных служб
- Возможность включения/отключения задач через API (enable/disable)
- Поддержка сложных расписаний

### Взаимодействие со скриптами
**Общий конфигурационный файл (JSON)**
- Приложение создает и обновляет JSON-конфиг с настройками
- Скрипты читают из него:
  - Пути к исходным файлам
  - Пути для результирующих файлов
  - Настройки email (отправитель, получатель, тема, тело)
  - Параметры SMTP-сервера
- Скрипты самостоятельно обрабатывают данные и отправляют письма
- Приложение только управляет расписанием запуска

### Отправка тестовых писем
**smtplib (встроенная библиотека Python)**
- Для функции "Отправить тестовое письмо" из GUI
- Без дополнительных зависимостей
- Поддержка TLS/SSL
- Возможность прикрепления файлов

### Производственный календарь РФ
**Простой подход - JSON-файл или API**
- Вариант 1: JSON-файл с праздниками (обновляется вручную раз в год)
- Вариант 2: API isdayoff.ru (простой GET-запрос для проверки дня)
- Минимум зависимостей, максимум простоты

### Конфигурация
**JSON файлы**
- Простота чтения/записи через встроенный модуль `json`
- Читаемость для человека при необходимости ручного редактирования
- Хранение:
  - Настроек скриптов (расписание, пути, состояние enabled/disabled)
  - Глобальных настроек email
  - Пароля для доступа к настройкам (хешированного)

### Упаковка и распространение
**PyInstaller**
- Создание standalone .exe файла (один файл или папка с exe + библиотеки)
- Все зависимости включены в дистрибутив
- Без необходимости установки Python на машине пользователя
- Простое распространение через копирование файла

## Принципы разработки

### KISS (Keep It Simple, Stupid)
- Простые решения вместо сложных архитектур
- Прямолинейная логика без излишних абстракций
- Если можно решить в 10 строк - не пишем 100

### Никакого оверинжиниринга
- Не создаем архитектуру "на вырост"
- Не используем паттерны проектирования там, где они не нужны
- Не добавляем функциональность "на всякий случай"
- Решаем конкретную задачу конкретным способом

### Читаемость кода
- Понятные названия переменных и функций
- Простая структура - один модуль = одна зона ответственности
- Минимум комментариев (код должен быть самодокументируемым)
- Короткие функции с понятной логикой

### Минимум зависимостей
- Используем встроенные библиотеки Python где возможно
- Сторонние библиотеки только когда действительно необходимо
- Меньше зависимостей = проще поддержка и упаковка

### Fail-fast принцип
- Если что-то пошло не так - сразу показываем ошибку пользователю
- Никаких молчаливых падений
- Простые и понятные сообщения об ошибках на русском языке
- Валидация входных данных перед выполнением операций

### Работает > Идеально
- Сначала делаем работающее решение
- Потом (если нужно) улучшаем
- Избегаем перфекционизма
- Главное - польза для пользователя

## Структура проекта

### Организация файлов и папок

```
ScriptManager/
│
├── main.py              # Точка входа приложения
├── gui.py               # GUI-интерфейс (CustomTkinter)
├── scheduler.py         # Работа с Windows Task Scheduler
├── config.py            # Чтение/запись конфигурации
├── utils.py             # Вспомогательные функции (хеширование пароля, проверка календаря)
│
├── config/              # Конфигурационные файлы (создаются автоматически)
│   ├── settings.json    # Настройки приложения и скриптов
│   ├── email.json       # Настройки email (защищены паролем)
│   └── calendar.json    # Производственный календарь РФ
│
├── scripts/             # Рабочие скрипты (пользователь размещает здесь)
│   ├── script1/
│   │   └── run.py
│   ├── script2/
│   │   └── process.bat
│   └── script3/
│       └── main.ps1
│
├── logs/                # Логи выполнения (создаются автоматически)
│   └── app.log
│
└── requirements.txt     # Зависимости Python
```

### Принципы организации

**Модульность**
- Один модуль = одна зона ответственности
- main.py - точка входа, минимум логики
- gui.py - только интерфейс
- scheduler.py - только работа с планировщиком
- config.py - только работа с конфигами
- utils.py - вспомогательные функции

**Портативность**
- Все конфиги и логи рядом с приложением
- Можно скопировать всю папку на другой компьютер
- Не используем системные пути (кроме Task Scheduler)

**Автоматизация**
- Папки config/ и logs/ создаются автоматически при первом запуске
- Если конфигов нет - создаются с дефолтными значениями
- Скрипты сканируются автоматически из scripts/

**Соглашения**
- Имя папки скрипта = имя скрипта в интерфейсе
- В каждой папке скрипта должен быть исполняемый файл (любое расширение)
- Приложение ищет первый исполняемый файл в папке

## Архитектура

### Компоненты системы

**1. GUI (gui.py) - Графический интерфейс**
- Отображает список скриптов из папки scripts/
- Управление настройками каждого скрипта (расписание, пути, вкл/выкл)
- Защищенное окно настроек email (требует пароль)
- Кнопка "Отправить тестовое письмо"
- Простой и понятный интерфейс для непродвинутых пользователей

**2. Config Manager (config.py) - Управление конфигурацией**
- Чтение/запись JSON файлов (settings.json, email.json, calendar.json)
- Создание дефолтных конфигов при первом запуске
- Валидация данных перед сохранением
- Работа с хешированным паролем

**3. Task Scheduler Manager (scheduler.py) - Работа с планировщиком Windows**
- Создание задач в Windows Task Scheduler
- Обновление расписания существующих задач
- Включение/отключение задач (enable/disable)
- Удаление задач при удалении скрипта
- Поддержка расписания: ежедневно / только рабочие дни

**4. Utils (utils.py) - Вспомогательные функции**
- Хеширование пароля (SHA-256)
- Проверка производственного календаря РФ (рабочий/нерабочий день)
- Отправка тестовых писем через SMTP
- Сканирование папки scripts/ и поиск исполняемых файлов
- Логирование операций

### Поток данных

```
┌──────────────┐
│  Пользователь │
└──────┬───────┘
       │
       ▼
┌──────────────────────────────────────────────────────────┐
│                    GUI (gui.py)                          │
│  - Список скриптов                                       │
│  - Настройка расписания, путей                           │
│  - Окно настроек email (защищено паролем)                │
└──────┬───────────────────────────┬───────────────────────┘
       │                           │
       ▼                           ▼
┌──────────────────┐      ┌────────────────────────┐
│  Config Manager  │      │  Scheduler Manager     │
│   (config.py)    │      │   (scheduler.py)       │
│                  │      │                        │
│  - Чтение/запись │      │  - Создание задач      │
│  - Валидация     │      │  - Обновление          │
└──────┬───────────┘      └────────┬───────────────┘
       │                           │
       ▼                           ▼
┌──────────────────┐      ┌────────────────────────┐
│   JSON файлы     │      │ Windows Task Scheduler │
│  (config/)       │      └────────┬───────────────┘
└──────────────────┘               │
                                   ▼
                          ┌─────────────────────┐
                          │  Запуск скрипта     │
                          │  (subprocess)       │
                          └────────┬────────────┘
                                   │
                                   ▼
                          ┌─────────────────────┐
                          │  Скрипт читает JSON │
                          │  Обрабатывает данные│
                          │  Отправляет Email   │
                          └─────────────────────┘
```

### Принципы архитектуры

**Простота**
- Линейная архитектура без сложных паттернов
- Прямые вызовы между модулями
- Никаких абстракций "на вырост"

**Разделение ответственности**
- GUI не общается напрямую с Task Scheduler - только через scheduler.py
- Config Manager - единственное место работы с JSON
- Utils не содержит бизнес-логику, только вспомогательные функции

**Независимость скриптов**
- Скрипты полностью независимы от приложения
- Взаимодействие только через JSON конфиг
- Скрипты могут работать даже без GUI (запуск из Task Scheduler)

**Конфигурация как источник правды**
- JSON файлы - единственный источник правды
- При запуске GUI синхронизируется с конфигами
- Task Scheduler создается на основе конфигов

## Модель данных

### Структура конфигурационных файлов

**config/settings.json** - настройки приложения и скриптов
```json
{
  "password_hash": "5e884898da28047151d0e56f8dc6292773603d0d6aabbdd62a11ef721d1542d8",
  "scripts": [
    {
      "name": "script1",
      "enabled": true,
      "schedule": {
        "time": "09:00",
        "mode": "workdays"
      },
      "paths": {
        "input": "C:/data/input",
        "output": "C:/data/output"
      },
      "executable": "scripts/script1/run.py"
    },
    {
      "name": "script2",
      "enabled": false,
      "schedule": {
        "time": "14:30",
        "mode": "daily"
      },
      "paths": {
        "input": "C:/reports/source",
        "output": "C:/reports/result"
      },
      "executable": "scripts/script2/process.bat"
    }
  ]
}
```

**config/email.json** - настройки email (защищены паролем в GUI)
```json
{
  "smtp": {
    "server": "smtp.gmail.com",
    "port": 587,
    "use_tls": true,
    "username": "sender@example.com",
    "password": "app_password_here"
  },
  "from": "sender@example.com",
  "to": "recipient@example.com",
  "subject": "Отчет от ScriptManager",
  "body": "Во вложении файл отчета"
}
```

**config/calendar.json** - производственный календарь РФ
```json
{
  "year": 2025,
  "holidays": [
    "2025-01-01",
    "2025-01-02",
    "2025-01-03",
    "2025-01-06",
    "2025-02-23",
    "2025-03-08",
    "2025-05-01",
    "2025-05-09",
    "2025-06-12",
    "2025-11-04"
  ]
}
```

### Принципы работы с данными

**Простота структуры**
- Плоские структуры без глубокой вложенности
- Понятные названия полей на английском
- Минимум типов данных (строки, булевы, массивы)

**Читаемость**
- JSON с отступами для ручного редактирования
- Комментарии в документации, не в файлах

**Валидация**
- Проверка обязательных полей перед сохранением
- Проверка формата времени (HH:MM)
- Проверка существования путей
- Проверка корректности email адресов

**Доступ**
- Только config.py имеет право читать/писать конфиги
- Скрипты читают конфиги в read-only режиме
- GUI не работает с файлами напрямую

## Работа с LLM

**LLM не применимо для данного проекта**

Причины:
- Простое CRUD-приложение с четкой детерминированной логикой
- Нет задач, требующих обработки естественного языка или генерации текста
- Добавление LLM противоречит принципу KISS
- Дополнительные зависимости существенно увеличат размер дистрибутива
- Потенциальные проблемы с конфиденциальностью данных (пароли, пути к файлам)
- Непредсказуемость результатов LLM неприемлема для управления расписанием

Все задачи приложения решаются стандартными алгоритмами:
- Управление расписанием - работа с Task Scheduler API
- Валидация данных - регулярные выражения и простые проверки
- Работа с конфигами - стандартный JSON
- Запуск скриптов - subprocess

## Мониторинг и Логирование

### Логирование

**Библиотека:** встроенный модуль `logging` Python

**Что логируем:**
- Запуск/остановка приложения
- Создание/обновление/удаление задач в Task Scheduler
- Изменения конфигурации
- Ошибки при валидации данных
- Результаты отправки тестовых писем
- Ошибки при работе с файлами/путями

**Уровни логирования:**
- INFO - обычные операции (запуск, сохранение конфигов)
- WARNING - потенциальные проблемы (путь не найден, но можно продолжить)
- ERROR - ошибки, которые мешают выполнению операции

**Формат лога:**
```
2025-01-15 09:30:15 - INFO - Приложение запущено
2025-01-15 09:30:16 - INFO - Загружено 3 скрипта из папки scripts/
2025-01-15 09:35:20 - INFO - Создана задача в Task Scheduler: script1
2025-01-15 09:40:10 - ERROR - Не удалось отправить тестовое письмо: SMTP Authentication Error
```

**Хранение логов:**
- Файл: `logs/app.log`
- Ротация: при достижении 10 МБ создается новый файл (app.log.1, app.log.2)
- Хранятся последние 5 файлов логов

### Мониторинг

**Минимальный мониторинг без сложных систем:**

**В GUI отображается:**
- Статус каждого скрипта (enabled/disabled)
- Время следующего запуска
- Последний результат выполнения (из Task Scheduler History)

**Никаких внешних систем мониторинга:**
- Не используем Prometheus, Grafana и подобное
- Не отправляем метрики куда-либо
- Вся информация доступна локально в GUI и логах

**Принцип:** Пользователь открывает приложение и сразу видит состояние всех скриптов

## Сценарии работы

### Сценарий 1: Первый запуск и миграция существующих скриптов

1. Пользователь запускает ScriptManager.exe
2. Приложение создает структуру папок: `config/`, `logs/`, `scripts/`
3. Создаются дефолтные конфиги
4. Пользователю предлагается установить пароль для доступа к настройкам
5. **Миграция:** Пользователь копирует существующие скрипты в `scripts/` (6 Python файлов)
6. Открывается главное окно со списком обнаруженных скриптов

### Сценарий 2: Настройка общих параметров (первичная конфигурация)

1. Пользователь нажимает "Настройки Email/Telegram" (защищено паролем)
2. Заполняет **общие настройки email:**
   - SMTP сервер: smtp.gmail.com
   - Порт: 465
   - Email адрес: gvu.lip@gmail.com
   - App password: (скрытый ввод)
3. Заполняет **общие настройки Telegram:**
   - Bot Token
   - Personal Chat ID
   - Group Chat ID
4. Нажимает "Сохранить"
5. Приложение создает `config/email.json` и `config/telegram.json`

### Сценарий 3: Настройка отдельного скрипта

1. Пользователь выбирает скрипт из списка (например, "CMS 36")
2. Открывается окно настроек скрипта:
   - **Расписание:**
     - Время: 09:00
     - Режим: Рабочие дни
   - **Пути:**
     - Исходный файл: `C:/Users/.../Артур ЦМС.xlsx`
     - Папка для результата: `C:/Users/.../Autosend/Воронеж`
   - **Специфичные настройки email:**
     - Получатель: cms-tab@yandex.ru
     - Тема: "Отчёт: {filename}" (поддержка переменных)
     - Тело письма: (шаблон)
3. Включает скрипт (enabled = true)
4. Нажимает "Сохранить и создать задачу"
5. Приложение:
   - Сохраняет настройки в `config/scripts/cms36.json`
   - Создает задачу в Windows Task Scheduler
   - Показывает подтверждение

### Сценарий 4: Тестирование отправки

1. Пользователь в настройках скрипта нажимает "Тест отправки"
2. Приложение запускает скрипт вручную
3. Отображается консоль с выводом скрипта в реальном времени
4. По завершении показывает:
   - Успешно отправлено / Ошибка
   - Лог выполнения
   - Уведомление в Telegram (если скрипт отправил)

### Сценарий 5: Автоматическая работа

1. Task Scheduler запускает скрипт в назначенное время
2. Скрипт:
   - Читает конфиги из `config/scripts/cms36.json`, `config/email.json`, `config/telegram.json`
   - Собирает данные из Excel
   - Создает отчет
   - Отправляет email с файлом
   - Отправляет уведомление в Telegram
3. Результаты логируются в `logs/app.log`
4. При следующем открытии GUI пользователь видит статус последнего выполнения

### Сценарий 6: Обработка ошибок

1. Если скрипт падает (например, файл не найден):
   - Telegram получает уведомление об ошибке
   - Ошибка записывается в лог
   - В GUI отображается красный статус "Последний запуск: ошибка"
2. Пользователь открывает GUI, видит ошибку
3. Кликает "Посмотреть лог"
4. Исправляет путь к файлу в настройках
5. Нажимает "Тест отправки" для проверки

## Деплой и Разработка

### Окружение разработки

**Установка окружения:**
```bash
# Создание виртуального окружения
python -m venv venv

# Активация (Windows)
venv\Scripts\activate

# Установка зависимостей
pip install -r requirements.txt
```

**requirements.txt:**
```
customtkinter==5.2.0
pywin32==306
pandas==2.1.0
openpyxl==3.1.2
requests==2.31.0
```

### Автоматизация разработки

**install.bat** - установка зависимостей:
```batch
@echo off
python -m venv venv
call venv\Scripts\activate
pip install -r requirements.txt
echo Зависимости установлены!
pause
```

**run.bat** - запуск в режиме разработки:
```batch
@echo off
call venv\Scripts\activate
python main.py
pause
```

**build.bat** - сборка .exe файла:
```batch
@echo off
call venv\Scripts\activate
pyinstaller --onefile --windowed --name="ScriptManager" --icon=icon.ico main.py
echo Сборка завершена! Файл в папке dist/
pause
```

### Деплой

**Способ распространения:**
1. Собираем .exe через PyInstaller
2. Создаем папку `ScriptManager_Release/`
3. Копируем туда:
   - `ScriptManager.exe`
   - `README.txt` (инструкция по первому запуску)
   - Пример `config/calendar.json` с праздниками на текущий год
4. Архивируем в .zip
5. Передаем пользователю

**Установка у пользователя:**
1. Распаковать архив в любую папку
2. Запустить `ScriptManager.exe`
3. При первом запуске установить пароль
4. Скопировать существующие скрипты в папку `scripts/`
5. Настроить каждый скрипт через GUI

**Обновление:**
- Просто заменить `ScriptManager.exe` на новую версию
- Все конфиги и скрипты сохраняются

### Конфигурирование

**Файловая структура конфигов:**

```
config/
├── settings.json          # Пароль и глобальные настройки
├── email.json             # Общие настройки SMTP
├── telegram.json          # Токены Telegram
├── calendar.json          # Производственный календарь РФ
└── scripts/               # Настройки каждого скрипта
    ├── cms36.json
    ├── cms48.json
    ├── kvadra.json
    ├── lts.json
    ├── rvk1.json
    └── rvk2.json
```

**Пример config/scripts/cms36.json:**
```json
{
  "name": "CMS 36",
  "enabled": true,
  "executable": "scripts/CMS 36/ArshinCMSvrn_gmail.py",
  "schedule": {
    "time": "09:00",
    "mode": "workdays"
  },
  "paths": {
    "input_file": "C:/Users/vdaur/OneDrive/.../Артур ЦМС.xlsx",
    "output_folder": "C:/Users/vdaur/OneDrive/.../Воронеж",
    "sheet_name": "Аршин ЦМС vrn"
  },
  "email": {
    "recipient": "cms-tab@yandex.ru",
    "subject": "Отчёт: {filename}",
    "body": "Добрый день!\n\nВо вложении файл с отчётом.\n\nС уважением, service pvk48"
  },
  "specific": {
    "summary_file": "C:/Users/vdaur/OneDrive/.../Итоговая аршин ЦМС.xlsx"
  }
}
```

**Пример config/telegram.json:**
```json
{
  "bot_token": "YOUR_BOT_TOKEN_HERE",
  "personal_chat_id": "312809937",
  "group_chat_id": "-1002458835967"
}
```

### Рефакторинг существующих скриптов

**Что нужно изменить в скриптах:**

1. **Убрать константы из кода**, заменить на чтение из JSON:
```python
import json

# Было:
# FILE_PATH = "C:/Users/..."
# EMAIL_ADDRESS = "gvu.lip@gmail.com"

# Стало:
with open('config/scripts/cms36.json', 'r', encoding='utf-8') as f:
    script_config = json.load(f)

with open('config/email.json', 'r', encoding='utf-8') as f:
    email_config = json.load(f)

with open('config/telegram.json', 'r', encoding='utf-8') as f:
    telegram_config = json.load(f)

FILE_PATH = script_config['paths']['input_file']
EMAIL_ADDRESS = email_config['smtp']['username']
TELEGRAM_TOKEN = telegram_config['bot_token']
```

2. **Структура остается прежней:**
   - Функции `gather_data()`, `create_output_file()`, `send_email()` не меняются
   - Только замена констант на чтение из конфигов

3. **Обратная совместимость:**
   - Если JSON не найден, использовать захардкоженные значения (fallback)
   - Логировать предупреждение

**Принцип:** Минимальные изменения в скриптах, максимум функциональности в GUI

---

## ВАЖНОЕ ОБНОВЛЕНИЕ ПОДХОДА

После анализа существующих скриптов принято решение о **полной интеграции** скриптов в приложение:

### Ключевые изменения подхода:

1. **Скрипты встроены в приложение**
   - Не пользовательские, а часть кода приложения
   - Находятся в `src/scripts/` как Python модули
   - 6 фиксированных скриптов: CMS36, CMS48, Kvadra, LTS, RVK1, RVK2

2. **Период сбора данных настраивается**
   - Каждый скрипт имеет: `collection_start_day` (например, "Thursday")
   - Период сбора: `collection_days` (например, 7)
   - **Умный перенос:** если день отправки выпадает на выходной, автоматический перенос на следующий рабочий день

3. **Никаких добавлений скриптов пользователем**
   - GUI только настраивает 6 существующих скриптов
   - Упрощает архитектуру
   - Исключает проблемы с неизвестными скриптами

### Обновленная структура проекта:

```
ScriptManager/
│
├── src/
│   ├── main.py
│   ├── gui.py
│   ├── scheduler.py
│   ├── config_manager.py
│   ├── utils.py
│   │
│   └── scripts/          # Встроенные скрипты как модули
│       ├── __init__.py
│       ├── base_script.py         # Базовый класс для всех скриптов
│       ├── cms36_script.py        # CMS 36 Воронеж
│       ├── cms48_script.py        # CMS 48
│       ├── kvadra_script.py       # Kvadra
│       ├── lts_script.py          # Липецктеплосеть
│       ├── rvk_kvadra_script.py   # RVK Kvadra
│       └── rvk_script.py          # RVK
│
├── config/
│   ├── settings.json
│   ├── email.json
│   ├── telegram.json
│   ├── calendar.json
│   └── scripts.json      # Настройки всех 6 скриптов
│
├── logs/
│   └── app.log
│
└── requirements.txt
```

### Обновленная модель данных config/scripts.json:

```json
{
  "cms36": {
    "name": "CMS 36 Воронеж",
    "enabled": true,
    "schedule": {
      "send_day": "Thursday",
      "send_time": "09:00",
      "mode": "workdays",
      "smart_reschedule": true
    },
    "collection": {
      "start_day": "Thursday",
      "duration_days": 7,
      "reference_day": "last_thursday"
    },
    "paths": {
      "input_file": "C:/Users/vdaur/OneDrive/.../Артур ЦМС.xlsx",
      "output_folder": "C:/Users/vdaur/OneDrive/.../Воронеж",
      "sheet_name": "Аршин ЦМС vrn",
      "summary_file": "C:/Users/vdaur/OneDrive/.../Итоговая аршин ЦМС.xlsx"
    },
    "email": {
      "recipient": "cms-tab@yandex.ru",
      "subject_template": "Отчёт: {filename}",
      "body_template": "Добрый день!\n\nВо вложении файл с отчётом.\n\nС уважением, service pvk48"
    }
  },
  "cms48": {
    "name": "CMS 48",
    "enabled": true,
    "schedule": {
      "send_day": "Friday",
      "send_time": "10:00",
      "mode": "workdays",
      "smart_reschedule": true
    },
    "collection": {
      "start_day": "Friday",
      "duration_days": 7,
      "reference_day": "last_friday"
    },
    "paths": {
      "input_file": "...",
      "output_folder": "...",
      "sheet_name": "..."
    },
    "email": {
      "recipient": "...",
      "subject_template": "...",
      "body_template": "..."
    }
  }
}
```

### Базовый класс для скриптов (src/scripts/base_script.py):

```python
from abc import ABC, abstractmethod
import json
from datetime import datetime, timedelta
from typing import Dict, Any

class BaseScript(ABC):
    """Базовый класс для всех скриптов сбора данных"""

    def __init__(self, script_id: str):
        self.script_id = script_id
        self.config = self.load_config()
        self.email_config = self.load_email_config()
        self.telegram_config = self.load_telegram_config()

    def load_config(self) -> Dict[str, Any]:
        """Загрузка конфига скрипта"""
        with open('config/scripts.json', 'r', encoding='utf-8') as f:
            all_configs = json.load(f)
        return all_configs[self.script_id]

    def load_email_config(self) -> Dict[str, Any]:
        with open('config/email.json', 'r', encoding='utf-8') as f:
            return json.load(f)

    def load_telegram_config(self) -> Dict[str, Any]:
        with open('config/telegram.json', 'r', encoding='utf-8') as f:
            return json.load(f)

    def calculate_collection_period(self) -> tuple:
        """
        Вычисляет период сбора данных на основе настроек
        Возвращает: (start_date, end_date)
        """
        collection = self.config['collection']
        today = datetime.now().date()

        # Находим последний день начала сбора (например, последний четверг)
        # Логика расчета на основе collection_start_day и reference_day

        start_date = ...  # Вычисляется динамически
        end_date = start_date + timedelta(days=collection['duration_days'])

        return start_date, end_date

    @abstractmethod
    def gather_data(self):
        """Сбор данных из источника (переопределяется в каждом скрипте)"""
        pass

    @abstractmethod
    def create_output_file(self, data):
        """Создание выходного файла (переопределяется в каждом скрипте)"""
        pass

    def send_email(self, filename: str, filepath: str):
        """Общий метод отправки email (одинаков для всех)"""
        # Использует self.email_config и self.config['email']
        pass

    def send_telegram_notification(self, message: str, to_group: bool = False):
        """Общий метод отправки в Telegram (одинаков для всех)"""
        # Использует self.telegram_config
        pass

    def run(self):
        """Основной метод выполнения скрипта"""
        try:
            data = self.gather_data()
            if data is not None and len(data) > 0:
                filename, filepath = self.create_output_file(data)
                self.send_email(filename, filepath)
                self.send_telegram_notification(f"Файл {filename} успешно отправлен", to_group=True)
            else:
                self.send_telegram_notification("Нет данных для отправки", to_group=True)
        except Exception as e:
            self.send_telegram_notification(f"ОШИБКА: {e}", to_group=True)
            raise
```

### Пример конкретного скрипта (src/scripts/cms36_script.py):

```python
from .base_script import BaseScript
import pandas as pd
from datetime import datetime

class CMS36Script(BaseScript):
    """Скрипт для CMS 36 Воронеж"""

    def __init__(self):
        super().__init__('cms36')

    def gather_data(self):
        """Специфичная логика сбора данных для CMS 36"""
        paths = self.config['paths']

        # Расчет периода сбора
        start_date, end_date = self.calculate_collection_period()

        # Чтение Excel
        df = pd.read_excel(
            paths['input_file'],
            sheet_name=paths['sheet_name']
        )

        # Фильтрация по датам (та же логика, что была раньше)
        df = df[df['№ в ГРСИ'].notna() & ...]

        # Фильтрация по периоду сбора
        df['Дата поверки'] = pd.to_datetime(df['Дата поверки'], errors='coerce')
        df = df[(df['Дата поверки'] >= start_date) & (df['Дата поверки'] <= end_date)]

        return df

    def create_output_file(self, data):
        """Специфичная логика создания файла для CMS 36"""
        paths = self.config['paths']

        # Та же логика формирования имени файла
        count_rows = len(data)
        dates = data['Дата поверки'].dropna().tolist()

        if dates:
            min_date = min(dates).strftime('%d.%m')
            max_date = max(dates).strftime('%d.%m.%y')
            date_range = f"{min_date} - {max_date}"
        else:
            date_range = "no_dates"

        filename = f"420 ({count_rows}) {date_range} Воронеж"
        filepath = os.path.join(paths['output_folder'], f"{filename}.xlsx")

        # Сохранение
        data.to_excel(filepath, index=False)

        return filename, filepath
```

### Умное управление расписанием:

**scheduler.py** теперь проверяет:
1. Запланированный день отправки (например, четверг)
2. Если четверг - выходной (проверка по calendar.json)
3. Автоматически переносит на следующий рабочий день
4. Создает задачу в Task Scheduler с правильной датой

**Алгоритм переноса:**
```python
def get_next_workday(planned_date: datetime, calendar: dict) -> datetime:
    """Находит следующий рабочий день если planned_date - выходной"""
    current_date = planned_date
    while is_holiday(current_date, calendar):
        current_date += timedelta(days=1)
    return current_date
```

### Механизм планирования и выполнения задач

#### Архитектура запуска скриптов

**Embedded Python + прямой запуск через Task Scheduler**

Приложение использует встроенный дистрибутив Python (embedded Python) для портативности:
- Python включен в папку приложения (`python/`)
- Task Scheduler запускает скрипты напрямую через `python/python.exe`
- Не требуется установка Python в системе
- Полная переносимость на любой компьютер

**Структура запуска:**
```
Windows Task Scheduler
    ↓
python\python.exe src\scripts\runner.py cms36
    ↓
CMS36Script().run()
    ↓
gather_data() → create_output_file() → send_email() → check_calendar() → reschedule_next_run()
```

#### Типы расписаний

**1. Еженедельные скрипты** (RVK, RVK Kvadra, LTS, Kvadra)
- Отправка каждую неделю в определенный день (четверг, пятница и т.д.)
- Период сбора: 7 дней
- `"schedule_type": "weekly"` в конфиге

**2. Ежедневные скрипты** (CMS 36 Воронеж, CMS 48 Липецк)
- Отправка каждый рабочий день в Аршин
- Период сбора: данные за сегодня (по дате внесения/поверки)
- `"schedule_type": "daily"` в конфиге

#### Умное управление расписанием

**Проблема:** Если день отправки выпадает на праздник/выходной, письмо не должно отправляться в этот день.

**Решение:** После каждой успешной отправки скрипт автоматически вычисляет следующую дату запуска с учетом календаря.

#### Алгоритм работы скрипта (пошагово)

**Шаг 1: Task Scheduler запускает скрипт**
```
python\python.exe src\scripts\runner.py cms36
```

**Шаг 2: Выполнение основной логики**
```python
def run(self):
    try:
        # 1. Сбор данных
        data = self.gather_data()

        # 2. Создание файла
        if data is not None and len(data) > 0:
            filename, filepath = self.create_output_file(data)

            # 3. Отправка email
            self.send_email(filename, filepath)

            # 4. Уведомление в Telegram
            self.send_telegram_notification(f"Файл {filename} успешно отправлен", to_group=True)

            # 5. ВАЖНО: Пересчет и обновление следующего запуска
            self.reschedule_next_run()
        else:
            self.send_telegram_notification("Нет данных для отправки", to_group=True)
    except Exception as e:
        self.send_telegram_notification(f"ОШИБКА: {e}", to_group=True)
        raise
```

**Шаг 3: Пересчет следующей даты запуска**
```python
def reschedule_next_run(self):
    """
    Вычисляет следующую дату запуска с учетом:
    - Периодичности (daily/weekly)
    - Производственного календаря
    - Режима работы (workdays/daily)
    """
    schedule = self.config['schedule']
    schedule_type = schedule.get('schedule_type', 'weekly')

    if schedule_type == 'weekly':
        # Для еженедельных: следующий четверг (или другой день недели)
        next_date = self.get_next_weekday(schedule['send_day'])
    else:
        # Для ежедневных: завтра
        next_date = datetime.now().date() + timedelta(days=1)

    # Проверка календаря: если выходной - ищем следующий рабочий день
    if schedule.get('smart_reschedule', True):
        next_date = self.get_next_workday(next_date)

    # Обновление задачи в Task Scheduler
    self.update_task_scheduler(next_date, schedule['send_time'])
```

**Шаг 4: Обновление Task Scheduler**
```python
def update_task_scheduler(self, next_date: date, time_str: str):
    """
    Обновляет задачу в Windows Task Scheduler
    Устанавливает новую дату и время следующего запуска
    """
    from scheduler import TaskSchedulerManager

    scheduler = TaskSchedulerManager()
    scheduler.update_task_date(
        task_name=f"ScriptManager_{self.script_id}",
        next_run_date=next_date,
        next_run_time=time_str
    )
```

#### Проверка производственного календаря

**Функция определения рабочего дня:**
```python
def get_next_workday(self, planned_date: date) -> date:
    """
    Находит следующий рабочий день если planned_date - выходной/праздник

    Алгоритм:
    1. Проверяем planned_date в calendar.json
    2. Если суббота/воскресенье или праздник → переходим к следующему дню
    3. Повторяем до первого рабочего дня
    """
    calendar = self.load_calendar()
    current_date = planned_date

    while self.is_holiday(current_date, calendar):
        current_date += timedelta(days=1)

    return current_date

def is_holiday(self, check_date: date, calendar: dict) -> bool:
    """
    Проверяет, является ли дата выходным

    Возвращает True если:
    - Суббота (weekday == 5)
    - Воскресенье (weekday == 6)
    - Праздник из списка calendar['holidays']
    """
    # Выходные
    if check_date.weekday() in [5, 6]:
        return True

    # Праздники РФ
    if check_date.strftime('%Y-%m-%d') in calendar.get('holidays', []):
        return True

    return False
```

#### Примеры работы умного расписания

**Пример 1: Еженедельный скрипт (RVK - Росводоканал)**

Конфигурация:
```json
{
  "schedule": {
    "send_day": "Thursday",
    "send_time": "09:00",
    "schedule_type": "weekly",
    "smart_reschedule": true
  }
}
```

Сценарий работы:
1. **19 декабря 2025 (четверг, рабочий день):**
   - Task Scheduler запускает скрипт в 09:00
   - Скрипт собирает данные с 12.12 по 19.12 (последние 7 дней)
   - Отправляет письмо на ev.shlykova@rosvodokanal.ru
   - Вычисляет следующую дату: 26 декабря (четверг)
   - Обновляет задачу в Task Scheduler на 26.12.2025 09:00

2. **26 декабря 2025 (четверг, но попадает на новогодние каникулы):**
   - Скрипт НЕ запускается в этот день
   - При предыдущем пересчете (19 декабря) было определено:
     - 26.12 - праздник (проверка по calendar.json)
     - Следующий рабочий день: 09 января 2026
     - Задача создана на 09.01.2026 09:00

3. **9 января 2026 (четверг, первый рабочий день):**
   - Task Scheduler запускает скрипт в 09:00
   - Скрипт собирает данные с 26.12 по 09.01
   - Отправляет письмо
   - Вычисляет следующую дату: 16 января (четверг, рабочий день)
   - Обновляет задачу на 16.01.2026 09:00

**Пример 2: Ежедневный скрипт (CMS 36 - Воронеж Аршин)**

Конфигурация:
```json
{
  "schedule": {
    "send_time": "10:00",
    "schedule_type": "daily",
    "mode": "workdays",
    "smart_reschedule": true
  }
}
```

Сценарий работы:
1. **Понедельник 22.12.2025:**
   - Запуск в 10:00
   - Собирает данные по дате поверки за последние 30 дней
   - Отправляет файл в систему Аршин (cms-tab@yandex.ru)
   - Следующий день: 23.12 (вторник, рабочий)
   - Задача на 23.12.2025 10:00

2. **Пятница 26.12.2025:**
   - Запуск в 10:00
   - Собирает данные за последние 30 дней
   - Отправляет письмо
   - Следующий день: 27.12 (суббота) → пропуск
   - 28.12 (воскресенье) → пропуск
   - 29.12-31.12 (праздники) → пропуск
   - Следующий рабочий: 09.01.2026
   - Задача на 09.01.2026 10:00

#### Интеграция с Windows Task Scheduler

**Создание задачи через GUI:**

Когда пользователь включает скрипт и нажимает "Сохранить", приложение:
1. Рассчитывает первую дату запуска (с учетом календаря)
2. Создает задачу в Task Scheduler с параметрами:
   - Имя: `ScriptManager_cms36`
   - Путь: `C:\...\python\python.exe`
   - Аргументы: `src\scripts\runner.py cms36`
   - Триггер: Однократный запуск на вычисленную дату и время
   - Права: Запуск от текущего пользователя

**Код создания задачи (scheduler.py):**
```python
def create_task(self, script_id: str, next_run_date: date, time_str: str, python_path: str):
    """
    Создает задачу в Windows Task Scheduler

    Важно: Задача создается как ОДНОКРАТНАЯ, а не повторяющаяся
    Следующий запуск устанавливается самим скриптом после выполнения
    """
    import win32com.client

    scheduler = win32com.client.Dispatch('Schedule.Service')
    scheduler.Connect()

    root_folder = scheduler.GetFolder('\\')
    task_def = scheduler.NewTask(0)

    # Триггер: однократный запуск
    trigger = task_def.Triggers.Create(1)  # 1 = TASK_TRIGGER_TIME
    trigger.StartBoundary = f"{next_run_date.isoformat()}T{time_str}:00"
    trigger.Enabled = True

    # Действие: запуск Python скрипта
    action = task_def.Actions.Create(0)  # 0 = TASK_ACTION_EXEC
    action.Path = python_path
    action.Arguments = f"src\\scripts\\runner.py {script_id}"
    action.WorkingDirectory = os.getcwd()

    # Регистрация задачи
    root_folder.RegisterTaskDefinition(
        f"ScriptManager_{script_id}",
        task_def,
        6,  # TASK_CREATE_OR_UPDATE
        None, None, 3  # TASK_LOGON_INTERACTIVE_TOKEN
    )

def update_task_date(self, task_name: str, next_run_date: date, next_run_time: str):
    """
    Обновляет дату следующего запуска существующей задачи
    Вызывается скриптом после успешной отправки
    """
    import win32com.client

    scheduler = win32com.client.Dispatch('Schedule.Service')
    scheduler.Connect()

    root_folder = scheduler.GetFolder('\\')
    task = root_folder.GetTask(task_name)
    task_def = task.Definition

    # Обновляем триггер
    trigger = task_def.Triggers.Item(1)
    trigger.StartBoundary = f"{next_run_date.isoformat()}T{next_run_time}:00"

    # Сохраняем изменения
    root_folder.RegisterTaskDefinition(
        task_name,
        task_def,
        6,  # TASK_CREATE_OR_UPDATE
        None, None, 3
    )
```

#### Обработка ошибок и сбоев

**Что происходит при ошибке скрипта:**

1. **Скрипт упал до отправки письма:**
   - Telegram получает уведомление об ошибке
   - Запись в `logs/app.log`
   - Задача в Task Scheduler НЕ обновляется
   - **При следующем запуске Task Scheduler повторит попытку (через 7 дней/1 день)**

2. **Скрипт отправил письмо, но упал при пересчете даты:**
   - Письмо отправлено успешно
   - Telegram получил уведомление об успехе
   - Запись ошибки в лог
   - **Задача НЕ обновлена → запустится снова в то же время через неделю/день**
   - Риск: дублирование отправки
   - **Защита:** проверка на существование файла в `create_output_file()`

3. **Task Scheduler не смог запустить скрипт:**
   - Просмотр через Task Scheduler History
   - GUI показывает статус "Ошибка запуска"
   - Пользователь может открыть лог и увидеть причину

**Механизм защиты от дублей:**
```python
def create_output_file(self, data):
    """Создание выходного файла с проверкой на существование"""
    filename = self.generate_filename(data)
    filepath = os.path.join(self.config['paths']['output_folder'], filename)

    # Защита от дублирования
    if os.path.exists(filepath):
        error_msg = f"Дубль файла {filename}, проверь! Реестр не отправлен!"
        self.send_telegram_notification(error_msg, to_group=True)
        raise FileExistsError(error_msg)

    data.to_excel(filepath, index=False)
    return filename, filepath
```

#### Портативность и встроенный Python

**Структура папки приложения:**
```
ScriptManager/
├── ScriptManager.exe           # Главный GUI
├── python/                     # Embedded Python
│   ├── python.exe
│   ├── python310.dll
│   └── Lib/
│       └── site-packages/
│           ├── pandas/
│           ├── openpyxl/
│           └── requests/
├── src/
│   └── scripts/
│       ├── runner.py           # Точка входа для Task Scheduler
│       ├── base_script.py
│       ├── cms36_script.py
│       └── ...
├── config/
└── logs/
```

**Преимущества embedded Python:**
- ✅ Не требует установки Python в системе
- ✅ Полная изоляция от системного Python
- ✅ Контролируемые версии библиотек
- ✅ Портативность - можно скопировать на любой компьютер
- ✅ Нет конфликтов версий с другими Python приложениями

**runner.py - точка входа для Task Scheduler:**
```python
"""
Запускается Task Scheduler'ом для выполнения конкретного скрипта
Использование: python runner.py <script_id>
Пример: python runner.py cms36
"""
import sys
from cms36_script import CMS36Script
from cms48_script import CMS48Script
from lts_script import LTSScript
# ... импорты других скриптов

SCRIPTS_MAP = {
    'cms36': CMS36Script,
    'cms48': CMS48Script,
    'kvadra': KvadraScript,
    'lts': LTSScript,
    'rvk1': RVK1Script,
    'rvk2': RVK2Script
}

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Использование: python runner.py <script_id>")
        sys.exit(1)

    script_id = sys.argv[1]

    if script_id not in SCRIPTS_MAP:
        print(f"Неизвестный скрипт: {script_id}")
        sys.exit(1)

    # Создание и запуск скрипта
    script = SCRIPTS_MAP[script_id]()
    script.run()
```

#### Мониторинг выполнения

**В GUI отображается:**
- **Статус скрипта:** Включен/Отключен
- **Следующий запуск:** 19.12.2025 09:00
- **Последнее выполнение:** 12.12.2025 09:05 - Успешно (или Ошибка)
- **Последний файл:** "420 (15) 05.12 - 12.12.25 Воронеж.xlsx"

**Получение статуса из Task Scheduler:**
```python
def get_task_status(self, task_name: str) -> dict:
    """
    Получает информацию о задаче из Task Scheduler

    Возвращает:
    - next_run: дата и время следующего запуска
    - last_run: дата и время последнего запуска
    - last_result: код результата (0 = успех, другое = ошибка)
    """
    import win32com.client

    scheduler = win32com.client.Dispatch('Schedule.Service')
    scheduler.Connect()

    task = scheduler.GetFolder('\\').GetTask(task_name)

    return {
        'next_run': task.NextRunTime,
        'last_run': task.LastRunTime,
        'last_result': task.LastTaskResult,
        'enabled': task.Enabled
    }
```

### Преимущества нового подхода:

✅ **Полный контроль** - все скрипты в одном месте
✅ **Единый код** - переиспользование через базовый класс
✅ **Никаких сюрпризов** - нет неизвестных скриптов
✅ **Умное расписание** - автоматический перенос на рабочие дни
✅ **Гибкая настройка** - каждый скрипт настраивается через GUI
✅ **Простая поддержка** - изменения в одном месте
✅ **Автоматическое перепланирование** - скрипты сами управляют следующим запуском
✅ **Защита от дублей** - проверка существования файлов
✅ **Полная портативность** - embedded Python, не требует установки

