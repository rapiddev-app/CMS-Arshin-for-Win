"""
Управление конфигурационными файлами ScriptManager
"""

import json
import os
import logging
import requests
from datetime import datetime
from typing import Dict, Any


class ConfigManager:
    """Класс для работы с JSON конфигурационными файлами"""

    def __init__(self):
        """Инициализация и создание структуры config/.

        По умолчанию на Windows конфиги берутся из `%APPDATA%\\ScriptManager\\config`,
        а "шаблонные" конфиги — из папки установки/проекта `config/`.

        Для тестов и отладки можно принудительно указать папку конфигов через
        переменную окружения `SCRIPTMANAGER_CONFIG_DIR` (или `SCRIPT_MANAGER_CONFIG_DIR`).

        :raises OSError: Если не удалось создать папку конфигов.
        """
        # Получаем корневую директорию проекта (Program Files)
        current_dir = os.path.dirname(os.path.abspath(__file__))
        self.project_root = os.path.dirname(current_dir)
        
        # Путь к "шаблонным" конфигам в папке установки
        self.install_config_dir = os.path.join(self.project_root, "config")

        # Принудительное указание каталога конфигов (удобно для тестов)
        forced_config_dir = (
            os.environ.get("SCRIPTMANAGER_CONFIG_DIR")
            or os.environ.get("SCRIPT_MANAGER_CONFIG_DIR")
        )

        # Определяем рабочую папку конфигурации (APPDATA)
        import sys
        if forced_config_dir:
            self.config_dir = forced_config_dir
        elif sys.platform == 'win32':
            app_data = os.environ.get('APPDATA')
            if app_data:
                self.config_dir = os.path.join(app_data, 'ScriptManager', 'config')
            else:
                self.config_dir = self.install_config_dir
        else:
            self.config_dir = self.install_config_dir

        # Создаём рабочую папку config если её нет
        if not os.path.exists(self.config_dir):
            os.makedirs(self.config_dir)
            logging.info(f"Создана папка config: {self.config_dir}")

        # === Миграция/Копирование шаблонов ===
        # Если в рабочей папке нет конфигов, копируем их из папки установки
        self._ensure_configs_exist()

        # Пути к конфигам (теперь указывают на APPDATA)
        self.settings_path = os.path.join(self.config_dir, "settings.json")
        self.email_gmail_path = os.path.join(self.config_dir, "email_gmail.json")
        self.email_yandex_path = os.path.join(self.config_dir, "email_yandex.json")
        self.telegram_path = os.path.join(self.config_dir, "telegram.json")
        self.calendar_path = os.path.join(self.config_dir, "calendar.json")
        self.scripts_path = os.path.join(self.config_dir, "scripts.json")

        # Создаём дефолтные конфиги если их нет (или копируем)
        self.create_default_configs()

    def _ensure_configs_exist(self):
        """Копирование конфигов из папки установки в APPDATA, если используем APPDATA"""
        if self.config_dir == self.install_config_dir:
            return

        import shutil
        
        # Список файлов для копирования
        config_files = [
            "settings.json", "email_gmail.json", "email_yandex.json", 
            "telegram.json", "calendar.json", "scripts.json"
        ]

        for filename in config_files:
            target_path = os.path.join(self.config_dir, filename)
            source_path = os.path.join(self.install_config_dir, filename)

            # Если файла нет в APPDATA, но он есть в Program Files -> копируем
            if not os.path.exists(target_path) and os.path.exists(source_path):
                try:
                    shutil.copy2(source_path, target_path)
                    logging.info(f"Скопирован конфиг {filename} в {self.config_dir}")
                except Exception as e:
                    logging.error(f"Не удалось скопировать {filename}: {e}")

    def _load_json(self, file_path: str) -> Dict[str, Any]:
        """
        Загрузка JSON файла

        Args:
            file_path: путь к JSON файлу

        Returns:
            Словарь с данными из JSON
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            logging.info(f"Загружен конфиг: {os.path.basename(file_path)}")
            return data
        except FileNotFoundError:
            logging.error(f"Файл не найден: {file_path}")
            raise
        except json.JSONDecodeError as e:
            logging.error(f"Ошибка парсинга JSON в {file_path}: {e}")
            raise

    def _save_json(self, file_path: str, data: Dict[str, Any]):
        """
        Сохранение данных в JSON файл

        Args:
            file_path: путь к JSON файлу
            data: данные для сохранения
        """
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            logging.info(f"Сохранён конфиг: {os.path.basename(file_path)}")
        except Exception as e:
            logging.error(f"Ошибка сохранения {file_path}: {e}")
            raise

    # ============ Методы для settings.json ============

    def load_settings(self) -> Dict[str, Any]:
        """Загрузка основных настроек"""
        return self._load_json(self.settings_path)

    def save_settings(self, settings: Dict[str, Any]):
        """Сохранение основных настроек"""
        self._save_json(self.settings_path, settings)

    # ============ Методы для email конфигов ============

    def load_email_config(self, provider: str = "gmail") -> Dict[str, Any]:
        """
        Загрузка email конфигурации

        Args:
            provider: "gmail" или "yandex"

        Returns:
            Словарь с настройками email
        """
        if provider == "gmail":
            return self._load_json(self.email_gmail_path)
        elif provider == "yandex":
            return self._load_json(self.email_yandex_path)
        else:
            raise ValueError(f"Неизвестный провайдер email: {provider}")

    def save_email_config(self, email_config: Dict[str, Any], provider: str = "gmail"):
        """
        Сохранение email конфигурации

        Args:
            email_config: настройки email
            provider: "gmail" или "yandex"
        """
        if provider == "gmail":
            self._save_json(self.email_gmail_path, email_config)
        elif provider == "yandex":
            self._save_json(self.email_yandex_path, email_config)
        else:
            raise ValueError(f"Неизвестный провайдер email: {provider}")

    # ============ Методы для telegram.json ============

    def load_telegram_config(self) -> Dict[str, Any]:
        """Загрузка Telegram конфигурации"""
        return self._load_json(self.telegram_path)

    def save_telegram_config(self, telegram_config: Dict[str, Any]):
        """Сохранение Telegram конфигурации"""
        self._save_json(self.telegram_path, telegram_config)

    # ============ Методы для calendar.json ============

    def load_calendar(self) -> Dict[str, Any]:
        """
        Загрузка календаря
        Проверяет актуальность года и обновляет если необходимо
        """
        calendar = self._load_json(self.calendar_path)

        # Проверяем актуальность года
        current_year = datetime.now().year
        if calendar.get("year", 0) < current_year:
            logging.warning(f"Календарь устарел (год {calendar.get('year')}), обновляем...")
            self.update_calendar_from_api(current_year)
            calendar = self._load_json(self.calendar_path)

        return calendar

    def update_calendar_from_api(self, year: int = None):
        """
        Загружает календарь с API isdayoff.ru

        Args:
            year: год для загрузки (по умолчанию текущий)
        """
        if year is None:
            year = datetime.now().year

        try:
            # Запрос к API isdayoff.ru для получения календаря на год
            url = f"https://isdayoff.ru/api/getdata?year={year}"
            response = requests.get(url, timeout=10)
            response.raise_for_status()

            # Ответ - строка из 365-366 символов (0=рабочий, 1=выходной, 2=сокращённый)
            days_string = response.text.strip()

            calendar_data = {
                "year": year,
                "days": days_string,
                "last_update": datetime.now().strftime("%Y-%m-%d")
            }

            self._save_json(self.calendar_path, calendar_data)
            logging.info(f"Календарь на {year} год загружен с API isdayoff.ru")

        except requests.RequestException as e:
            logging.error(f"Ошибка загрузки календаря с API: {e}")
            logging.warning("Создаётся резервный календарь с основными праздниками")
            self._create_fallback_calendar(year)

    def _create_fallback_calendar(self, year: int):
        """
        Создаёт резервный календарь с основными праздниками РФ
        Используется если API недоступен
        """
        # Основные праздники РФ 2025 (индексы дней с 1 января)
        # 1-8 января (новогодние), 23 февраля, 8 марта, 1 мая, 9 мая, 12 июня, 4 ноября
        # Плюс субботы и воскресенья

        # Создаём строку из 365 дней (0 по умолчанию)
        days = ['0'] * 365

        # Отмечаем субботы и воскресенья (упрощённо, без точного расчёта)
        # Для 2025: 1 января - среда
        # Каждый 6-й и 7-й день недели
        for day in range(365):
            weekday = (day + 2) % 7  # 1 января 2025 = среда (2)
            if weekday == 5 or weekday == 6:  # Суббота или воскресенье
                days[day] = '1'

        # Основные праздники 2025 (дни с 1 января, индекс с 0)
        holidays = [
            0, 1, 2, 3, 4, 5, 6, 7,  # 1-8 января
            53,  # 23 февраля
            66,  # 8 марта
            120,  # 1 мая
            128,  # 9 мая
            162,  # 12 июня
            307,  # 4 ноября
        ]

        for holiday in holidays:
            if holiday < len(days):
                days[holiday] = '1'

        days_string = ''.join(days)

        calendar_data = {
            "year": year,
            "days": days_string,
            "last_update": datetime.now().strftime("%Y-%m-%d"),
            "source": "fallback"
        }

        self._save_json(self.calendar_path, calendar_data)
        logging.info(f"Создан резервный календарь на {year} год")

    # ============ Создание дефолтных конфигов ============

    def create_default_configs(self):
        """Создание всех дефолтных конфигов если их нет"""

        # 1. settings.json
        if not os.path.exists(self.settings_path):
            default_settings = {
                "password_hash": "",
                "last_update": datetime.now().strftime("%Y-%m-%d")
            }
            self._save_json(self.settings_path, default_settings)
            logging.info("Создан дефолтный settings.json")

        # 2. email_gmail.json
        if not os.path.exists(self.email_gmail_path):
            default_gmail = {
                "smtp": {
                    "server": "smtp.gmail.com",
                    "port": 465,
                    "use_ssl": True,
                    "username": "gvu.lip@gmail.com",
                    "password": "ecrm vrjp kvwg qdax"
                }
            }
            self._save_json(self.email_gmail_path, default_gmail)
            logging.info("Создан дефолтный email_gmail.json")

        # 3. email_yandex.json
        if not os.path.exists(self.email_yandex_path):
            default_yandex = {
                "smtp": {
                    "server": "smtp.yandex.ru",
                    "port": 465,
                    "use_ssl": True,
                    "username": "cms.lip@yandex.ru",
                    "password": "jdmtcwmccjdhjjoe"
                }
            }
            self._save_json(self.email_yandex_path, default_yandex)
            logging.info("Создан дефолтный email_yandex.json")

        # 4. telegram.json
        if not os.path.exists(self.telegram_path):
            default_telegram = {
                "bot_token": "YOUR_BOT_TOKEN_HERE",
                "personal_chat_id": "312809937",
                "group_chat_id": "-1002458835967"
            }
            self._save_json(self.telegram_path, default_telegram)
            logging.info("Создан дефолтный telegram.json")

        # 5. calendar.json - загружаем с API или создаём резервный
        if not os.path.exists(self.calendar_path):
            current_year = datetime.now().year
            logging.info(f"Загружаем календарь на {current_year} год...")
            self.update_calendar_from_api(current_year)

    # ============ Методы для работы со scripts.json ============

    def load_scripts_config(self) -> Dict[str, Any]:
        """
        Загрузка конфигурации всех скриптов

        Returns:
            Словарь с настройками всех скриптов
        """
        return self._load_json(self.scripts_path)

    def load_script_config(self, script_id: str) -> Dict[str, Any]:
        """
        Загрузка конфигурации конкретного скрипта

        Args:
            script_id: идентификатор скрипта (cms36, cms48, kvadra, lts, rvk, rvk_kvadra)

        Returns:
            Словарь с настройками скрипта
        """
        scripts = self.load_scripts_config()
        if script_id not in scripts:
            raise ValueError(f"Скрипт {script_id} не найден в конфигурации")
        return scripts[script_id]

    def save_script_config(self, script_id: str, config: Dict[str, Any]):
        """
        Сохранение конфигурации конкретного скрипта

        Args:
            script_id: идентификатор скрипта
            config: словарь с настройками скрипта
        """
        scripts = self.load_scripts_config()
        scripts[script_id] = config
        self._save_json(self.scripts_path, scripts)
        logging.info(f"Сохранена конфигурация скрипта: {script_id}")

    def update_script_field(self, script_id: str, field: str, value: Any):
        """
        Обновление отдельного поля в конфигурации скрипта

        Args:
            script_id: идентификатор скрипта
            field: название поля (enabled, source_file, email_recipient и т.д.)
            value: новое значение
        """
        config = self.load_script_config(script_id)
        config[field] = value
        self.save_script_config(script_id, config)
        logging.info(f"Обновлено поле {field} для скрипта {script_id}")
