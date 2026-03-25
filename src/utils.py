"""
Вспомогательные функции для ScriptManager
"""

import logging
import os
from logging.handlers import RotatingFileHandler


def setup_logging():
    """
    Настройка логирования приложения

    Создает:
    - Папку logs/ если её нет
    - Файл logs/app.log с ротацией
    - Формат: "YYYY-MM-DD HH:MM:SS - LEVEL - Message"
    - Ротация: максимум 10 МБ, последние 5 файлов
    """
    try:
        # Получаем корневую директорию проекта (на уровень выше src)
        current_dir = os.path.dirname(os.path.abspath(__file__))
        project_root = os.path.dirname(current_dir)

        # Определяем путь к логам
        import sys
        if sys.platform == 'win32':
            app_data = os.environ.get('APPDATA')
            if app_data:
                logs_dir = os.path.join(app_data, 'ScriptManager', 'logs')
            else:
                logs_dir = os.path.join(project_root, "logs")
        else:
            logs_dir = os.path.join(project_root, "logs")

        # Создаем папку logs если её нет
        if not os.path.exists(logs_dir):
            os.makedirs(logs_dir)

        # Путь к файлу логов
        log_file = os.path.join(logs_dir, "app.log")

        # Настройка формата логов
        log_format = "%(asctime)s - %(levelname)s - %(message)s"
        date_format = "%Y-%m-%d %H:%M:%S"

        # Создаем обработчик с ротацией
        # maxBytes=10*1024*1024 = 10 МБ
        # backupCount=5 = храним последние 5 файлов
        file_handler = RotatingFileHandler(
            log_file,
            maxBytes=10 * 1024 * 1024,
            backupCount=5,
            encoding='utf-8'
        )
        file_handler.setLevel(logging.INFO)
        file_handler.setFormatter(logging.Formatter(log_format, date_format))

        # Консольный вывод для отладки (с правильной кодировкой для Windows)
        import sys
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(logging.INFO)
        console_handler.setFormatter(logging.Formatter(log_format, date_format))

        # Устанавливаем UTF-8 для вывода в консоль (Windows)
        if sys.platform == 'win32':
            try:
                sys.stdout.reconfigure(encoding='utf-8')
            except (AttributeError, OSError):
                # Для старых версий Python или если reconfigure не поддерживается
                pass

        # Настройка корневого логгера
        logger = logging.getLogger()
        logger.setLevel(logging.INFO)
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)

        logging.info("Логирование инициализировано успешно")

    except Exception as e:
        # Если не удалось настроить логирование, выводим в консоль
        print(f"ОШИБКА при настройке логирования: {e}")
        raise
