"""
Runner - точка входа для запуска скриптов через Windows Task Scheduler

Использование:
    python runner.py <script_id>

Примеры:
    python runner.py cms36
    python runner.py rvk
    python runner.py kvadra

Этот файл импортирует нужный скрипт и запускает его метод run()
"""

import sys
import os
import logging

# Добавляем путь к src
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(__file__))))

from utils import setup_logging


# Импорты всех скриптов
from scripts.cms36_script import CMS36Script
from scripts.cms48_script import CMS48Script
from scripts.kvadra_script import KvadraScript
from scripts.lts_script import LTSScript
from scripts.rvk_script import RVKScript
from scripts.rvk_kvadra_script import RVKKvadraScript


# Маппинг script_id -> класс скрипта
SCRIPTS_MAP = {
    'cms36': CMS36Script,
    'cms48': CMS48Script,
    'kvadra': KvadraScript,
    'lts': LTSScript,
    'rvk': RVKScript,
    'rvk_kvadra': RVKKvadraScript
}


def main():
    """
    Главная функция runner

    Читает script_id из аргументов командной строки,
    создаёт экземпляр скрипта и запускает его
    """
    # Настраиваем логирование
    setup_logging()

    # Проверяем аргументы
    if len(sys.argv) < 2:
        error_msg = "Использование: python runner.py <script_id>"
        logging.error(error_msg)
        print(error_msg)
        print("\nДоступные script_id:")
        for script_id in SCRIPTS_MAP.keys():
            print(f"  - {script_id}")
        sys.exit(1)

    # Получаем script_id
    script_id = sys.argv[1].lower()

    # Проверяем что такой скрипт существует
    if script_id not in SCRIPTS_MAP:
        error_msg = f"Неизвестный скрипт: {script_id}"
        logging.error(error_msg)
        print(error_msg)
        print("\nДоступные script_id:")
        for sid in SCRIPTS_MAP.keys():
            print(f"  - {sid}")
        sys.exit(1)

    # Создаём и запускаем скрипт
    try:
        logging.info(f"=" * 60)
        logging.info(f"RUNNER: Запуск скрипта '{script_id}'")
        logging.info(f"=" * 60)

        script_class = SCRIPTS_MAP[script_id]
        script = script_class()
        script.run()

        logging.info(f"=" * 60)
        logging.info(f"RUNNER: Скрипт '{script_id}' завершён успешно")
        logging.info(f"=" * 60)

        sys.exit(0)

    except Exception as e:
        error_msg = f"RUNNER: Критическая ошибка при выполнении скрипта '{script_id}': {e}"
        logging.error(error_msg, exc_info=True)
        # stdout и stderr: GUI объединяет оба потока; дублируем в stderr для инструментов,
        # которые читают только один из потоков.
        print(error_msg, file=sys.stderr)
        print(error_msg)
        sys.exit(1)


if __name__ == "__main__":
    main()
