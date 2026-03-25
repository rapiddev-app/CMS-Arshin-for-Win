"""
ScriptManager - Менеджер автоматизации скриптов для Windows

Точка входа приложения
"""

import logging
from utils import setup_logging
from gui import ScriptManagerGUI


def main():
    """Главная функция приложения"""
    try:
        # Инициализация логирования
        setup_logging()

        logging.info("ScriptManager запущен")
        logging.info("Версия: 1.0.0")

        # Запуск GUI
        app = ScriptManagerGUI()
        app.run()

    except Exception as e:
        logging.error(f"Ошибка при запуске приложения: {e}", exc_info=True)
        raise


if __name__ == "__main__":
    main()
