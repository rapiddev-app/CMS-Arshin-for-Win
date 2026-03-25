"""
Тестовый скрипт для проверки работы календаря и умного расписания
"""

import sys
import os
from datetime import datetime, timedelta

# Добавляем путь к src
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from utils import setup_logging
from config_manager import ConfigManager
from scripts.base_script import BaseScript


class TestScript(BaseScript):
    """Тестовый скрипт для проверки календаря"""

    def gather_data(self):
        """Заглушка для сбора данных"""
        return ["test_data"]

    def create_output_file(self, data):
        """Заглушка для создания файла"""
        return "test_file.xlsx", "test_path.xlsx"


def test_is_workday():
    """Тест проверки рабочих дней"""
    print("\n=== Тест 1: Проверка рабочих дней ===")

    script = TestScript("test_script")

    # Тестируем различные даты 2025 года
    test_dates = [
        (datetime(2025, 12, 19), True, "Пятница 19.12.2025 - рабочий"),
        (datetime(2025, 12, 20), False, "Суббота 20.12.2025 - выходной"),
        (datetime(2025, 12, 21), False, "Воскресенье 21.12.2025 - выходной"),
        (datetime(2025, 12, 22), True, "Понедельник 22.12.2025 - рабочий"),
        (datetime(2025, 1, 1), False, "1 января - праздник"),
        (datetime(2025, 1, 2), False, "2 января - праздник"),
        (datetime(2025, 1, 3), False, "3 января - праздник"),
        (datetime(2025, 1, 6), False, "6 января - праздник"),
        (datetime(2025, 1, 7), True, "7 января - рабочий"),
    ]

    for test_date, expected, description in test_dates:
        result = script.is_workday(test_date)
        status = "✅" if result == expected else "❌"
        print(f"{status} {description}: {result} (ожидалось {expected})")


def test_get_next_workday():
    """Тест получения следующего рабочего дня"""
    print("\n=== Тест 2: Получение следующего рабочего дня ===")

    script = TestScript("test_script")

    # Если сегодня пятница 19.12 -> следующий рабочий - понедельник 22.12
    friday = datetime(2025, 12, 19)
    next_workday = script.get_next_workday(friday)
    expected = datetime(2025, 12, 22)
    status = "✅" if next_workday.date() == expected.date() else "❌"
    print(f"{status} С пятницы 19.12 -> {next_workday.strftime('%A %d.%m.%Y')} "
          f"(ожидалось {expected.strftime('%A %d.%m.%Y')})")

    # Если сегодня 31.12 (среда, но потом праздники) -> следующий рабочий - 09.01
    dec31 = datetime(2025, 12, 31)
    next_workday = script.get_next_workday(dec31)
    expected = datetime(2026, 1, 9)
    status = "✅" if next_workday.date() == expected.date() else "❌"
    print(f"{status} С 31.12.2025 -> {next_workday.strftime('%A %d.%m.%Y')} "
          f"(ожидалось {expected.strftime('%A %d.%m.%Y')})")


def test_get_next_run_date_daily():
    """Тест вычисления следующей даты для daily скриптов"""
    print("\n=== Тест 3: Следующая дата для daily скриптов ===")

    script = TestScript("test_script")

    # Пятница -> следующий рабочий день понедельник
    friday = datetime(2025, 12, 19, 10, 0)
    # Временно меняем текущую дату (для теста это неактуально, но показываем логику)

    next_date = script.get_next_run_date('daily', smart_reschedule=True)
    print(f"Сегодня {datetime.now().strftime('%A %d.%m.%Y')}")
    print(f"Следующий daily запуск: {next_date.strftime('%A %d.%m.%Y')}")

    # Без умного перепланирования
    next_date_no_smart = script.get_next_run_date('daily', smart_reschedule=False)
    print(f"Следующий daily запуск (без smart): {next_date_no_smart.strftime('%A %d.%m.%Y')}")


def test_get_next_run_date_weekly():
    """Тест вычисления следующей даты для weekly скриптов"""
    print("\n=== Тест 4: Следующая дата для weekly скриптов ===")

    script = TestScript("test_script")

    # Если сегодня пятница, следующий четверг
    next_thursday = script.get_next_run_date('weekly', send_day='Thursday', smart_reschedule=True)
    print(f"Сегодня {datetime.now().strftime('%A %d.%m.%Y')}")
    print(f"Следующий Thursday: {next_thursday.strftime('%A %d.%m.%Y')}")

    # Следующая пятница
    next_friday = script.get_next_run_date('weekly', send_day='Friday', smart_reschedule=True)
    print(f"Следующий Friday: {next_friday.strftime('%A %d.%m.%Y')}")


def test_calculate_collection_period():
    """Тест расчёта периода сбора данных"""
    print("\n=== Тест 5: Расчёт периода сбора данных ===")

    script = TestScript("test_script")

    # Daily скрипт (сегодня)
    start, end = script.calculate_collection_period()
    print(f"Daily скрипт: с {start.strftime('%d.%m.%Y')} по {end.strftime('%d.%m.%Y')}")

    # Weekly скрипт (четверг, последние 7 дней от четверга)
    actual_run = datetime(2025, 12, 25)  # Четверг 25.12
    start, end = script.calculate_collection_period(actual_run, base_weekday=3)
    print(f"Weekly (четверг): с {start.strftime('%d.%m.%Y')} по {end.strftime('%d.%m.%Y')}")

    # Weekly скрипт (перенесён с четверга на пятницу из-за праздника)
    actual_run = datetime(2025, 1, 10)  # Пятница 10.01 (четверг был праздник)
    start, end = script.calculate_collection_period(actual_run, base_weekday=3)
    print(f"Weekly (перенос с чт на пт): с {start.strftime('%d.%m.%Y')} по {end.strftime('%d.%m.%Y')}")


def main():
    """Запуск всех тестов"""
    print("=" * 60)
    print("Тестирование работы календаря и умного расписания")
    print("=" * 60)

    # Настраиваем логирование
    setup_logging()

    # Запускаем тесты
    test_is_workday()
    test_get_next_workday()
    test_get_next_run_date_daily()
    test_get_next_run_date_weekly()
    test_calculate_collection_period()

    print("\n" + "=" * 60)
    print("Все тесты завершены!")
    print("=" * 60)


if __name__ == "__main__":
    main()
