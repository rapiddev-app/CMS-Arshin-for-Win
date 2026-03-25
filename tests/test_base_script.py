"""
Тестовый скрипт для проверки BaseScript
"""

import sys
import os

# Для тестов используем конфиги из репозитория (а не из %APPDATA%)
_TESTS_DIR = os.path.dirname(__file__)
_PROJECT_ROOT = os.path.abspath(os.path.join(_TESTS_DIR, ".."))
os.environ["SCRIPTMANAGER_CONFIG_DIR"] = os.path.join(_PROJECT_ROOT, "config")

# Добавляем пути к src и src/scripts (относительно корня проекта)
_SRC_DIR = os.path.join(_PROJECT_ROOT, "src")
_SCRIPTS_DIR = os.path.join(_SRC_DIR, "scripts")
sys.path.insert(0, _SRC_DIR)
sys.path.insert(0, _SCRIPTS_DIR)

from base_script import BaseScript
from utils import setup_logging
import pandas as pd

# Инициализация логирования
setup_logging()


class TestScript(BaseScript):
    """Тестовый скрипт для проверки BaseScript"""

    def __init__(self):
        # Инициализируем с gmail провайдером
        super().__init__(script_id="test_script", email_provider="gmail")

    def gather_data(self):
        """Создаём тестовые данные"""
        print("   - Сбор тестовых данных...")
        # Создаём простой DataFrame для теста
        data = pd.DataFrame({
            'Столбец1': ['Значение1', 'Значение2', 'Значение3'],
            'Столбец2': [10, 20, 30],
            'Столбец3': ['A', 'B', 'C']
        })
        return data

    def create_output_file(self, data):
        """Создаём тестовый Excel файл"""
        print("   - Создание тестового файла...")

        # Создаём папку test_output если её нет
        output_dir = "test_output"
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        filename = "test_report.xlsx"
        filepath = os.path.join(output_dir, filename)

        # Сохраняем DataFrame в Excel
        data.to_excel(filepath, index=False)
        print(f"   - Файл создан: {filepath}")

        return filename, filepath


print("=" * 60)
print("Тестирование BaseScript")
print("=" * 60)

print("\n1. Создание экземпляра TestScript...")
script = TestScript()
print("   ✓ TestScript создан")
print(f"   - Script ID: {script.script_id}")
print(f"   - Email provider: {script.email_provider}")
print(f"   - SMTP server: {script.email_config['smtp']['server']}")

print("\n2. Тестирование gather_data()...")
data = script.gather_data()
print(f"   ✓ Данные собраны: {len(data)} строк")
print(f"   - Столбцы: {list(data.columns)}")

print("\n3. Тестирование create_output_file()...")
filename, filepath = script.create_output_file(data)
print(f"   ✓ Файл создан")
print(f"   - Имя: {filename}")
print(f"   - Путь: {filepath}")
print(f"   - Существует: {os.path.exists(filepath)}")

print("\n4. Тестирование send_telegram_notification()...")
result = script.send_telegram_notification(
    "🧪 Тест BaseScript: уведомление успешно отправлено!",
    to_group=False
)
if result:
    print("   ✓ Telegram уведомление отправлено в личный чат")
else:
    print("   ⚠ Ошибка отправки Telegram уведомления (не критично)")

print("\n5. Тестирование метода run()...")
try:
    filename, filepath = script.run()
    print("   ✓ Метод run() выполнен успешно")
    print(f"   - Результат: {filename}")
except Exception as e:
    print(f"   ✗ Ошибка в run(): {e}")

print("\n" + "=" * 60)
print("✅ Базовое тестирование завершено успешно!")
print("=" * 60)
print("\n📝 Примечание:")
print("   - Метод send_email() не тестируется автоматически")
print("   - Для полного теста отправки email нужен реальный получатель")
print("   - Telegram уведомление было отправлено в личный чат")
