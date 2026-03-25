"""
Тестовый скрипт для проверки ConfigManager
"""

import sys
sys.path.insert(0, 'src')

from config_manager import ConfigManager
from utils import setup_logging

# Инициализация логирования
setup_logging()

print("=" * 60)
print("Тестирование ConfigManager")
print("=" * 60)

# Создание экземпляра ConfigManager
print("\n1. Создание ConfigManager...")
cm = ConfigManager()
print("   ✓ ConfigManager создан")

# Проверка settings.json
print("\n2. Загрузка settings.json...")
settings = cm.load_settings()
print(f"   ✓ Settings загружены: {settings}")

# Проверка email_gmail.json
print("\n3. Загрузка email_gmail.json...")
gmail_config = cm.load_email_config("gmail")
print(f"   ✓ Gmail конфиг загружен")
print(f"     - Server: {gmail_config['smtp']['server']}")
print(f"     - Port: {gmail_config['smtp']['port']}")
print(f"     - Username: {gmail_config['smtp']['username']}")

# Проверка email_yandex.json
print("\n4. Загрузка email_yandex.json...")
yandex_config = cm.load_email_config("yandex")
print(f"   ✓ Yandex конфиг загружен")
print(f"     - Server: {yandex_config['smtp']['server']}")
print(f"     - Port: {yandex_config['smtp']['port']}")
print(f"     - Username: {yandex_config['smtp']['username']}")

# Проверка telegram.json
print("\n5. Загрузка telegram.json...")
telegram_config = cm.load_telegram_config()
print(f"   ✓ Telegram конфиг загружен")
print(f"     - Bot token начинается с: {telegram_config['bot_token'][:20]}...")
print(f"     - Personal chat ID: {telegram_config['personal_chat_id']}")
print(f"     - Group chat ID: {telegram_config['group_chat_id']}")

# Проверка calendar.json
print("\n6. Загрузка calendar.json...")
calendar = cm.load_calendar()
print(f"   ✓ Календарь загружен")
print(f"     - Год: {calendar['year']}")
print(f"     - Последнее обновление: {calendar.get('last_update', 'N/A')}")
print(f"     - Источник: {calendar.get('source', 'API')}")
print(f"     - Дней в календаре: {len(calendar.get('days', ''))}")

print("\n" + "=" * 60)
print("✅ Все тесты пройдены успешно!")
print("=" * 60)
