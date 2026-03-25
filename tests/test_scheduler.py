"""
Тестовый скрипт для проверки TaskSchedulerManager
"""

import sys
import os
sys.path.insert(0, 'src')

from scheduler import TaskSchedulerManager
from utils import setup_logging
from datetime import datetime, timedelta

# Инициализация логирования
setup_logging()

print("=" * 60)
print("Тестирование TaskSchedulerManager")
print("=" * 60)

print("\n1. Создание экземпляра TaskSchedulerManager...")
try:
    manager = TaskSchedulerManager()
    print("   ✓ TaskSchedulerManager создан")
except Exception as e:
    print(f"   ✗ Ошибка: {e}")
    sys.exit(1)

# Тестовые данные
test_script_id = "test_task"
test_script_path = os.path.abspath("test_base_script.py")
next_run = datetime.now() + timedelta(days=1, hours=9)  # Завтра в 9:00

print("\n2. Создание тестовой задачи...")
print(f"   - Script ID: {test_script_id}")
print(f"   - Script path: {test_script_path}")
print(f"   - Next run: {next_run.strftime('%Y-%m-%d %H:%M')}")

success = manager.create_task(
    script_id=test_script_id,
    script_path=test_script_path,
    next_run_time=next_run,
    description="Тестовая задача для проверки TaskSchedulerManager"
)

if success:
    print("   ✓ Задача создана успешно")
else:
    print("   ✗ Ошибка создания задачи")
    sys.exit(1)

print("\n3. Получение статуса задачи...")
status = manager.get_task_status(test_script_id)
if status:
    print("   ✓ Статус получен:")
    print(f"      - Enabled: {status['enabled']}")
    print(f"      - Next run: {status['next_run']}")
    print(f"      - Last run: {status['last_run']}")
    print(f"      - Last result: {status['last_result']}")
else:
    print("   ⚠ Не удалось получить статус")

print("\n4. Обновление времени запуска...")
new_run_time = datetime.now() + timedelta(days=2, hours=10)  # Послезавтра в 10:00
print(f"   - Новое время: {new_run_time.strftime('%Y-%m-%d %H:%M')}")

success = manager.update_task_date(test_script_id, new_run_time)
if success:
    print("   ✓ Время обновлено")

    # Проверяем обновление
    status = manager.get_task_status(test_script_id)
    if status:
        print(f"   - Проверка: {status['next_run']}")
else:
    print("   ✗ Ошибка обновления")

print("\n5. Тестирование disable/enable...")
print("   - Отключение задачи...")
if manager.disable_task(test_script_id):
    print("   ✓ Задача отключена")
    status = manager.get_task_status(test_script_id)
    if status:
        print(f"      Enabled: {status['enabled']}")
else:
    print("   ✗ Ошибка отключения")

print("   - Включение задачи...")
if manager.enable_task(test_script_id):
    print("   ✓ Задача включена")
    status = manager.get_task_status(test_script_id)
    if status:
        print(f"      Enabled: {status['enabled']}")
else:
    print("   ✗ Ошибка включения")

print("\n6. Получение всех задач ScriptManager...")
all_tasks = manager.get_all_tasks()
print(f"   ✓ Найдено задач: {len(all_tasks)}")
for script_id, task_status in all_tasks.items():
    print(f"   - {script_id}:")
    print(f"      Enabled: {task_status['enabled']}")
    print(f"      Next run: {task_status['next_run']}")

print("\n7. Удаление тестовой задачи...")
success = manager.delete_task(test_script_id)
if success:
    print("   ✓ Задача удалена")

    # Проверяем удаление
    status = manager.get_task_status(test_script_id)
    if status is None:
        print("   ✓ Задача действительно удалена")
    else:
        print("   ⚠ Задача всё ещё существует")
else:
    print("   ✗ Ошибка удаления")

print("\n" + "=" * 60)
print("✅ Тестирование TaskSchedulerManager завершено!")
print("=" * 60)
print("\n📝 Примечание:")
print("   - Проверьте Task Scheduler вручную (taskschd.msc)")
print("   - Убедитесь, что тестовая задача была создана и удалена")
