"""
Менеджер задач Windows Task Scheduler
"""

import logging
from datetime import datetime
from typing import Optional, Dict, Any
import win32com.client
import pythoncom


class TaskSchedulerManager:
    """
    Класс для управления задачами в Windows Task Scheduler

    Использует pywin32 (win32com.client) для взаимодействия с планировщиком.
    Все задачи ScriptManager имеют префикс "ScriptManager_"
    """

    TASK_PREFIX = "ScriptManager_"
    FOLDER_PATH = "\\"  # Корневая папка планировщика

    def __init__(self):
        """Инициализация менеджера задач"""
        try:
            # Инициализируем COM (необходимо для pywin32)
            pythoncom.CoInitialize()

            # Подключаемся к планировщику задач
            self.scheduler = win32com.client.Dispatch('Schedule.Service')
            self.scheduler.Connect()

            # Получаем корневую папку
            self.root_folder = self.scheduler.GetFolder(self.FOLDER_PATH)

            logging.info("TaskSchedulerManager инициализирован")

        except Exception as e:
            logging.error(f"Ошибка инициализации TaskSchedulerManager: {e}")
            raise

    def _get_task_name(self, script_id: str) -> str:
        """
        Формирование полного имени задачи

        Args:
            script_id: идентификатор скрипта (cms36, cms48, kvadra, lts, rvk, rvk_kvadra)

        Returns:
            Полное имя задачи с префиксом
        """
        return f"{self.TASK_PREFIX}{script_id}"

    def create_task(self, script_id: str, script_path: str, script_args: str,
                   next_run_time: datetime, description: str = "",
                   python_exe: str = "python.exe", enabled: bool = True) -> bool:
        """
        Создание новой задачи в планировщике

        Args:
            script_id: идентификатор скрипта
            script_path: полный путь к Python скрипту
            next_run_time: дата и время следующего запуска
            description: описание задачи (опционально)
            python_exe: путь к интерпретатору Python (по умолчанию "python.exe")
            enabled: должна ли задача быть активна

        Returns:
            True если задача создана успешно, иначе False
        """
        try:
            task_name = self._get_task_name(script_id)

            # Создаём новую задачу
            task_definition = self.scheduler.NewTask(0)

            # Настройки задачи
            task_definition.RegistrationInfo.Description = description or f"Автоматический запуск скрипта {script_id}"
            task_definition.RegistrationInfo.Author = "ScriptManager"

            # Параметры выполнения
            task_definition.Settings.Enabled = enabled
            task_definition.Settings.Hidden = False
            task_definition.Settings.StartWhenAvailable = True
            task_definition.Settings.DisallowStartIfOnBatteries = False
            task_definition.Settings.StopIfGoingOnBatteries = False
            task_definition.Settings.ExecutionTimeLimit = "PT2H"  # Максимум 2 часа

            # Создаём триггер (однократный запуск)
            trigger = task_definition.Triggers.Create(1)  # 1 = TASK_TRIGGER_TIME (однократно)
            trigger.StartBoundary = next_run_time.strftime("%Y-%m-%dT%H:%M:%S")
            trigger.Enabled = True

            # Создаём действие (запуск Python скрипта)
            action = task_definition.Actions.Create(0)  # 0 = TASK_ACTION_EXEC
            action.Path = python_exe
            action.Arguments = f'"{script_path}" {script_args}'

            # Регистрируем задачу
            self.root_folder.RegisterTaskDefinition(
                task_name,
                task_definition,
                6,  # TASK_CREATE_OR_UPDATE
                None,  # user
                None,  # password
                3,  # TASK_LOGON_INTERACTIVE_TOKEN
                None   # sddl
            )

            logging.info(f"Задача создана: {task_name} -> {next_run_time.strftime('%Y-%m-%d %H:%M')}")
            return True

        except Exception as e:
            logging.error(f"Ошибка создания задачи {script_id}: {e}")
            return False

    def update_task_date(self, script_id: str, next_run_time: datetime) -> bool:
        """
        Обновление даты следующего запуска задачи

        Args:
            script_id: идентификатор скрипта
            next_run_time: новая дата и время следующего запуска

        Returns:
            True если задача обновлена успешно, иначе False
        """
        try:
            task_name = self._get_task_name(script_id)

            # Получаем существующую задачу
            task = self.root_folder.GetTask(task_name)
            task_definition = task.Definition

            # Обновляем триггер
            if task_definition.Triggers.Count > 0:
                trigger = task_definition.Triggers.Item(1)  # Первый триггер
                trigger.StartBoundary = next_run_time.strftime("%Y-%m-%dT%H:%M:%S")

                # Перерегистрируем задачу
                self.root_folder.RegisterTaskDefinition(
                    task_name,
                    task_definition,
                    6,  # TASK_CREATE_OR_UPDATE
                    None,
                    None,
                    3,
                    None
                )

                logging.info(f"Задача обновлена: {task_name} -> {next_run_time.strftime('%Y-%m-%d %H:%M')}")
                return True
            else:
                logging.warning(f"У задачи {task_name} нет триггеров")
                return False

        except Exception as e:
            logging.error(f"Ошибка обновления задачи {script_id}: {e}")
            return False

    def delete_task(self, script_id: str) -> bool:
        """
        Удаление задачи из планировщика

        Args:
            script_id: идентификатор скрипта

        Returns:
            True если задача удалена успешно, иначе False
        """
        try:
            task_name = self._get_task_name(script_id)
            self.root_folder.DeleteTask(task_name, 0)

            logging.info(f"Задача удалена: {task_name}")
            return True

        except Exception as e:
            logging.error(f"Ошибка удаления задачи {script_id}: {e}")
            return False

    def get_task_status(self, script_id: str) -> Optional[Dict[str, Any]]:
        """
        Получение информации о задаче

        Args:
            script_id: идентификатор скрипта

        Returns:
            Словарь с информацией о задаче:
            {
                'enabled': bool,
                'next_run': datetime or None,
                'last_run': datetime or None,
                'last_result': int (0 = успех, другое = код ошибки)
            }
            Или None если задача не найдена
        """
        try:
            task_name = self._get_task_name(script_id)
            task = self.root_folder.GetTask(task_name)

            # Получаем информацию
            enabled = task.Enabled
            last_run_time = task.LastRunTime
            last_task_result = task.LastTaskResult
            next_run_time = task.NextRunTime

            # Преобразуем COM даты в datetime
            def com_time_to_datetime(com_time):
                """Преобразование COM времени в datetime"""
                try:
                    if com_time and str(com_time) != "1899-12-30 00:00:00":
                        return datetime.fromisoformat(str(com_time))
                    return None
                except:
                    return None

            return {
                'enabled': enabled,
                'next_run': com_time_to_datetime(next_run_time),
                'last_run': com_time_to_datetime(last_run_time),
                'last_result': last_task_result
            }

        except Exception as e:
            logging.error(f"Ошибка получения статуса задачи {script_id}: {e}")
            return None

    def enable_task(self, script_id: str) -> bool:
        """
        Включение задачи

        Args:
            script_id: идентификатор скрипта

        Returns:
            True если задача включена успешно, иначе False
        """
        try:
            task_name = self._get_task_name(script_id)
            task = self.root_folder.GetTask(task_name)
            task_definition = task.Definition

            task_definition.Settings.Enabled = True

            self.root_folder.RegisterTaskDefinition(
                task_name,
                task_definition,
                6,
                None,
                None,
                3,
                None
            )

            logging.info(f"Задача включена: {task_name}")
            return True

        except Exception as e:
            logging.error(f"Ошибка включения задачи {script_id}: {e}")
            return False

    def disable_task(self, script_id: str) -> bool:
        """
        Отключение задачи

        Args:
            script_id: идентификатор скрипта

        Returns:
            True если задача отключена успешно, иначе False
        """
        try:
            task_name = self._get_task_name(script_id)
            task = self.root_folder.GetTask(task_name)
            task_definition = task.Definition

            task_definition.Settings.Enabled = False

            self.root_folder.RegisterTaskDefinition(
                task_name,
                task_definition,
                6,
                None,
                None,
                3,
                None
            )

            logging.info(f"Задача отключена: {task_name}")
            return True

        except Exception as e:
            logging.error(f"Ошибка отключения задачи {script_id}: {e}")
            return False

    def get_all_tasks(self) -> Dict[str, Dict[str, Any]]:
        """
        Получение списка всех задач ScriptManager

        Returns:
            Словарь {script_id: task_status}
        """
        all_tasks = {}

        try:
            # Получаем все задачи из папки
            tasks = self.root_folder.GetTasks(0)

            for task in tasks:
                task_name = task.Name

                # Фильтруем только задачи ScriptManager
                if task_name.startswith(self.TASK_PREFIX):
                    script_id = task_name[len(self.TASK_PREFIX):]
                    status = self.get_task_status(script_id)
                    if status:
                        all_tasks[script_id] = status

            logging.info(f"Найдено задач ScriptManager: {len(all_tasks)}")
            return all_tasks

        except Exception as e:
            logging.error(f"Ошибка получения списка задач: {e}")
            return {}

    def __del__(self):
        """Деструктор - освобождаем COM ресурсы"""
        try:
            pythoncom.CoUninitialize()
        except:
            pass
