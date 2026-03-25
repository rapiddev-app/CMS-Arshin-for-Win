"""
Базовый класс для всех скриптов сбора данных
"""

from abc import ABC, abstractmethod
import logging
import smtplib
import requests
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from typing import Any, Tuple, Optional
from datetime import datetime, timedelta
import os
import sys

# Добавляем путь к src для импорта config_manager
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(__file__))))

from config_manager import ConfigManager


class BaseScript(ABC):
    """
    Базовый абстрактный класс для всех скриптов сбора данных

    Содержит общую логику:
    - Загрузка конфигураций (email, telegram)
    - Отправка email через SMTP
    - Отправка уведомлений в Telegram
    - Обработка ошибок

    Каждый скрипт переопределяет:
    - gather_data() - сбор данных из источника
    - create_output_file() - создание выходного файла
    """

    def __init__(self, script_id: str, email_provider: str = "gmail"):
        """
        Инициализация базового скрипта

        Args:
            script_id: идентификатор скрипта (cms36, cms48, kvadra, lts, rvk, rvk_kvadra)
            email_provider: провайдер email ("gmail" или "yandex")
        """
        self.script_id = script_id
        self.email_provider = email_provider

        # Загружаем конфигурации
        self.config_manager = ConfigManager()
        self.email_config = self.config_manager.load_email_config(email_provider)
        self.telegram_config = self.config_manager.load_telegram_config()

        logging.info(f"Инициализирован скрипт: {script_id} (email: {email_provider})")

    @abstractmethod
    def gather_data(self) -> Any:
        """
        Сбор данных из источника

        Этот метод должен быть переопределён в каждом скрипте.
        Реализует специфичную логику сбора данных (из Excel, API и т.д.)

        Returns:
            Собранные данные (обычно pandas DataFrame или None)
        """
        pass

    @abstractmethod
    def create_output_file(self, data: Any) -> Tuple[str, str]:
        """
        Создание выходного файла с данными

        Этот метод должен быть переопределён в каждом скрипте.
        Реализует специфичную логику формирования файла.

        Args:
            data: данные для сохранения

        Returns:
            Tuple (filename, filepath) - имя файла и полный путь
        """
        pass

    def send_email(self, filename: Optional[str], filepath: Optional[str], recipient: str,
                   subject: str, body: str):
        """
        Отправка email с вложенным файлом (или без него) через SMTP

        Args:
            filename: имя файла для отображения (None если без файла)
            filepath: полный путь к файлу (None если без файла)
            recipient: email получателя
            subject: тема письма
            body: тело письма
        """
        try:
            # Создаём сообщение
            msg = MIMEMultipart()
            msg['From'] = self.email_config['smtp']['username']
            # Проверяем наличие тестового получателя (через переменную окружения)
            test_recipient = os.environ.get('TEST_EMAIL_RECIPIENT')
            if test_recipient:
                logging.warning(f"РЕЖИМ ТЕСТИРОВАНИЯ: Переопределение получателя {recipient} -> {test_recipient}")
                recipient = test_recipient
                subject = f"[TEST] {subject}"

            msg['To'] = recipient
            msg['Subject'] = subject

            # Добавляем тело письма
            msg.attach(MIMEText(body, 'plain', 'utf-8'))

            # Прикрепляем файл, если он указан
            if filepath and filename and os.path.exists(filepath):
                with open(filepath, 'rb') as f:
                    attachment = MIMEApplication(f.read(), _subtype="xlsx")
                    attachment.add_header('Content-Disposition', 'attachment',
                                        filename=filename)
                    msg.attach(attachment)

            # Отправка через SMTP
            smtp_server = self.email_config['smtp']['server']
            smtp_port = self.email_config['smtp']['port']
            smtp_username = self.email_config['smtp']['username']
            smtp_password = self.email_config['smtp']['password']

            with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
                server.login(smtp_username, smtp_password)
                server.send_message(msg)

            logging.info(f"Email успешно отправлен на {recipient}: {filename}")
            return True

        except Exception as e:
            logging.error(f"Ошибка отправки email: {e}")
            raise

    def send_telegram_notification(self, message: str, to_group: bool = False):
        """
        Отправка уведомления в Telegram и дублирование на почту

        Args:
            message: текст сообщения
            to_group: отправить в группу (True) или личный чат (False)
        """
        try:
            chat_id = (self.telegram_config['group_chat_id'] if to_group
                      else self.telegram_config['personal_chat_id'])

            bot_token = self.telegram_config['bot_token']
            url = f"https://api.telegram.org/bot{bot_token}/sendMessage"

            data = {
                "chat_id": chat_id,
                "text": message,
                "parse_mode": "HTML"
            }

            response = requests.post(url, data=data, timeout=10)
            response.raise_for_status()

            logging.info(f"Telegram уведомление отправлено: {message[:50]}...")
            
            # Дублирование сообщения на почту v.daurov@hotmail.com
            try:
                self.send_email(
                    filename=None,
                    filepath=None,
                    recipient="v.daurov@hotmail.com",
                    subject=f"Уведомление от системы Аршин ({self.script_id})",
                    body=f"Копия системного уведомления:\n\n{message}"
                )
            except Exception as email_err:
                logging.error(f"Ошибка дублирования Telegram сообщения почтой: {email_err}")

            return True

        except Exception as e:
            logging.error(f"Ошибка отправки Telegram уведомления: {e}")
            # Не пробрасываем ошибку дальше - уведомление не критично
            return False

    def check_and_update_calendar(self) -> bool:
        """
        Проверка актуальности календаря и автоматическое обновление при необходимости

        Проверяет:
        1. Не устарел ли календарь (current_year > calendar_year)
        2. Не скоро ли закончится год (осталось меньше 30 дней до нового года)

        Если нужно обновление - пытается загрузить новый календарь через API

        Returns:
            True если календарь актуален или обновлён успешно, False если ошибка
        """
        try:
            calendar_data = self.config_manager.load_calendar()
            calendar_year = calendar_data.get('year', 2025)
            current_year = datetime.now().year

            # Проверяем что календарь не устарел
            if current_year > calendar_year:
                logging.warning(f"Календарь устарел! Текущий год: {current_year}, в календаре: {calendar_year}")
                logging.info(f"Попытка загрузить календарь на {current_year} год...")
                return self._update_calendar_for_year(current_year)

            # Проверяем что до конца года осталось больше 30 дней
            days_until_new_year = (datetime(current_year + 1, 1, 1) - datetime.now()).days
            if days_until_new_year <= 30 and calendar_year == current_year:
                logging.info(f"До конца года осталось {days_until_new_year} дней. Загружаю календарь на {current_year + 1} год...")
                return self._update_calendar_for_year(current_year + 1)

            return True

        except Exception as e:
            logging.error(f"Ошибка при проверке календаря: {e}")
            return False

    def _update_calendar_for_year(self, year: int) -> bool:
        """
        Загрузка нового производственного календаря через API

        Args:
            year: год для загрузки календаря

        Returns:
            True если успешно загружен и сохранён, False при ошибке
        """
        try:
            import requests

            # API для получения производственного календаря РФ
            api_url = f"https://isdayoff.ru/api/getdata?year={year}&cc=ru"

            logging.info(f"Загрузка календаря с {api_url}...")
            response = requests.get(api_url, timeout=10)
            response.raise_for_status()

            days_string = response.text.strip()

            # Проверяем что получили 365 или 366 символов
            expected_days = 366 if (year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)) else 365
            if len(days_string) != expected_days:
                raise ValueError(f"Получено {len(days_string)} дней, ожидалось {expected_days}")

            # Формируем новый календарь
            new_calendar = {
                "year": year,
                "holidays": [],  # Можно заполнить позже вручную
                "days": days_string,
                "last_update": datetime.now().strftime("%Y-%m-%d")
            }

            # Сохраняем
            calendar_path = os.path.join(
                os.path.dirname(os.path.dirname(os.path.dirname(__file__))),
                "config", "calendar.json"
            )

            with open(calendar_path, 'w', encoding='utf-8') as f:
                import json
                json.dump(new_calendar, f, ensure_ascii=False, indent=2)

            logging.info(f"Календарь на {year} год успешно загружен и сохранён!")

            # Отправляем уведомление в Telegram
            self.send_telegram_notification(
                f"📅 Календарь обновлён!\n\nЗагружен производственный календарь РФ на {year} год.\nАвтоматическое обновление прошло успешно.",
                to_group=True
            )

            return True

        except Exception as e:
            error_msg = f"Ошибка при загрузке календаря на {year} год: {e}"
            logging.error(error_msg)
            self.send_telegram_notification(
                f"⚠️ Ошибка обновления календаря!\n\n{error_msg}\n\nТребуется ручное обновление config/calendar.json",
                to_group=True
            )
            return False

    def is_workday(self, date: datetime) -> bool:
        """
        Проверка, является ли дата рабочим днём

        Args:
            date: дата для проверки

        Returns:
            True если рабочий день, False если выходной или праздник
        """
        # Проверяем актуальность календаря перед использованием
        self.check_and_update_calendar()

        calendar_data = self.config_manager.load_calendar()
        year = calendar_data['year']
        days_string = calendar_data['days']

        # Проверяем что дата в том же году что и календарь
        if date.year != year:
            logging.warning(f"Календарь загружен для {year} года, а проверяется {date.year}")
            # Простая проверка на выходные (суббота=5, воскресенье=6)
            return date.weekday() < 5

        # Определяем день года (1 января = индекс 0)
        day_of_year = (date - datetime(date.year, 1, 1)).days

        if day_of_year >= len(days_string):
            return True  # За пределами календаря считаем рабочим днём

        # '0' = рабочий день, '1' = выходной/праздник
        return days_string[day_of_year] == '0'

    def get_next_workday(self, date: datetime) -> datetime:
        """
        Получить следующий рабочий день после заданной даты

        Args:
            date: начальная дата

        Returns:
            Ближайший следующий рабочий день
        """
        current = date
        max_iterations = 10  # Защита от бесконечного цикла

        for _ in range(max_iterations):
            current += timedelta(days=1)
            if self.is_workday(current):
                return current

        # Если не нашли за 10 дней, возвращаем как есть
        logging.warning(f"Не найден рабочий день в течение 10 дней после {date}")
        return current

    def calculate_collection_period(self, actual_run_date: Optional[datetime] = None,
                                    base_weekday: Optional[int] = None) -> Tuple[datetime, datetime]:
        """
        Расчёт периода сбора данных для скрипта

        Для daily скриптов:
            start_date = end_date = actual_run_date (или сегодня)

        Для weekly скриптов:
            start_date = предыдущий base_weekday (фиксированный)
            end_date = actual_run_date (может быть перенесён на следующий рабочий)

        Args:
            actual_run_date: фактическая дата запуска (по умолчанию сегодня)
            base_weekday: базовый день недели для weekly (0=пн, 3=чт)

        Returns:
            Tuple (start_date, end_date)

        Примеры:
            Daily:
                calculate_collection_period() → (сегодня, сегодня)

            Weekly (четверг запланирован, четверг рабочий):
                calculate_collection_period(datetime(2025,1,14), 3)
                → (datetime(2025,1,7), datetime(2025,1,14))

            Weekly (четверг был выходной, перенесено на пятницу):
                calculate_collection_period(datetime(2025,1,15), 3)
                → (datetime(2025,1,7), datetime(2025,1,15))
        """
        if actual_run_date is None:
            actual_run_date = datetime.now()

        # Для daily скриптов
        if base_weekday is None:
            return actual_run_date, actual_run_date

        # Для weekly скриптов
        end_date = actual_run_date

        # Находим предыдущий base_weekday от фактической даты запуска
        days_since_base = (actual_run_date.weekday() - base_weekday) % 7

        if days_since_base == 0:
            # Если сегодня сам базовый день (например четверг), берём прошлый четверг
            days_back = 7
        else:
            # Иначе берём последний прошедший базовый день
            days_back = days_since_base

        start_date = actual_run_date - timedelta(days=days_back)

        logging.info(f"Период сбора: с {start_date.strftime('%Y-%m-%d')} "
                    f"по {end_date.strftime('%Y-%m-%d')} ({days_back} дней)")

        return start_date, end_date

    def get_next_run_date(self, schedule_type: str, send_day: Optional[str] = None,
                         smart_reschedule: bool = True) -> datetime:
        """
        Вычисление следующей даты запуска скрипта

        Args:
            schedule_type: тип расписания ('daily' или 'weekly')
            send_day: день недели для weekly ('Monday', 'Tuesday', и т.д.)
            smart_reschedule: учитывать производственный календарь

        Returns:
            Дата и время следующего запуска

        Примеры:
            Daily скрипт (сегодня понедельник, рабочий):
                get_next_run_date('daily') → завтра (вторник)

            Daily скрипт (сегодня пятница):
                get_next_run_date('daily', smart_reschedule=True) → понедельник

            Weekly скрипт (send_day='Thursday', сегодня четверг):
                get_next_run_date('weekly', 'Thursday') → следующий четверг
        """
        today = datetime.now()

        if schedule_type == 'daily':
            # Для ежедневных скриптов: следующий день
            next_date = today + timedelta(days=1)

            # Если включено умное перепланирование - переносим на рабочий день
            if smart_reschedule:
                while not self.is_workday(next_date):
                    next_date += timedelta(days=1)

        elif schedule_type == 'weekly':
            if not send_day:
                raise ValueError("Для weekly расписания требуется указать send_day")

            # Преобразуем название дня в номер (0=Monday, 6=Sunday)
            weekday_map = {
                'Monday': 0, 'Tuesday': 1, 'Wednesday': 2, 'Thursday': 3,
                'Friday': 4, 'Saturday': 5, 'Sunday': 6
            }
            target_weekday = weekday_map.get(send_day)
            if target_weekday is None:
                raise ValueError(f"Некорректный день недели: {send_day}")

            # Вычисляем сколько дней до следующего target_weekday
            current_weekday = today.weekday()
            days_ahead = (target_weekday - current_weekday) % 7

            if days_ahead == 0:
                # Если сегодня сам целевой день - берём следующую неделю
                days_ahead = 7

            next_date = today + timedelta(days=days_ahead)

            # Если включено умное перепланирование - переносим на рабочий день
            if smart_reschedule:
                while not self.is_workday(next_date):
                    next_date += timedelta(days=1)

        else:
            raise ValueError(f"Неизвестный тип расписания: {schedule_type}")

        logging.info(f"Следующий запуск {self.script_id}: {next_date.strftime('%Y-%m-%d')}")
        return next_date

    def reschedule_next_run(self, schedule_type: str, send_time: str,
                           send_day: Optional[str] = None, smart_reschedule: bool = True):
        """
        Пересчёт и обновление следующего запуска в Task Scheduler

        Вызывается после успешной отправки письма для планирования
        следующего запуска с учётом производственного календаря.

        Args:
            schedule_type: тип расписания ('daily' или 'weekly')
            send_time: время отправки в формате 'HH:MM' (например, '09:00')
            send_day: день недели для weekly ('Monday', 'Tuesday', и т.д.)
            smart_reschedule: учитывать производственный календарь

        Примеры:
            Daily скрипт:
                reschedule_next_run('daily', '10:00')
                → планирует на завтра 10:00 (или следующий рабочий)

            Weekly скрипт:
                reschedule_next_run('weekly', '09:00', 'Thursday')
                → планирует на следующий четверг 09:00 (или следующий рабочий)
        """
        try:
            # Вычисляем следующую дату запуска
            next_date = self.get_next_run_date(schedule_type, send_day, smart_reschedule)

            # Добавляем время
            time_parts = send_time.split(':')
            next_run = next_date.replace(
                hour=int(time_parts[0]),
                minute=int(time_parts[1]),
                second=0,
                microsecond=0
            )

            # Обновляем задачу в Task Scheduler
            from scheduler import TaskSchedulerManager
            scheduler = TaskSchedulerManager()
            success = scheduler.update_task_date(self.script_id, next_run)

            if success:
                logging.info(f"Следующий запуск обновлён: {next_run.strftime('%Y-%m-%d %H:%M')}")
            else:
                logging.error(f"Не удалось обновить задачу в Task Scheduler")

        except Exception as e:
            logging.error(f"Ошибка при пересчёте следующего запуска: {e}")
            # Не пробрасываем ошибку дальше - это не критично

    def run(self):
        """
        Основной метод выполнения скрипта

        Вызывает последовательно:
        1. gather_data() - сбор данных
        2. create_output_file() - создание файла
        3. Отправка результата (реализуется в наследниках)

        Обрабатывает ошибки и отправляет уведомления в Telegram
        """
        try:
            logging.info(f"Запуск скрипта {self.script_id}...")

            # Сбор данных
            data = self.gather_data()

            # Проверка наличия данных
            if data is None or (hasattr(data, '__len__') and len(data) == 0):
                message = f"Нет данных для отправки ({self.script_id})"
                logging.warning(message)
                self.send_telegram_notification(message, to_group=True)
                return

            # Создание выходного файла
            filename, filepath = self.create_output_file(data)

            logging.info(f"Скрипт {self.script_id} успешно выполнен: {filename}")

            # Возвращаем данные для дальнейшей обработки в наследниках
            return filename, filepath

        except Exception as e:
            error_message = f"ОШИБКА в скрипте {self.script_id}: {e}"
            logging.error(error_message, exc_info=True)
            self.send_telegram_notification(error_message, to_group=True)
            raise
