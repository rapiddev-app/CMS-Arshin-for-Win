"""
CMS 36 Воронеж - скрипт для ежедневной отправки данных в Аршин

Daily скрипт:
- Собирает данные за последние 30 дней
- Отправляет ежедневно (по рабочим дням)
- Email: cms-tab@yandex.ru
- Telegram: группа + личный чат
"""

import os
import sys
import pandas as pd
from datetime import datetime
from typing import Tuple, Any
import logging

# Добавляем путь к src для импорта
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(__file__))))

from scripts.base_script import BaseScript
from config_manager import ConfigManager


class CMS36Script(BaseScript):
    """
    Скрипт CMS 36 Воронеж

    Daily скрипт для отправки данных в систему Аршин ЦМС
    Собирает данные из Excel файла "Артур ЦМС.xlsx", лист "Аршин ЦМС vrn"
    """

    def __init__(self):
        super().__init__(script_id='cms36', email_provider='gmail')

        # Загружаем конфиг скрипта
        self.config = ConfigManager().load_script_config('cms36')

    def gather_data(self) -> Any:
        """
        Сбор данных из Excel файла

        Логика:
        1. Читает данные из указанного листа
        2. Фильтрует строки где № в ГРСИ не пустой
        3. Проверяет даты поверки (должны быть в диапазоне последних 30 дней)
        4. Возвращает отфильтрованный DataFrame

        Returns:
            Tuple(DataFrame, List[dates]) или (DataFrame(), []) при ошибке
        """
        paths = self.config['paths']

        today = pd.Timestamp('today').normalize()
        min_valid_date = today - pd.Timedelta(days=30)

        try:
            # Читаем Excel файл
            df = pd.read_excel(
                paths['input_file'],
                sheet_name=paths['sheet_name']
            )

            logging.info(f"Прочитан файл: {paths['input_file']}, строк: {len(df)}")

            # Фильтрация: убираем строки где № в ГРСИ пустой или "нет данных"
            df = df[
                df['№ в ГРСИ'].notna() &
                (df['№ в ГРСИ'].astype(str).str.strip().str.lower() != 'нет данных')
            ]

            # Выбираем нужные колонки (от '№ в ГРСИ' до 'МПИ (мес)')
            df = df.loc[:, '№ в ГРСИ':'МПИ (мес)']

            # Проверка даты поверки
            if 'Дата поверки' in df.columns:
                df['Дата поверки'] = pd.to_datetime(df['Дата поверки'], errors='coerce')

                # Проверяем что даты в допустимом диапазоне
                invalid_dates = df[
                    (df['Дата поверки'] > today) |
                    (df['Дата поверки'] < min_valid_date)
                ]

                if not invalid_dates.empty:
                    error_msg = "Дата поверки вне заданных пределов файл не отправлен! Воронеж"
                    logging.error(error_msg)
                    self.send_telegram_notification(error_msg, to_group=True)
                    return (pd.DataFrame(), [])

            # Итоговая проверка
            if df.empty:
                logging.warning("После фильтрации данных не осталось")
                return (pd.DataFrame(), [])

            # Собираем даты для имени файла
            dates = df['Дата поверки'].dropna().tolist()

            logging.info(f"После фильтрации осталось строк: {len(df)}, дат: {len(dates)}")

            return (df, dates)

        except FileNotFoundError:
            error_msg = f"Файл {paths['input_file']} не найден!"
            logging.error(error_msg)
            self.send_telegram_notification(error_msg, to_group=True)
            return (pd.DataFrame(), [])

        except ValueError as e:
            error_msg = f"Лист {paths['sheet_name']} не найден в файле!"
            logging.error(error_msg)
            self.send_telegram_notification(error_msg, to_group=True)
            return (pd.DataFrame(), [])

    def create_output_file(self, data: Any) -> Tuple[str, str]:
        """
        Создание выходного Excel файла

        Формат имени: "420 ({count}) {date_range} Воронеж.xlsx"
        Пример: "420 (15) 05.12 - 12.12.25 Воронеж.xlsx"

        Args:
            data: Tuple (DataFrame, List[dates])

        Returns:
            Tuple (filename, filepath)
        """
        df, dates = data
        paths = self.config['paths']

        count_rows = len(df)

        # Формируем диапазон дат
        if dates:
            min_date = min(dates).strftime('%d.%m')
            max_date = max(dates).strftime('%d.%m.%y')
            date_range = max_date if min_date == max_date else f"{min_date} - {max_date}"
        else:
            date_range = "no_dates"

        # Имя файла
        filename = f"420 ({count_rows}) {date_range} Воронеж"
        filepath = os.path.join(paths['output_folder'], f"{filename}.xlsx")

        # Проверка на существование файла (защита от дублей)
        if os.path.exists(filepath):
            error_msg = "Дубль файла реестра Воронеж, проверь! Реестр не отправлен!"
            logging.error(error_msg)
            self.send_telegram_notification(error_msg, to_group=True)
            raise FileExistsError(error_msg)

        # Форматируем даты для Excel
        if 'Дата поверки' in df.columns:
            df['Дата поверки'] = df['Дата поверки'].apply(
                lambda x: x.strftime('%d.%m.%y') if pd.notna(x) else ""
            )

        # Сохраняем файл
        df.to_excel(filepath, index=False)
        logging.info(f"Создан файл: {filename}.xlsx ({count_rows} строк)")

        return filename, filepath

    def log_to_summary_file(self, output_filename: str, count_rows: int):
        """
        Логирование в итоговый файл отчётности

        Args:
            output_filename: имя созданного файла
            count_rows: количество строк
        """
        try:
            summary_path = self.config['paths'].get('summary_file')
            if not summary_path:
                return

            # Читаем или создаём итоговый файл
            if os.path.exists(summary_path):
                df_summary = pd.read_excel(summary_path, sheet_name='Лист1')
            else:
                df_summary = pd.DataFrame(columns=['Дата отправки', 'Количество', 'Название файла'])

            # Добавляем новую запись
            new_record = pd.DataFrame([{
                'Дата отправки': datetime.today().date(),
                'Количество': count_rows,
                'Название файла': f"{output_filename}.xlsx"
            }])

            df_summary = pd.concat([df_summary, new_record], ignore_index=True)
            df_summary['Дата отправки'] = pd.to_datetime(
                df_summary['Дата отправки'], 
                dayfirst=True, 
                format='mixed'
            ).dt.strftime('%d.%m.%Y')

            # Сохраняем
            df_summary.to_excel(summary_path, sheet_name='Лист1', index=False)
            logging.info(f"Обновлён итоговый файл: {summary_path}")

        except Exception as e:
            error_msg = f"<b>ОШИБКА по Воронежу:</b> Не удалось обновить итоговый файл. Причина: {e}"
            logging.error(error_msg)
            self.send_telegram_notification(error_msg, to_group=True)

    def run(self):
        """
        Основной метод выполнения скрипта

        Последовательность:
        1. Сбор данных (gather_data)
        2. Создание файла (create_output_file)
        3. Отправка email
        4. Логирование в итоговый файл
        5. Telegram уведомление
        6. Перепланирование следующего запуска (daily)
        """
        try:
            logging.info(f"=== Запуск скрипта CMS36 Воронеж ===")

            # Сбор данных
            data = self.gather_data()
            df, dates = data

            # Проверка наличия данных
            if df.empty or len(df) < 1:
                message = "Отправка файла отменена: данных о поверке по Воронежу нет или дата вне допустимого диапазона."
                logging.warning(message)
                self.send_telegram_notification(message, to_group=True)
                # Перепланирование перед ранним выходом
                schedule = self.config['schedule']
                self.reschedule_next_run(
                    schedule_type='daily',
                    send_time=schedule['send_time'],
                    smart_reschedule=schedule.get('smart_reschedule', True)
                )
                return

            # Создание выходного файла
            filename, filepath = self.create_output_file(data)

            # Отправка email
            email_config = self.config['email']
            self.send_email(
                filename=f"{filename}.xlsx",
                filepath=filepath,
                recipient=email_config['recipient'],
                subject=email_config['subject_template'].format(filename=filename),
                body=email_config['body_template']
            )

            # Логирование в итоговый файл
            self.log_to_summary_file(filename, len(df))

            # Успешное уведомление
            success_msg = f"Письмо с файлом {filename} успешно отправлено на {email_config['recipient']}"
            logging.info(success_msg)
            self.send_telegram_notification(success_msg, to_group=True)

            # Перепланирование следующего запуска (daily скрипт)
            schedule = self.config['schedule']
            self.reschedule_next_run(
                schedule_type='daily',
                send_time=schedule['send_time'],
                smart_reschedule=schedule.get('smart_reschedule', True)
            )

            logging.info(f"=== Скрипт CMS36 завершён успешно ===")

        except Exception as e:
            error_message = f"ОШИБКА в скрипте CMS36 Воронеж: {e}"
            logging.error(error_message, exc_info=True)
            self.send_telegram_notification(error_message, to_group=True)
            raise


if __name__ == "__main__":
    # Для прямого запуска через Task Scheduler
    from utils import setup_logging
    setup_logging()

    script = CMS36Script()
    script.run()
