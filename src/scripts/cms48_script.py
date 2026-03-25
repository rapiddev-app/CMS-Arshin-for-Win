"""
CMS 48 Липецк - скрипт для ежедневной отправки данных в Аршин

Daily скрипт (аналогичен CMS36):
- Собирает данные из нескольких файлов
- Отправляет ежедневно (по рабочим дням)
- Email: cms-tab@yandex.ru
"""

import os
import sys
import warnings
import pandas as pd
from datetime import datetime, timedelta
from typing import Tuple, Any
import logging

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(__file__))))

from scripts.base_script import BaseScript
from config_manager import ConfigManager


class CMS48Script(BaseScript):
    """Скрипт CMS 48 Липецк - daily, multiple files"""

    def __init__(self):
        super().__init__(script_id='cms48', email_provider='gmail')
        self.config = ConfigManager().load_script_config('cms48')

    def gather_data(self) -> Any:
        """
        Сбор данных из нескольких Excel файлов

        Особенности:
        - Обрабатывает несколько файлов (Квадра, ЛТС, РВК)
        - Объединяет данные в один DataFrame
        - Фильтрует по дате поверки за последние 30 дней
        """
        paths = self.config['paths']

        today = datetime.today().date()
        min_valid_date = today - timedelta(days=30)

        all_data = pd.DataFrame()
        files_processed = 0

        target_files = paths.get('target_files', [
            "Реестр РирЭнерго 26",
            "Реестр Липецктеплосети 26",
            "Реестр в РВК 26"
        ])

        logging.info("CMS48: Начинаю сбор данных из файлов...")

        try:
            # Читаем все файлы из папки
            one_drive_folder = paths['input_folder']

            for file in os.listdir(one_drive_folder):
                # Проверяем, что это Excel и содержит ключевое слово
                if file.endswith(".xlsx") and any(keyword in file for keyword in target_files):
                    file_path = os.path.join(one_drive_folder, file)
                    logging.info(f"CMS48: Обрабатываю файл: {file}")

                    try:
                        # openpyxl пишет UserWarning про Data Validation в stderr — в GUI это
                        # мешало видеть реальную причину ошибки; для чтения файла предупреждение не нужно
                        with warnings.catch_warnings():
                            warnings.filterwarnings(
                                "ignore",
                                message=".*Data Validation extension.*",
                                category=UserWarning,
                            )
                            df = pd.read_excel(file_path, sheet_name=paths['sheet_name'])

                        # Фильтрация по № в ГРСИ
                        df = df[
                            df['№ в ГРСИ'].notna() &
                            (df['№ в ГРСИ'].astype(str).str.strip().str.lower() != 'нет данных')
                        ]
                        df = df.loc[:, '№ в ГРСИ':'МПИ (мес)']

                        # Проверка даты поверки
                        if 'Дата поверки' in df.columns:
                            df['Дата поверки'] = pd.to_datetime(df['Дата поверки'], errors='coerce')

                            invalid_dates = df[
                                (df['Дата поверки'] > pd.Timestamp(today)) |
                                (df['Дата поверки'] < pd.Timestamp(min_valid_date))
                            ]

                            if not invalid_dates.empty:
                                logging.warning(f"CMS48: Файл {file} содержит недопустимые даты")
                                continue

                        all_data = pd.concat([all_data, df], ignore_index=True)
                        files_processed += 1

                    except Exception as e:
                        logging.error(f"CMS48: Ошибка при обработке {file}: {e}")
                        continue

            logging.info(f"CMS48: Обработано файлов: {files_processed}, строк: {len(all_data)}")

            if all_data.empty:
                return (pd.DataFrame(), [])

            dates = all_data['Дата поверки'].dropna().tolist() if 'Дата поверки' in all_data.columns else []

            return (all_data, dates)

        except Exception as e:
            error_msg = f"<b>ОШИБКА по Липецку:</b> Не удалось обработать файлы.\n<b>Причина:</b> {e}\nПроцесс остановлен."
            logging.error(error_msg, exc_info=True)
            self.send_telegram_notification(error_msg, to_group=True)
            return (pd.DataFrame(), [])

    def create_output_file(self, data: Any) -> Tuple[str, str]:
        """Создание выходного файла"""
        df, dates = data
        paths = self.config['paths']

        count_rows = len(df)

        if dates:
            min_date = min(dates).strftime('%d.%m')
            max_date = max(dates).strftime('%d.%m.%y')
            date_range = max_date if min_date == max_date else f"{min_date} - {max_date}"
        else:
            date_range = "no_dates"

        filename = f"420 ({count_rows}) {date_range}"
        filepath = os.path.join(paths['output_folder'], f"{filename}.xlsx")

        # Проверка на дубли
        if os.path.exists(filepath):
            telegram_msg = (
                f"<b>ОШИБКА по Липецку:</b> Дубль файла <code>{filename}.xlsx</code>. "
                "Реестр не отправлен!"
            )
            logging.error(telegram_msg)
            self.send_telegram_notification(telegram_msg, to_group=True)
            raise FileExistsError(
                f"Дубль файла {filename}.xlsx: такой реестр уже есть в папке выгрузки. "
                "Удалите или переименуйте старый файл и повторите запуск."
            )

        # Форматируем даты
        if 'Дата поверки' in df.columns:
            df['Дата поверки'] = df['Дата поверки'].apply(
                lambda x: x.strftime('%d.%m.%y') if pd.notna(x) else ""
            )

        df.to_excel(filepath, index=False)
        logging.info(f"CMS48: Создан файл {filename}.xlsx ({count_rows} строк)")

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
            error_msg = f"<b>ОШИБКА по Липецку:</b> Не удалось обновить итоговый файл <code>Итоговая аршин ЦМС.xlsx</code>.\n<b>Причина:</b> {e}"
            logging.error(error_msg)
            self.send_telegram_notification(error_msg, to_group=True)

    def run(self):
        """Основной метод выполнения"""
        try:
            logging.info(f"=== Запуск скрипта CMS48 Липецк ===")

            data = self.gather_data()
            df, dates = data

            if df.empty or len(df) < 1:
                message = "CMS48: Данных о поверке по Липецку нет"
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

            filename, filepath = self.create_output_file(data)

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

            success_msg = f"Письмо с файлом {filename} успешно отправлено на {email_config['recipient']}"
            logging.info(success_msg)
            self.send_telegram_notification(success_msg, to_group=True)

            schedule = self.config['schedule']
            self.reschedule_next_run(
                schedule_type='daily',
                send_time=schedule['send_time'],
                smart_reschedule=schedule.get('smart_reschedule', True)
            )

            logging.info(f"=== Скрипт CMS48 завершён успешно ===")

        except Exception as e:
            error_message = f"<b>КРИТИЧЕСКАЯ ОШИБКА по Липецку:</b> {e}"
            logging.error(error_message, exc_info=True)
            self.send_telegram_notification(error_message, to_group=True)
            raise


if __name__ == "__main__":
    from utils import setup_logging
    setup_logging()
    script = CMS48Script()
    script.run()
