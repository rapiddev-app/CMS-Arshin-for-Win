"""
Kvadra - скрипт для еженедельной отправки данных в Квадра

Weekly скрипт:
- Собирает данные за последние 7 дней по "Дата вноса"
- Отправляет каждую пятницу
- Email: cms.lip@yandex.ru через Yandex SMTP
"""

import os
import sys
import pandas as pd
from datetime import datetime
from typing import Tuple, Any
import logging

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(__file__))))

from scripts.base_script import BaseScript
from config_manager import ConfigManager


class KvadraScript(BaseScript):
    """Скрипт Kvadra - weekly, использует Yandex email"""

    def __init__(self):
        super().__init__(script_id='kvadra', email_provider='yandex')
        self.config = ConfigManager().load_script_config('kvadra')

    def gather_data(self) -> Any:
        """
        Сбор данных из Excel файла

        Особенность: фильтрация по "Дата вноса" (не "Дата поверки")
        """
        paths = self.config['paths']

        # Получаем день недели из конфига
        schedule = self.config['schedule']
        send_day = schedule.get('send_day', 'Friday')
        
        weekday_map = {
            'Monday': 0, 'Tuesday': 1, 'Wednesday': 2, 'Thursday': 3,
            'Friday': 4, 'Saturday': 5, 'Sunday': 6
        }
        base_weekday = weekday_map.get(send_day, 4)

        # Расчёт периода сбора
        start_date, end_date = self.calculate_collection_period(
            actual_run_date=datetime.now(),
            base_weekday=base_weekday
        )

        try:
            df = pd.read_excel(
                paths['input_file'],
                sheet_name=paths['sheet_name']
            )

            logging.info(f"Kvadra: Прочитан файл, строк: {len(df)}")

            # Преобразуем 'Дата вноса' к дате
            df['Дата вноса'] = pd.to_datetime(df['Дата вноса'], errors='coerce').dt.date

            # Фильтрация по 'Дата вноса'
            df_filtered = df[
                (df['Дата вноса'] > start_date.date()) &
                (df['Дата вноса'] <= end_date.date())
            ]

            logging.info(f"Kvadra: Строк после фильтрации по 'Дата вноса': {len(df_filtered)}")

            # Анализ по 'Дата поверки' для логов
            if 'Дата поверки' in df.columns:
                df['Дата поверки'] = pd.to_datetime(df['Дата поверки'], errors='coerce').dt.date
                in_range = df[
                    (df['Дата поверки'] > start_date.date()) &
                    (df['Дата поверки'] <= end_date.date())
                ]
                logging.info(f"Kvadra: Строк по 'Дата поверки': {len(in_range)}")

            # Выбираем нужные столбцы
            df_filtered = df_filtered.loc[:, '№ п/п':'Номер свидетельства извещения']

            # Форматируем 'Дата поверки' для вывода
            if 'Дата поверки' in df_filtered.columns:
                df_filtered['Дата поверки'] = pd.to_datetime(
                    df_filtered['Дата поверки'], errors='coerce'
                ).dt.strftime('%d.%m.%Y')

            return df_filtered

        except Exception as e:
            error_msg = f"Ошибка при чтении файла Kvadra: {e}"
            logging.error(error_msg, exc_info=True)
            self.send_telegram_notification(error_msg, to_group=True)
            return pd.DataFrame()

    def create_output_file(self, data: Any) -> Tuple[str, str]:
        """Создание выходного файла"""
        df = data
        paths = self.config['paths']
        count_rows = len(df)

        schedule = self.config['schedule']
        send_day = schedule.get('send_day', 'Friday')
        weekday_map = {
            'Monday': 0, 'Tuesday': 1, 'Wednesday': 2, 'Thursday': 3,
            'Friday': 4, 'Saturday': 5, 'Sunday': 6
        }
        base_weekday = weekday_map.get(send_day, 4)

        start_date, end_date = self.calculate_collection_period(
            actual_run_date=datetime.now(),
            base_weekday=base_weekday
        )
        
        date_range = f"{start_date.strftime('%d.%m')} - {end_date.strftime('%d.%m.%Y')}"
        filename = f"Поверки ИПУ ЦМС ({count_rows}) {date_range}"
        filepath = os.path.join(paths['output_folder'], f"{filename}.xlsx")

        df.to_excel(filepath, index=False)

        logging.info(f"Kvadra: Создан файл {filename}.xlsx ({count_rows} строк)")
        return filename, filepath

    def run(self):
        """Основной метод выполнения"""
        try:
            logging.info(f"=== Запуск скрипта Kvadra ===")

            data = self.gather_data()
            if data.empty:
                message = "Данных о поверке ГВС-Квадра по Липецку нет, отчёт не отправлен."
                logging.warning(message)
                self.send_telegram_notification(message, to_group=True)
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

            success_msg = f"Письмо с файлом {filename} успешно отправлено на {email_config['recipient']}, перешли его срочно Леликовой ( новый адрес poverka_ipu@lipetsk.quadra.ru. )!"
            logging.info(success_msg)
            self.send_telegram_notification(success_msg, to_group=True)

            schedule = self.config['schedule']
            schedule = self.config['schedule']
            send_day = schedule.get('send_day', 'Friday')
            
            self.reschedule_next_run(
                schedule_type='weekly',
                send_time=schedule['send_time'],
                send_day=send_day,
                smart_reschedule=schedule.get('smart_reschedule', True)
            )

            logging.info(f"=== Скрипт Kvadra завершён успешно ===")

        except Exception as e:
            error_message = f"ОШИБКА в скрипте Kvadra: {e}"
            logging.error(error_message, exc_info=True)
            self.send_telegram_notification(error_message, to_group=True)
            raise


if __name__ == "__main__":
    from utils import setup_logging
    setup_logging()
    script = KvadraScript()
    script.run()
