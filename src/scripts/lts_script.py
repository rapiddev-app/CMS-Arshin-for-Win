"""
LTS (Липецктеплосеть) - скрипт для еженедельной отправки данных

Weekly скрипт (аналогичен RVK):
- Собирает данные за последние 7 дней
- Отправляет каждый четверг
- Email: mup-sbyt@mail.ru
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


class LTSScript(BaseScript):
    """Скрипт LTS (Липецктеплосеть) - weekly"""

    def __init__(self):
        super().__init__(script_id='lts', email_provider='gmail')
        self.config = ConfigManager().load_script_config('lts')

    def gather_data(self) -> Any:
        """Сбор данных (аналогично RVK)"""
        # Получаем день недели из конфига
        schedule = self.config['schedule']
        paths = self.config['paths']
        send_day = schedule.get('send_day', 'Thursday')
        
        weekday_map = {
            'Monday': 0, 'Tuesday': 1, 'Wednesday': 2, 'Thursday': 3,
            'Friday': 4, 'Saturday': 5, 'Sunday': 6
        }
        base_weekday = weekday_map.get(send_day, 3)

        start_date, end_date = self.calculate_collection_period(
            actual_run_date=datetime.now(),
            base_weekday=base_weekday
        )

        try:
            df = pd.read_excel(
                paths['input_file'],
                sheet_name=paths['sheet_name'],
                header=[0, 1],
                dtype={'Unnamed: 1_level_0_Лицевой счет': str}
            )

            if ('Unnamed: 1_level_0', 'Лицевой счет') in df.columns:
                df[('Unnamed: 1_level_0', 'Лицевой счет')] = df[(
                    'Unnamed: 1_level_0', 'Лицевой счет'
                )].apply(lambda x: str(int(float(x))) if pd.notna(x) else '')

            date_columns = [
                'ИПУ 1_Дата выпуска счетчика', 'ИПУ 2_Дата выпуска счетчика',
                'ИПУ 3_Дата выпуска счетчика', 'ИПУ 4_Дата выпуска счетчика',
                'ИПУ 1_Дата поверки', 'ИПУ 2_Дата поверки',
                'ИПУ 3_Дата поверки', 'ИПУ 4_Дата поверки',
                'ИПУ 1_Дата очередной поверки', 'ИПУ 2_Дата очередной поверки',
                'ИПУ 3_Дата очередной поверки', 'ИПУ 4_Дата очередной поверки'
            ]

            for col in date_columns:
                if col in df.columns.get_level_values(1):
                    df[col] = pd.to_datetime(df[col], errors='coerce')

            df.columns = ['_'.join(col).strip() for col in df.columns.values]

            date_column = 'ИПУ 1_Дата поверки'
            if date_column in df.columns:
                df_filtered = df.dropna(subset=[date_column]).copy()
                df_filtered.loc[:, 'Дата поверки'] = pd.to_datetime(
                    df_filtered[date_column], errors='coerce'
                )
            else:
                raise KeyError(f"Колонка '{date_column}' не найдена")

            df_filtered = df_filtered[
                (df_filtered['Дата поверки'] >= start_date) &
                (df_filtered['Дата поверки'] <= end_date)
            ]

            df_filtered = df_filtered.iloc[:, :43]

            for col in date_columns:
                if col in df_filtered.columns:
                    df_filtered[col] = pd.to_datetime(
                        df_filtered[col], errors='coerce'
                    ).dt.strftime('%d.%m.%Y')

            logging.info(f"LTS: После фильтрации осталось строк: {len(df_filtered)}")
            return df_filtered

        except Exception as e:
            error_msg = f"Ошибка при чтении файла LTS: {e}"
            logging.error(error_msg, exc_info=True)
            self.send_telegram_notification(error_msg, to_group=True)
            return pd.DataFrame()

    def create_output_file(self, data: Any) -> Tuple[str, str]:
        """Создание выходного файла"""
        df = data
        paths = self.config['paths']
        count_rows = len(df)

        schedule = self.config['schedule']
        send_day = schedule.get('send_day', 'Thursday')
        weekday_map = {
            'Monday': 0, 'Tuesday': 1, 'Wednesday': 2, 'Thursday': 3,
            'Friday': 4, 'Saturday': 5, 'Sunday': 6
        }
        base_weekday = weekday_map.get(send_day, 3)

        start_date, end_date = self.calculate_collection_period(
            actual_run_date=datetime.now(),
            base_weekday=base_weekday
        )

        date_range = f"{start_date.strftime('%d.%m')} - {end_date.strftime('%d.%m.%Y')}"
        filename = f"Реестр поверок ИПУ ({count_rows}) {date_range} (Липецктеплосеть) ЦМС"
        filepath = os.path.join(paths['output_folder'], f"{filename}.xlsx")

        if os.path.exists(filepath):
            try:
                os.remove(filepath)
            except Exception as e:
                logging.warning(f"Не удалось удалить файл: {e}")

        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, header=True)

        logging.info(f"LTS: Создан файл {filename}.xlsx ({count_rows} строк)")
        return filename, filepath

    def run(self):
        """Основной метод выполнения"""
        try:
            logging.info(f"=== Запуск скрипта LTS ===")

            data = self.gather_data()
            if data.empty:
                message = "Данных о поверке по Липецктеплосеть нет"
                logging.warning(message)
                self.send_telegram_notification(message, to_group=True)
                # Перепланирование перед ранним выходом
                schedule = self.config['schedule']
                send_day = schedule.get('send_day', 'Thursday')
                self.reschedule_next_run(
                    schedule_type='weekly',
                    send_time=schedule['send_time'],
                    send_day=send_day,
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

            success_msg = f"Письмо с файлом {filename} успешно отправлено на {email_config['recipient']}"
            logging.info(success_msg)
            self.send_telegram_notification(success_msg, to_group=True)

            schedule = self.config['schedule']
            send_day = schedule.get('send_day', 'Thursday')
            
            self.reschedule_next_run(
                schedule_type='weekly',
                send_time=schedule['send_time'],
                send_day=send_day,
                smart_reschedule=schedule.get('smart_reschedule', True)
            )

            logging.info(f"=== Скрипт LTS завершён успешно ===")

        except Exception as e:
            error_message = f"ОШИБКА в скрипте LTS: {e}"
            logging.error(error_message, exc_info=True)
            self.send_telegram_notification(error_message, to_group=True)
            raise


if __name__ == "__main__":
    from utils import setup_logging
    setup_logging()
    script = LTSScript()
    script.run()
