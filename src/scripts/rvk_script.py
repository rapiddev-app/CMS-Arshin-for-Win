"""
RVK - скрипт для еженедельной отправки данных в Росводоканал

Weekly скрипт:
- Собирает данные за последние 7 дней (не включая сегодня)
- Отправляет каждый четверг (или следующий рабочий день)
- Email: ev.shlykova@rosvodokanal.ru
- Telegram: группа + личный чат
"""

import os
import sys
import pandas as pd
from datetime import datetime, timedelta
from typing import Tuple, Any
import logging

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(__file__))))

from scripts.base_script import BaseScript
from config_manager import ConfigManager


class RVKScript(BaseScript):
    """
    Скрипт RVK (Росводоканал)

    Weekly скрипт для отправки реестра поверок ИПУ
    Собирает данные из Excel файла с многоуровневыми заголовками
    """

    def __init__(self):
        super().__init__(script_id='rvk', email_provider='gmail')
        self.config = ConfigManager().load_script_config('rvk')

    def gather_data(self) -> Any:
        """
        Сбор данных из Excel файла

        Особенности:
        - Двухуровневые заголовки (header=[0, 1])
        - Фильтрация по "ИПУ 1_Дата поверки" за последние 7 дней
        - Выбор колонок A-AQ (43 столбца)
        """
        paths = self.config['paths']

        # Получаем день недели из конфига
        schedule = self.config['schedule']
        send_day = schedule.get('send_day', 'Thursday')
        
        weekday_map = {
            'Monday': 0, 'Tuesday': 1, 'Wednesday': 2, 'Thursday': 3,
            'Friday': 4, 'Saturday': 5, 'Sunday': 6
        }
        base_weekday = weekday_map.get(send_day, 3)

        # Расчёт периода сбора (weekly: последние 7 дней)
        start_date, end_date = self.calculate_collection_period(
            actual_run_date=datetime.now(),
            base_weekday=base_weekday
        )

        try:
            # Читаем с многоуровневыми заголовками
            df = pd.read_excel(
                paths['input_file'],
                sheet_name=paths['sheet_name'],
                header=[0, 1],
                dtype={'Unnamed: 1_level_0_Лицевой счет': str}
            )

            logging.info(f"Прочитан файл: {paths['input_file']}, строк: {len(df)}")

            # Приводим "Лицевой счет" к строковому типу
            if ('Unnamed: 1_level_0', 'Лицевой счет') in df.columns:
                df[('Unnamed: 1_level_0', 'Лицевой счет')] = df[(
                    'Unnamed: 1_level_0', 'Лицевой счет'
                )].apply(lambda x: str(int(float(x))) if pd.notna(x) else '')

            # Преобразование дат
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

            # Объединяем уровни заголовков
            df.columns = ['_'.join(col).strip() for col in df.columns.values]

            # Фильтрация по дате поверки
            date_column = 'ИПУ 1_Дата поверки'
            if date_column in df.columns:
                df_filtered = df.dropna(subset=[date_column]).copy()
                df_filtered.loc[:, 'Дата поверки'] = pd.to_datetime(
                    df_filtered[date_column], errors='coerce'
                )
            else:
                raise KeyError(f"Колонка '{date_column}' не найдена")

            # Фильтр по периоду сбора
            df_filtered = df_filtered[
                (df_filtered['Дата поверки'] >= start_date) &
                (df_filtered['Дата поверки'] <= end_date)
            ]

            # Выбираем первые 43 столбца (A-AQ)
            df_filtered = df_filtered.iloc[:, :43]

            # Форматируем все даты для вывода
            for col in date_columns:
                if col in df_filtered.columns:
                    df_filtered[col] = pd.to_datetime(
                        df_filtered[col], errors='coerce'
                    ).dt.strftime('%d.%m.%Y')

            logging.info(f"После фильтрации осталось строк: {len(df_filtered)}")

            return df_filtered

        except FileNotFoundError:
            error_msg = f"Файл {paths['input_file']} не найден!"
            logging.error(error_msg)
            self.send_telegram_notification(error_msg, to_group=True)
            return pd.DataFrame()

        except Exception as e:
            error_msg = f"Ошибка при чтении файла: {e}"
            logging.error(error_msg, exc_info=True)
            self.send_telegram_notification(error_msg, to_group=True)
            return pd.DataFrame()

    def create_output_file(self, data: Any) -> Tuple[str, str]:
        """
        Создание выходного Excel файла

        Формат: "Реестр поверок ИПУ ({count}) {date_range} (РВК) ЦМС.xlsx"
        """
        df = data
        paths = self.config['paths']

        count_rows = len(df)

        # Расчёт диапазона дат
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

        # Имя файла
        filename = f"Реестр поверок ИПУ ({count_rows}) {date_range} (РВК) ЦМС"
        filepath = os.path.join(paths['output_folder'], f"{filename}.xlsx")

        # Удаляем существующий файл если есть (для РВК это нормально)
        if os.path.exists(filepath):
            try:
                os.remove(filepath)
                logging.info(f"Удалён существующий файл: {filename}.xlsx")
            except Exception as e:
                logging.warning(f"Не удалось удалить файл: {e}")

        # Сохраняем
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, header=True)

        logging.info(f"Создан файл: {filename}.xlsx ({count_rows} строк)")

        return filename, filepath

    def run(self):
        """
        Основной метод выполнения скрипта

        Последовательность:
        1. Сбор данных за последние 7 дней
        2. Создание файла
        3. Отправка email
        4. Telegram уведомление
        5. Перепланирование (weekly, Thursday)
        """
        try:
            logging.info(f"=== Запуск скрипта RVK ===")

            # Сбор данных
            data = self.gather_data()

            if data.empty or len(data) < 1:
                message = "Данных о поверке по РВК нет, отчёт не отправлен."
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

            # Создание файла
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

            # Успешное уведомление
            success_msg = f"Письмо с файлом {filename} успешно отправлено на {email_config['recipient']}"
            logging.info(success_msg)
            self.send_telegram_notification(success_msg, to_group=True)

            # Перепланирование (weekly)
            schedule = self.config['schedule']
            send_day = schedule.get('send_day', 'Thursday')

            self.reschedule_next_run(
                schedule_type='weekly',
                send_time=schedule['send_time'],
                send_day=send_day,
                smart_reschedule=schedule.get('smart_reschedule', True)
            )

            logging.info(f"=== Скрипт RVK завершён успешно ===")

        except Exception as e:
            error_message = f"ОШИБКА в скрипте RVK: {e}"
            logging.error(error_message, exc_info=True)
            self.send_telegram_notification(error_message, to_group=True)
            raise


if __name__ == "__main__":
    from utils import setup_logging
    setup_logging()

    script = RVKScript()
    script.run()
