import xlwings as xw
import pandas as pd
import sqlite3
import os
import sys
import logging
from datetime import datetime

# Настройка кодировки для консоли Windows
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')
    # Альтернативный вариант, если верхний не сработает
    # locale.setpreferredencoding('UTF-8')

def _setup_logger():
    """Configure module-level logger that writes into logs/entry_db_6kx.log."""
    logger = logging.getLogger("entry_db_6kx")
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    log_dir = os.path.abspath(os.path.join(script_dir, '..', 'logs'))
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, 'entry_db_6kx.log')

    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    return logger


logger = _setup_logger()

def process_single_6kx_file():
    """
    Основная функция обработки единичного файла 6КХ из Excel.
    Получает путь из именованной таблицы tPathF6KX на листе sys,
    обрабатывает файл и записывает данные в SQLite базу.
    """
    try:
        # Шаг 1: Получение активной книги Excel
        wb = xw.Book.caller()
        logger.info("✓ Подключен к активной книге Excel")
        
        # Шаг 2: Получение пути к файлу из именованной таблицы
        try:
            # Получаем лист sys
            sys_sheet = wb.sheets['sys']
            
            # Таблица tPathF6KX содержит столбец Path с путями к файлам
            # path_table = sys_sheet.tables['tPathF6KX']
            # table_df = path_table.range.options(pd.DataFrame, header=1, index=False).value
            # if 'Path' not in table_df or table_df['Path'].dropna().empty:
            #     raise ValueError("Таблица tPathF6KX не содержит заполненного столбца 'Path'")
            # file_path = table_df['Path'].dropna().iloc[0]
            
            # Таблица tPathF6KX содержит столбец Path с путями к файлам
            path_table = sys_sheet.tables['tPathF6KX']

            # Получаем диапазон таблицы и читаем как DataFrame
            # header=1 означает, что первая строка - заголовки
            table_df = path_table.range.options(pd.DataFrame, header=1, index=False).value

            # Проверяем наличие столбца Path и его заполненность
            if 'Path' not in table_df.columns or table_df['Path'].dropna().empty:
                raise ValueError("Таблица tPathF6KX не содержит заполненного столбца 'Path'")

            # Получаем первый непустой путь
            file_path = table_df['Path'].dropna().iloc[0]


            logger.info("✓ Получен путь к файлу: %s", file_path)
            
        except Exception as e:
            logger.error("❌ Ошибка при получении пути файла: %s", e)
            return False
        
        # Шаг 3: Проверка существования файла
        if not os.path.exists(file_path):
            logger.error("❌ Файл не существует: %s", file_path)
            return False

        # Шаг 4: Проверка существования базы данных и таблиц
        db_path = r'r:\Подразделения\РИСК-менеджмент\Внутренние\3 - РИСК ЛИКВИДНОСТИ\1 - БАЛАНС\СКРИПТЫ\PyScripts\DataBase_6KX_6NX\database\liquidity_data.db'
        
        if not os.path.exists(db_path):
            logger.error("❌ База данных не существует: %s", db_path)
            logger.error("❌ ОСТАНОВКА: Сначала создайте базу данных")
            return False
            
        # Проверяем существование необходимых таблиц
        if not check_required_tables(db_path):
            logger.error("❌ ОСТАНОВКА: Необходимые таблицы не найдены в базе данных")
            return False

        # Шаг 5: Чтение данных из файла Excel
        try:
            # Читаем файл, пропуская первые 8 строк (как в оригинальном скрипте)
            df = pd.read_excel(file_path, skiprows=8, dtype=str)
            logger.info("✓ Файл прочитан. Строк данных: %s", len(df))
            
        except Exception as e:
            logger.error("❌ Ошибка при чтении файла: %s", e)
            return False

        # Шаг 6: Валидация данных
        required_columns = ['REC_NO', 'EKP', 'R030', 'T100']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            logger.error("❌ Отсутствуют обязательные колонки: %s", missing_columns)
            return False
            
        if df.empty or df['EKP'].isna().all():
            logger.error("❌ Файл не содержит данных")
            return False

        # Шаг 7: Извлечение даты из имени файла
        filename = os.path.basename(file_path)
        try:
            # Формат: 6КХ_DDMMYYYY.xlsx
            date_part = filename.split('_')[1].split('.')[0]
            date_obj = datetime.strptime(date_part, '%d%m%Y')
            file_date = date_obj.strftime('%Y-%m-%d')
            logger.info("✓ Извлечена дата: %s", file_date)
            
        except Exception as e:
            logger.error("❌ Ошибка при извлечении даты из файла: %s", e)
            return False

        # Шаг 8: Обработка данных для Combined_6KX_Data
        df_combined = df[['REC_NO', 'EKP', 'R030', 'T100']].copy()
        df_combined['Date'] = file_date
        
        # Рассчитываем R031 (тип валюты)
        def calculate_r031(r030_value):
            if str(r030_value) == '980':
                return 'NV'  # Национальная валюта
            elif str(r030_value) == '#':
                return '#'   # Неопределенная
            else:
                return 'FCY' # Иностранная валюта
        
        df_combined['R031'] = df_combined['R030'].apply(calculate_r031)
        
        # Переупорядочиваем колонки
        df_combined = df_combined[['Date', 'REC_NO', 'EKP', 'R030', 'R031', 'T100']]
        logger.info("✓ Подготовлены данные Combined_6KX_Data: %s строк", len(df_combined))

        # Шаг 9: Подготовка данных LCR_Combined
        # Фильтруем данные для A6K081 и A6K082
        lcr_081 = df_combined[df_combined['EKP'] == 'A6K081']
        lcr_082 = df_combined[df_combined['EKP'] == 'A6K082']
        
        # Создаем запись для LCR_Combined
        lcr_data = {
            'Date': file_date,
            'LCRвв': None,
            'LCRів': None,
            'Min_NRM': 1.00,
            'Target': 1.10
        }
        
        # Заполняем значения LCR (делим на 100 как в оригинальном коде)
        if not lcr_081.empty:
            lcr_data['LCRвв'] = float(lcr_081.iloc[0]['T100']) / 100
            
        if not lcr_082.empty:
            lcr_data['LCRів'] = float(lcr_082.iloc[0]['T100']) / 100
            
        logger.info("✓ Подготовлены данные LCR_Combined для даты %s", file_date)

        # Шаг 10: Запись в базу данных SQLite
        try:
            with sqlite3.connect(db_path) as conn:
                # Записываем данные Combined_6KX_Data
                df_combined.to_sql('DB_6KX', conn, if_exists='append', index=False)
                logger.info("✓ Записано в DB_6KX: %s строк", len(df_combined))
                
                # Записываем данные LCR_Combined
                pd.DataFrame([lcr_data]).to_sql('LCR_Combined', conn, if_exists='append', index=False)
                logger.info("✓ Записано в LCR_Combined: 1 строка")
                
        except Exception as e:
            logger.error("❌ Ошибка при записи в базу данных: %s", e)
            return False

        logger.info("🎉 Обработка файла завершена успешно!")
        return True
        
    except Exception as e:
        logger.exception("❌ Критическая ошибка")
        return False


def check_required_tables(db_path):
    """
    Проверяет существование необходимых таблиц в базе данных.
    
    Args:
        db_path: Путь к файлу базы данных
        
    Returns:
        True если все таблицы существуют, False иначе
    """
    required_tables = ['DB_6KX', 'LCR_Combined']
    
    try:
        with sqlite3.connect(db_path) as conn:
            cursor = conn.cursor()
            
            # Получаем список всех таблиц
            cursor.execute("""
                SELECT name FROM sqlite_master 
                WHERE type='table'
            """)
            existing_tables = [row[0] for row in cursor.fetchall()]
            
            # Проверяем наличие каждой необходимой таблицы
            missing_tables = []
            for table in required_tables:
                if table not in existing_tables:
                    missing_tables.append(table)
            
            if missing_tables:
                logger.error("❌ Отсутствуют таблицы: %s", ', '.join(missing_tables))
                logger.error("❌ Необходимо создать таблицы перед использованием скрипта")
                return False
            else:
                logger.info("✓ Все необходимые таблицы найдены: %s", ', '.join(required_tables))
                return True
                
    except Exception as e:
        logger.error("❌ Ошибка при проверке таблиц: %s", e)
        return False


if __name__ == "__main__":
    # Запуск основной функции
    process_single_6kx_file()
