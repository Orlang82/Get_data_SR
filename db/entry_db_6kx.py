import xlwings as xw
import pandas as pd
import sqlite3
import os
import sys
from datetime import datetime

# Настройка кодировки для консоли Windows
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')
    # Альтернативный вариант, если верхний не сработает
    # locale.setpreferredencoding('UTF-8')

def process_single_6kx_file():
    """
    Основная функция обработки единичного файла 6КХ из Excel.
    Получает путь из именованной таблицы tPathF6KX на листе sys,
    обрабатывает файл и записывает данные в SQLite базу.
    """
    try:
        # Шаг 1: Получение активной книги Excel
        wb = xw.Book.caller()
        print("✓ Подключен к активной книге Excel")
        
        # Шаг 2: Получение пути к файлу из именованной таблицы
        try:
            # Получаем лист sys
            sys_sheet = wb.sheets['sys']
            
            # Получаем диапазон именованной таблицы tPathF6KX
            file_path_range = wb.names['tPathF6KX'].refers_to_range
            file_path = file_path_range.value
            
            print(f"✓ Получен путь к файлу: {file_path}")
            
        except Exception as e:
            print(f"❌ Ошибка при получении пути файла: {e}")
            return False

        # Шаг 3: Проверка существования файла
        if not os.path.exists(file_path):
            print(f"❌ Файл не существует: {file_path}")
            return False

        # Шаг 4: Проверка существования базы данных и таблиц
        db_path = r'r:\Подразделения\РИСК-менеджмент\Внутренние\3 - РИСК ЛИКВИДНОСТИ\1 - БАЛАНС\СКРИПТЫ\PyScripts\DataBase_6KX_6NX\database\liquidity_data.db'
        
        if not os.path.exists(db_path):
            print(f"❌ База данных не существует: {db_path}")
            print("❌ ОСТАНОВКА: Сначала создайте базу данных")
            return False
            
        # Проверяем существование необходимых таблиц
        if not check_required_tables(db_path):
            print("❌ ОСТАНОВКА: Необходимые таблицы не найдены в базе данных")
            return False

        # Шаг 5: Чтение данных из файла Excel
        try:
            # Читаем файл, пропуская первые 8 строк (как в оригинальном скрипте)
            df = pd.read_excel(file_path, skiprows=8, dtype=str)
            print(f"✓ Файл прочитан. Строк данных: {len(df)}")
            
        except Exception as e:
            print(f"❌ Ошибка при чтении файла: {e}")
            return False

        # Шаг 6: Валидация данных
        required_columns = ['REC_NO', 'EKP', 'R030', 'T100']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            print(f"❌ Отсутствуют обязательные колонки: {missing_columns}")
            return False
            
        if df.empty or df['EKP'].isna().all():
            print("❌ Файл не содержит данных")
            return False

        # Шаг 7: Извлечение даты из имени файла
        filename = os.path.basename(file_path)
        try:
            # Формат: 6КХ_DDMMYYYY.xlsx
            date_part = filename.split('_')[1].split('.')[0]
            date_obj = datetime.strptime(date_part, '%d%m%Y')
            file_date = date_obj.strftime('%d.%m.%Y')
            print(f"✓ Извлечена дата: {file_date}")
            
        except Exception as e:
            print(f"❌ Ошибка при извлечении даты из файла: {e}")
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
        print(f"✓ Подготовлены данные Combined_6KX_Data: {len(df_combined)} строк")

        # Шаг 9: Подготовка данных LCR_Combined
        # Фильтруем данные для A6K081 и A6K082
        lcr_081 = df_combined[df_combined['EKP'] == 'A6K081']
        lcr_082 = df_combined[df_combined['EKP'] == 'A6K082']
        
        # Создаем запись для LCR_Combined
        lcr_data = {
            'Date': file_date,
            'LCRвв': None,
            'LCRів': None,
            'Min_NRM': 100.0,
            'Target': 110.0
        }
        
        # Заполняем значения LCR (делим на 100 как в оригинальном коде)
        if not lcr_081.empty:
            lcr_data['LCRвв'] = float(lcr_081.iloc[0]['T100']) / 100
            
        if not lcr_082.empty:
            lcr_data['LCRів'] = float(lcr_082.iloc[0]['T100']) / 100
            
        print(f"✓ Подготовлены данные LCR_Combined для даты {file_date}")

        # Шаг 10: Запись в базу данных SQLite
        try:
            with sqlite3.connect(db_path) as conn:
                # Записываем данные Combined_6KX_Data
                df_combined.to_sql('DB_6KX', conn, if_exists='append', index=False)
                print(f"✓ Записано в DB_6KX: {len(df_combined)} строк")
                
                # Записываем данные LCR_Combined
                pd.DataFrame([lcr_data]).to_sql('LCR_Combined', conn, if_exists='append', index=False)
                print(f"✓ Записано в LCR_Combined: 1 строка")
                
        except Exception as e:
            print(f"❌ Ошибка при записи в базу данных: {e}")
            return False

        print("🎉 Обработка файла завершена успешно!")
        return True
        
    except Exception as e:
        print(f"❌ Критическая ошибка: {e}")
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
                print(f"❌ Отсутствуют таблицы: {', '.join(missing_tables)}")
                print("❌ Необходимо создать таблицы перед использованием скрипта")
                return False
            else:
                print(f"✓ Все необходимые таблицы найдены: {', '.join(required_tables)}")
                return True
                
    except Exception as e:
        print(f"❌ Ошибка при проверке таблиц: {e}")
        return False


if __name__ == "__main__":
    # Запуск основной функции
    process_single_6kx_file()
