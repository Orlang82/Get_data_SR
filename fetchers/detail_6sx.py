import xlwings as xw
import pandas as pd
import logging
import os
from db.oracle import query
from utils.path_utils import get_sql_path
from utils.excel_writer import paste_to_excel_smart

# Настройка логирования (отключено по умолчанию)
ENABLE_LOGGING = True  # Установите True для включения логов

def _setup_logger():
    """Настройка логгера для модуля detail_6sx."""
    if not ENABLE_LOGGING:
        return logging.getLogger("detail_6sx_disabled")

    logger = logging.getLogger("detail_6sx")
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    log_dir = os.path.abspath(os.path.join(script_dir, '..', 'logs'))
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, 'detail_6sx.log')

    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    return logger

logger = _setup_logger()


def fetch_6sx_data():
    """
    Получает и обрабатывает данные для формирования перечня счетов 6S.

    Returns:
        tuple: (acc_calc, acc_exclude) - два DataFrame для записи в Excel
    """
    # Получаем активную книгу Excel
    wb = xw.Book.caller()

    # Шаг 1: Получаем отчетную дату из именованной ячейки RDATE на листе menu
    try:
        rdate = wb.names['RDATE'].refers_to_range.value
    except KeyError:
        raise ValueError("Именованная ячейка 'RDATE' не найдена в книге Excel")
    except Exception as e:
        raise ValueError(f"Ошибка получения даты RDATE: {e}")

    # Шаг 2: Получаем перечень счетов, которые уже исключены из расчета 6SX
    sql_path_exclude = get_sql_path("SR_6SX_EXCLUDE_template.sql")
    with open(sql_path_exclude, encoding="utf-8") as f:
        sql_exclude = f.read().strip().rstrip(";")

    df_exclude = query(sql_exclude)

    # Формируем множество номеров счетов для быстрой проверки (только по ACCOUNT_NUMBER)
    excluded_accounts = set(df_exclude['ACCOUNT_NUMBER'].tolist())

    # Шаг 3: Получаем перечень счетов с остатками на текущий день
    sql_path_account = get_sql_path("SR_6SX_ACCOUNT_template.sql")
    with open(sql_path_account, encoding="utf-8") as f:
        sql_account = f.read().strip().rstrip(";")

    df_account = query(sql_account, {"date_param": rdate})

    # Добавляем колонку для пометок
    df_account['mark'] = None

    # Шаг 4: Логика обработки - помечаем записи для исключения
    # ВАЖНО: сначала проверяем наличие в списке исключений, потом остальные условия

    # 4.1: Сравнение с перечнем ранее исключенных счетов (только по ACCOUNT_NUMBER)
    mask_pre_excluded = df_account['ACCOUNT_NUMBER'].isin(excluded_accounts)
    df_account.loc[mask_pre_excluded, 'mark'] = 'pre_excluded'

    # 4.2: Проверка валюты 980 (UAH) - код валюты может быть как число, так и текст
    # Проверяем только для еще не помеченных
    mask_980 = (df_account['CUR'] == 980) | (df_account['CUR'] == '980')
    df_account.loc[mask_980 & (df_account['mark'].isna()), 'mark'] = 'exclude'

    # 4.3: Проверка наличия слов "транз" или "транс" в названии (регистронезависимо)
    # Проверяем только для еще не помеченных
    mask_tranz = df_account['NAME_ACC'].str.contains(
        r'транз|транс',
        case=False,
        na=False,
        regex=True
    )
    df_account.loc[mask_tranz & (df_account['mark'].isna()), 'mark'] = 'exclude'

    # Шаг 5: Формируем два списка
    # acc_exclude - записи помеченные как exclude или pre_excluded
    acc_exclude = df_account[df_account['mark'].isin(['exclude', 'pre_excluded'])].copy()

    # acc_calc - записи не помеченные (mark = None)
    acc_calc = df_account[df_account['mark'].isna()].copy()

    # Выбираем только нужные колонки
    columns_to_keep = ['R020', 'ACCOUNT_NUMBER', 'CUR', 'SUM_UAH', 'NAME_ACC']
    acc_calc = acc_calc[columns_to_keep]
    acc_exclude = acc_exclude[columns_to_keep + ['mark']]

    return acc_calc, acc_exclude


def apply_exclude_formatting(sheet_name, table_name, df_exclude):
    """
    Применяет форматирование к таблице t6S_EXCLUDE:
    - pre_excluded: серый шрифт (35%), зачеркнутый
    - exclude: красный шрифт, полужирный

    Args:
        sheet_name (str): Имя листа Excel
        table_name (str): Имя таблицы Excel
        df_exclude (pd.DataFrame): DataFrame с данными исключений (должен содержать колонку 'mark')
    """
    wb = xw.Book.caller()
    sheet = wb.sheets[sheet_name]
    table = sheet.api.ListObjects(table_name)

    # Получаем начальную строку данных таблицы (после заголовка)
    start_row = table.HeaderRowRange.Row + 1
    start_col = table.Range.Column

    # Применяем форматирование построчно
    for i, (idx, row) in enumerate(df_exclude.iterrows()):
        # Номер строки в Excel (используем счетчик i, а не индекс из DataFrame)
        excel_row = start_row + i

        # Получаем диапазон для всей строки таблицы
        row_range = sheet.range((excel_row, start_col)).resize(1, 5)  # 5 колонок

        if row['mark'] == 'pre_excluded':
            # Серый шрифт 35% (RGB: 166, 166, 166) и зачеркнутый
            row_range.api.Font.Color = 0xA6A6A6  # В Excel используется BGR формат
            row_range.api.Font.Strikethrough = True
            row_range.api.Font.Bold = False

        elif row['mark'] == 'exclude':
            # Красный шрифт (RGB: 255, 0, 0) и полужирный
            row_range.api.Font.Color = 0x0000FF  # Красный в BGR формате
            row_range.api.Font.Bold = True
            row_range.api.Font.Strikethrough = False


def paste_to_excel_detail_6sx(sheet_name="6SX_ACC"):
    """
    Основная функция для вставки данных 6SX в Excel.
    Вызывается из main.py через xlwings.

    Args:
        sheet_name (str): Имя листа Excel (по умолчанию "6SX_ACC")
    """
    # Получаем обработанные данные
    acc_calc, acc_exclude = fetch_6sx_data()

    # Записываем acc_calc в таблицу t6S_TO_CALC
    paste_to_excel_smart(sheet_name, "t6S_TO_CALC", acc_calc)

    # Записываем acc_exclude в таблицу t6S_EXCLUDE (без колонки mark)
    acc_exclude_output = acc_exclude[['R020', 'ACCOUNT_NUMBER', 'CUR', 'SUM_UAH', 'NAME_ACC']]
    paste_to_excel_smart(sheet_name, "t6S_EXCLUDE", acc_exclude_output)

    # Применяем форматирование к таблице t6S_EXCLUDE
    apply_exclude_formatting(sheet_name, "t6S_EXCLUDE", acc_exclude)
