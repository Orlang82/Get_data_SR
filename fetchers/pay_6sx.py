import xlwings as xw
import pandas as pd
import logging
import os
from db.oracle import query
from utils.path_utils import get_sql_path
from utils.excel_writer import paste_to_excel_smart
from fetchers.detail_6sx import fetch_6sx_data

# Логирование (отключено по умолчанию, установите True для включения)
ENABLE_LOGGING = True

def _setup_logger():
    """Настройка логгера для модуля pay_6sx."""
    if not ENABLE_LOGGING:
        return logging.getLogger("pay_6sx_disabled")

    logger = logging.getLogger("pay_6sx")
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    log_dir = os.path.abspath(os.path.join(script_dir, '..', 'logs'))
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, 'pay_6sx.log')

    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    return logger

logger = _setup_logger()


def fetch_pay_6sx_data():
    """
    Получает перечень документов, формирующих остатки для счетов 6S.

    Алгоритм:
    1. Читает отчетную дату RDATE из Excel.
    2. Получает перечень счетов к расчету (acc_calc) через fetch_6sx_data().
    3. Для каждого счета выполняет SQL-запрос и собирает результаты.

    Returns:
        pd.DataFrame: перечень документов с колонками R020, ACCOUNT_DT, CUR,
                      ACCOUNT_CT, DESCRIPTION, SUM_UAH
    """
    logger.info("=== Начало fetch_pay_6sx_data ===")

    # Получаем отчетную дату из именованной ячейки RDATE на листе menu
    wb = xw.Book.caller()
    try:
        rdate = wb.names['RDATE'].refers_to_range.value
        logger.info(f"Отчетная дата RDATE: {rdate}")
    except KeyError:
        logger.error("Именованная ячейка 'RDATE' не найдена в книге Excel")
        raise ValueError("Именованная ячейка 'RDATE' не найдена в книге Excel")

    # Получаем перечень счетов к расчету (без исключенных)
    acc_calc, _ = fetch_6sx_data()
    logger.info(f"Получено счетов для обработки: {len(acc_calc)}")

    # Если счетов нет — возвращаем пустой DataFrame
    if acc_calc.empty:
        logger.info("acc_calc пуст, возвращаем пустой DataFrame")
        return pd.DataFrame(columns=['R020', 'ACCOUNT_DT', 'CUR', 'ACCOUNT_CT', 'DESCRIPTION', 'SUM_UAH'])

    # Читаем SQL-шаблон запроса документов
    sql_path = get_sql_path("SR_6SX_PAY_template.sql")
    with open(sql_path, encoding="utf-8") as f:
        sql = f.read().strip().rstrip(";")

    # Выполняем запрос для каждого счета и собираем результаты
    results = []
    for _, row in acc_calc.iterrows():
        params = {
            "date_param": rdate,
            "data_acc": row['ACCOUNT_NUMBER'],  # строка — SUBSTR корректно даст R020
            "data_cur": row['CUR'],
        }
        df = query(sql, params)
        if not df.empty:
            results.append(df)
        logger.info(f"Счет {row['ACCOUNT_NUMBER']} ({row['CUR']}): найдено документов {len(df)}")

    # Объединяем все результаты в один DataFrame
    if results:
        df_all = pd.concat(results, ignore_index=True)
    else:
        df_all = pd.DataFrame(columns=['R020', 'ACCOUNT_DT', 'CUR', 'ACCOUNT_CT', 'DESCRIPTION', 'SUM_UAH'])

    logger.info(f"Итого документов: {len(df_all)}")
    logger.info("=== Конец fetch_pay_6sx_data ===")
    return df_all


def paste_to_excel_pay_6sx(sheet_name="6SX_ACC"):
    """
    Записывает перечень документов 6SX в таблицу t6S_PAY на листе 6SX_ACC.
    Вызывается из main.py через xlwings.

    Args:
        sheet_name (str): Имя листа Excel (по умолчанию "6SX_ACC")
    """
    logger.info("=== Начало paste_to_excel_pay_6sx ===")
    try:
        df = fetch_pay_6sx_data()
        paste_to_excel_smart(sheet_name, "t6S_PAY", df)
        logger.info("t6S_PAY записана успешно")
    except Exception as e:
        logger.error(f"Ошибка в paste_to_excel_pay_6sx: {e}", exc_info=True)
        raise
    logger.info("=== Конец paste_to_excel_pay_6sx ===")
