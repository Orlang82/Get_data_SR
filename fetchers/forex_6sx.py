import logging
import os
import pandas as pd
from db.oracle import query
from utils.path_utils import get_sql_path
from utils.excel_writer import paste_to_excel_smart
from utils.parser_forex import parse_forex_numbers
from fetchers.pay_6sx import fetch_pay_6sx_data

# Логирование (установите True для включения)
ENABLE_LOGGING = True

def _setup_logger():
    """Настройка логгера для модуля forex_6sx."""
    if not ENABLE_LOGGING:
        return logging.getLogger("forex_6sx_disabled")
    logger = logging.getLogger("forex_6sx")
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    log_dir = os.path.abspath(os.path.join(script_dir, '..', 'logs'))
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, 'forex_6sx.log')
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    fh = logging.FileHandler(log_file, encoding='utf-8')
    fh.setFormatter(formatter)
    logger.addHandler(fh)
    return logger

logger = _setup_logger()


def fetch_forex_6sx_data() -> pd.DataFrame:
    """
    Формирует перечень forex-сделок по документам 6S.

    Алгоритм:
    1. Получает данные pay_6sx (поле DESCRIPTION).
    2. Парсит DESCRIPTION через parse_forex_numbers — извлекает номера сделок.
    3. Выполняет SQL SR_6SX_FOREX_template.sql с динамическим IN-clause.

    Returns:
        pd.DataFrame: колонки DOC_NO, DESCRIPTION, S135
    """
    logger.info("=== Начало fetch_forex_6sx_data ===")

    # Получаем документы 6S с полем DESCRIPTION
    df_pay = fetch_pay_6sx_data()
    logger.info(f"Получено строк из pay_6sx: {len(df_pay)}")

    if df_pay.empty or 'DESCRIPTION' not in df_pay.columns:
        logger.info("Нет данных DESCRIPTION, возвращаем пустой DataFrame")
        return pd.DataFrame(columns=['DOC_NO', 'DESCRIPTION', 'S135'])

    # Парсим все описания и собираем уникальные номера сделок
    all_numbers = []
    for desc in df_pay['DESCRIPTION'].dropna():
        all_numbers.extend(parse_forex_numbers(str(desc)))

    # Убираем дубли, сохраняем порядок
    unique_numbers = list(dict.fromkeys(all_numbers))
    logger.info(f"Найдено уникальных номеров сделок: {len(unique_numbers)}")

    if not unique_numbers:
        logger.info("Нет номеров сделок, возвращаем пустой DataFrame")
        return pd.DataFrame(columns=['DOC_NO', 'DESCRIPTION', 'S135'])

    # Читаем SQL-шаблон
    sql_path = get_sql_path("SR_6SX_FOREX_template.sql")
    with open(sql_path, encoding="utf-8") as f:
        sql = f.read().strip().rstrip(";")

    # Строим динамический IN-clause: заменяем :data_number на :v0, :v1, ...
    placeholders = ", ".join(f":v{i}" for i in range(len(unique_numbers)))
    sql = sql.replace(":data_number", placeholders)
    params = {f"v{i}": num for i, num in enumerate(unique_numbers)}

    df_result = query(sql, params)
    logger.info(f"Итого forex-сделок: {len(df_result)}")
    logger.info("=== Конец fetch_forex_6sx_data ===")
    return df_result


def paste_to_excel_forex_6sx(sheet_name="6SX_ACC"):
    """
    Записывает перечень forex-сделок в таблицу t6S_FOREX на листе 6SX_ACC.
    Вызывается из main.py через xlwings.
    """
    logger.info("=== Начало paste_to_excel_forex_6sx ===")
    try:
        df = fetch_forex_6sx_data()
        paste_to_excel_smart(sheet_name, "t6S_FOREX", df)
        logger.info("t6S_FOREX записана успешно")
    except Exception as e:
        logger.error(f"Ошибка в paste_to_excel_forex_6sx: {e}", exc_info=True)
        raise
    logger.info("=== Конец paste_to_excel_forex_6sx ===")
