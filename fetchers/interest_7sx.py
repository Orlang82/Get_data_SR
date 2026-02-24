import os
import logging
import pandas as pd
import xlwings as xw

from db.oracle import query
from utils.excel_writer import paste_to_excel
from utils.path_utils import get_sql_path

# --- Логирование (отключить при необходимости установив False) ---
ENABLE_LOGGING = True

def _setup_logger():
    if not ENABLE_LOGGING:
        return logging.getLogger("interest_7sx_disabled")
    logger = logging.getLogger("interest_7sx")
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    log_dir = os.path.abspath(os.path.join(script_dir, '..', 'logs'))
    os.makedirs(log_dir, exist_ok=True)
    handler = logging.FileHandler(os.path.join(log_dir, 'interest_7sx.log'), encoding='utf-8')
    handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logger.addHandler(handler)
    return logger

logger = _setup_logger()

# --- Границы тенора (дни, метка) — редактируются здесь ---
TENOR_BOUNDS = [
    (31,   "1M"),
    (92,   "3M"),
    (183,  "6M"),
    (365,  "1Y"),
    (729,  "2Y"),
    (1094, "3Y"),
    (1459, "4Y"),
    (1824, "5Y"),
    (2554, "7Y"),
    (3649, "10Y"),
    (5474, "15Y"),
    (7299, "20Y"),
]

# --- Коэффициенты для расчёта SUM_CALC по тенору — редактируются здесь ---
TENOR_COEF = {
    "1M":    0,
    "3M":    0.0020,
    "6M":    0.0040,
    "1Y":    0.0070,
    "2Y":    0.0125,
    "3Y":    0.0175,
    "4Y":    0.0225,
    "5Y":    0.0275,
    "7Y":    0.0325,
    "10Y":   0.0375,
    "15Y":   0.0450,
    "20Y":   0.0525,
    "UP20Y": 0.0600,
}


def _tenor_label(days):
    """Возвращает метку тенора по количеству дней."""
    if days is None or (isinstance(days, float) and pd.isna(days)):
        return ""
    for bound, label in TENOR_BOUNDS:
        if days <= bound:
            return label
    return "UP20Y"


def fetch_interest_7sx():
    """Получает данные процентного риска торговой книги 7S."""
    # Читаем отчётную дату из именованной ячейки RDATE7SX листа Calculation
    wb = xw.Book.caller()
    rdate = wb.names['RDATE7SX'].refers_to_range.value
    logger.info(f"Отчётная дата RDATE7SX: {rdate}")

    # Загружаем SQL-шаблон
    sql_path = get_sql_path("SR_7S_INTEREST_RISK_template.sql")
    with open(sql_path, encoding="utf-8") as f:
        sql = f.read().strip().rstrip(";")

    # Выполняем запрос с подстановкой даты
    df = query(sql, {"date_param": rdate})
    logger.info(f"Получено строк: {len(df)}")

    # Приводим отчётную дату к Timestamp для арифметики
    date_ts = pd.Timestamp(rdate)

    # COUNT_DAY: разница DATE_END минус отчётная дата в днях
    df['COUNT_DAY'] = (pd.to_datetime(df['DATE_END']) - date_ts).dt.days

    # TENOR: бакет по количеству дней
    df['TENOR'] = df['COUNT_DAY'].apply(_tenor_label)

    # SUM_CALC: SUM_UAH умноженное на коэффициент тенора
    df['SUM_CALC'] = df.apply(
        lambda r: r['SUM_UAH'] * TENOR_COEF[r['TENOR']] if r['TENOR'] != "" else "",
        axis=1
    )

    return df


def paste_to_excel_interest_7sx():
    """Записывает данные процентного риска 7S в таблицу t7S_INTEREST листа TRAD_F01X."""
    df = fetch_interest_7sx()
    paste_to_excel("TRAD_F01X", "t7S_INTEREST", df)
