"""
Модуль для получения и вставки данных по финансовым заимствованиям (FZ)
с применением коэффициента конверсии (CCF) для отчета 6JX.
"""

from db.oracle import query
from utils.date_utils import get_previous_working_day
from utils.excel_writer import paste_to_excel
from utils.path_utils import get_sql_path


def fetch_to_fz_ccf_6jx():
    """
    Получает данные по финансовым заимствованиям с применением CCF из Oracle.

    Функция:
    1. Загружает SQL-запрос из файла SR_6JX_FZ_CCF_template.sql
    2. Использует дату предыдущего рабочего дня в качестве параметра
    3. Выполняет запрос к базе данных Oracle

    Returns:
        pandas.DataFrame: Таблица с данными по финансовым заимствованиям
    """
    # Путь к текущему SQL-файлу
    sql_path = get_sql_path("SR_6JX_FZ_CCF_template.sql")

    # Открываем и читаем SQL-запрос из файла
    with open(sql_path, encoding="utf-8") as f:
        sql = f.read().strip().rstrip(";")

    # Получаем предыдущий рабочий день для подстановки в запрос
    date_param = get_previous_working_day()

    # Выполняем запрос с параметром даты
    return query(sql, {"date_param": date_param})


def paste_to_excel_fz_ccf_6jx():
    """
    Вставляет данные по финансовым заимствованиям с CCF в Excel-файл.

    Функция получает данные через fetch_to_fz_ccf_6jx() и вставляет их
    в лист "FZ_for_6JX" Excel-файла "F6JX_Details".
    """
    # Получаем DataFrame с данными
    df = fetch_to_fz_ccf_6jx()

    # Вставляем данные в Excel: книга "F6JX_Details", лист "FZ_for_6JX"
    paste_to_excel("F6JX_Details", "FZ_for_6JX", df)
