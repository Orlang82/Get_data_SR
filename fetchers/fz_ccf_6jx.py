from db.oracle import query
from utils.date_utils import get_previous_working_day
from utils.excel_writer import paste_to_excel
from utils.path_utils import get_sql_path

def fetch_to_fz_ccf_6jx():
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
    # Получаем DataFrame с данными
    df = fetch_to_fz_ccf_6jx()
    # Вставляем данные в Excel
    paste_to_excel("F6JX_Details", "FZ_for_6JX", df)
