from db.oracle import query
from utils.date_utils import get_previous_working_day
from utils.excel_writer import paste_to_excel
from utils.path_utils import get_sql_path

def fetch_to_9000grp():
    # Путь к текущему файлу
    sql_path = get_sql_path("SR_CHECK_9000_template.sql")
    with open(sql_path, encoding="utf-8") as f:
        sql = f.read().strip().rstrip(";")
    date_param = get_previous_working_day()
    return query(sql, {"date_param": date_param})

def paste_to_excel_9000grp():
    df = fetch_to_9000grp()
    paste_to_excel("Нрк_TEST", "tSUM9000", df)
