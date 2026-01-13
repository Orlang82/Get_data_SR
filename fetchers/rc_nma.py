from db.oracle import query
from utils.date_utils import get_previous_working_day
from utils.excel_writer import paste_to_excel
from utils.path_utils import get_sql_path

def fetch_to_rc_nma():
    # Путь к текущему файлу
    sql_path = get_sql_path("SR_RC_NMA.sql")
    with open(sql_path, encoding="utf-8") as f:
        sql = f.read().strip().rstrip(";")
    
    # Определяем даты параметров в зависимости от режима прогноза
    date_param = get_previous_working_day()    
        
    return query(sql, {"date_param": date_param})

def paste_to_excel_rc_nma():
    df = fetch_to_rc_nma()
    paste_to_excel("Calculation_6RX", "NMA", df)