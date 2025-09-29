from db.oracle import query
from utils.date_utils import get_previous_working_day
from utils.excel_writer import paste_to_excel
from utils.path_utils import get_sql_path
from utils.date_utils import forecast_date

def fetch_to_dz_spot():
    # Путь к текущему файлу
    sql_path = get_sql_path("SR_CHECK_DZ_SPOT_template.sql")
    with open(sql_path, encoding="utf-8") as f:
        sql = f.read().strip().rstrip(";")

    # Определяем даты параметров в зависимости от режима прогноза
    if not forecast_date():
        date_param = get_previous_working_day()    
    else:
        date_param = forecast_date()
    
    return query(sql, {"date_param": date_param})

def paste_to_excel_dz_spot():
    df = fetch_to_dz_spot()
    paste_to_excel("Нрк_TEST", "tDZ_Spot", df)
    