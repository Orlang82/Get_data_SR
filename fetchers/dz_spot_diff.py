from db.oracle import query
from utils.date_utils import get_previous_working_day
from pandas.tseries.offsets import BDay
from utils.excel_writer import paste_to_excel
from utils.path_utils import get_sql_path

def fetch_to_diff_spot():

    sql_path = get_sql_path("SR_DIFF_DZ_SPOT_template.sql")
    with open(sql_path, encoding="utf-8") as f:
        sql = f.read().strip().rstrip(";")

    date_param_old = get_previous_working_day() - BDay(1)
    date_param = get_previous_working_day()

    # Форматируем даты как строки в формате 'DD.MM.YYYY'
    date_param_old_str = date_param_old.strftime("%d.%m.%Y")
    date_param_str = date_param.strftime("%d.%m.%Y")

    # Подставляем значения прямо в текст SQL-запроса
    sql = sql.replace(":date_param_old", f"'{date_param_old_str}'")
    sql = sql.replace(":date_param", f"'{date_param_str}'")

    # Теперь параметров не передаём — всё уже подставлено
    return query(sql)

def paste_to_excel_diff_spot():
    df = fetch_to_diff_spot()
    paste_to_excel("Нрк_TEST", "tDZ_Spot_Diff", df)
