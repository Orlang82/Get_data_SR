import xlwings as xw
from db.oracle import query
from utils.excel_writer import paste_to_excel
from utils.path_utils import get_sql_path

def fetch_to_compens_579():
    # Получаем текущую книгу
    wb = xw.Book.caller()

    # Получаем значения из именованных ячеек
    date_param = wb.names['RDATE'].refers_to_range.value
        
    # Приводим даты к строкам (DD.MM.YYYY)
    date_param_str = date_param.strftime("%d.%m.%Y")

    sql_path = get_sql_path("SR_COMPENSATION_579_template.sql")
    with open(sql_path, encoding="utf-8") as f:
        sql = f.read().strip().rstrip(";")

    # Подставляем значения в SQL
    sql = sql.replace(":date_param", f"'{date_param_str}'")
   
    return query(sql)

def paste_to_excel_comp_579():
    df = fetch_to_compens_579()
    paste_to_excel("menu", "Check_579", df)