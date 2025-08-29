import xlwings as xw
from db.oracle import query
from utils.excel_writer import paste_to_excel
from utils.path_utils import get_sql_path

def fetch_to_diff_acc():
    # Получаем текущую книгу и лист DIFF
    wb = xw.Book.caller()

    # Получаем значения из именованных ячеек
    date_param_old = wb.names['date_start'].refers_to_range.value
    date_param = wb.names['date_end'].refers_to_range.value
    date_r020 = wb.names['d_r020'].refers_to_range.value
    
    # Приводим даты к строкам (DD.MM.YYYY)
    date_param_old_str = date_param_old.strftime("%d.%m.%Y")
    date_param_str = date_param.strftime("%d.%m.%Y")
    # Обработка date_r020 в зависимости от типа
    if isinstance(date_r020, str):
        date_r020_str = date_r020
    else:
        # Если это число (включая float), приводим к int, затем к строке
        date_r020_str = str(int(date_r020))

    sql_path = get_sql_path("SR_DIFF_ACC_template.sql")
    with open(sql_path, encoding="utf-8") as f:
        sql = f.read().strip().rstrip(";")

    # Подставляем значения в SQL9
    sql = sql.replace(":date_param_old", f"'{date_param_old_str}'")
    sql = sql.replace(":date_param", f"'{date_param_str}'")
    sql = sql.replace(":date_r020", f"'{date_r020_str}'")
    if date_r020_str == '2600':
        sql = sql.replace(":over_param", "ACS.BASE_AMOUNT > 0")
    else:
        sql = sql.replace(":over_param", "ACS.BASE_AMOUNT <> 0")

    return query(sql)

def paste_to_excel_diff_acc():
    df = fetch_to_diff_acc()
    paste_to_excel("DIFF", "tDiffAcc", df)
