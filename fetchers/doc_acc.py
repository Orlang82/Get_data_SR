import xlwings as xw
from db.oracle import query
from utils.excel_writer import paste_to_excel
from utils.path_utils import get_sql_path

def fetch_to_doc_acc():
    # Получаем текущую книгу и лист DIFF
    wb = xw.Book.caller()

    # Получаем значения из именованных ячеек
    date_param = wb.names['date_end'].refers_to_range.value
    date_acc = wb.names['num_acc'].refers_to_range.value
    
    # Приводим даты к строкам (DD.MM.YYYY)
    date_param_str = date_param.strftime("%d.%m.%Y")
    date_acc_str = str(int(date_acc))

    sql_path = get_sql_path("SR_DOC_ACC_template.sql")
    with open(sql_path, encoding="utf-8") as f:
        sql = f.read().strip().rstrip(";")

    # Подставляем значения в SQL9
    sql = sql.replace(":date_param", f"'{date_param_str}'")
    sql = sql.replace(":date_acc", f"'{date_acc_str}'")
    
    return query(sql)

def paste_to_excel_doc_acc():
    df = fetch_to_doc_acc()
    paste_to_excel("DIFF", "tDetailAcc", df)
