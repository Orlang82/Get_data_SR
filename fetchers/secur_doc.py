from db.oracle import query
from utils.excel_writer import paste_to_excel
from utils.path_utils import get_sql_path

def fetch_to_secur_doc():
    # Получаем путь к SQL-файлу шаблона
    sql_path = get_sql_path("SR_SECUR_DOC_template.sql")
    # Открываем и читаем SQL-запрос из файла
    with open(sql_path, encoding="utf-8") as f:
        sql = f.read().strip().rstrip(";")
    # Выполняем запрос к базе данных и возвращаем результат
    return query(sql)

def paste_to_excel_secur_doc():
        # Получаем DataFrame с данными по ценным бумагам
    df = fetch_to_secur_doc()
    # Вставляем данные в Excel
    paste_to_excel("ОВДП", "tISIN", df)
