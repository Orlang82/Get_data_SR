from db.oracle import query
from utils.date_utils import get_previous_working_day
from utils.excel_writer import paste_to_excel
from utils.path_utils import get_sql_path
from utils.date_utils import forecast_date


def fetch_to_banks_42x():
    """
    Получает данные из Oracle по форме 42X (межбанковские операции).

    Функция загружает SQL-запрос из шаблона, подставляет параметр даты
    в зависимости от режима работы (актуальные данные или прогноз)
    и выполняет запрос к базе данных.

    Returns:
        DataFrame: результат выполнения SQL-запроса с данными по форме 42X
    """
    # Получаем путь к SQL-шаблону для формы 42X
    sql_path = get_sql_path("SR_BANKS_42X_template.sql")

    # Читаем содержимое SQL-файла
    with open(sql_path, encoding="utf-8") as f:
        sql = f.read().strip().rstrip(";")

    # Определяем дату для запроса в зависимости от режима работы
    if not forecast_date():
        # Если прогнозная дата не установлена - берем предыдущий рабочий день
        date_param = get_previous_working_day()
    else:
        # Если установлена прогнозная дата - используем её
        date_param = forecast_date()

    # Выполняем запрос к Oracle с параметром даты
    return query(sql, {"date_param": date_param})


def paste_to_excel_banks_42x():
    """
    Загружает данные по форме 42X и вставляет их в Excel.

    Функция получает данные из базы данных и записывает их
    в указанный лист Excel в именованную таблицу.
    """
    # Получаем данные из базы
    df = fetch_to_banks_42x()

    # Вставляем данные в Excel на лист "F42X" в таблицу "tActualForecast42X"
    paste_to_excel("F42X", "tActualForecast42X", df)
    