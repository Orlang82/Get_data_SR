# Импортируем класс BDay (Business Day) из модуля pandas.tseries.offsets,
# чтобы работать с рабочими днями (исключая выходные и праздники)
from pandas.tseries.offsets import BDay
import pandas as pd
import xlwings as xw

# Определяем функцию для получения предыдущего рабочего дня
def get_previous_working_day():
    # Получаем сегодняшнюю дату и время с помощью pd.Timestamp.today()
    # Затем вычитаем один рабочий день (BDay(1)) — это автоматически учитывает выходные и праздники
    # После этого приводим результат к дате (без времени) с помощью .date()
    return (pd.Timestamp.today() - BDay(1)).date()

def forecast_date():
    # Получаем текущую книгу и лист DIFF
    wb = xw.Book.caller()
    # Получаем значения из именованных ячеек
    date_forecast = wb.names['ForecastDate'].refers_to_range.value
    
 # Проверяем, что значение из ячейки не пустое (не None)
    if date_forecast:
        # Если в ячейке есть дата, форматируем ее и возвращаем
        # return date_forecast.strftime("%d.%m.%Y")
        return date_forecast
    else:
        # Если ячейка пуста, функция вернет None (пустоту)
        return None