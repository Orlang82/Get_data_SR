# Импортируем класс BDay (Business Day) из модуля pandas.tseries.offsets,
# чтобы работать с рабочими днями (исключая выходные и праздники)
from pandas.tseries.offsets import BDay
import pandas as pd

# Определяем функцию для получения предыдущего рабочего дня
def get_previous_working_day():
    # Получаем сегодняшнюю дату и время с помощью pd.Timestamp.today()
    # Затем вычитаем один рабочий день (BDay(1)) — это автоматически учитывает выходные и праздники
    # После этого приводим результат к дате (без времени) с помощью .date()
    return (pd.Timestamp.today() - BDay(1)).date()