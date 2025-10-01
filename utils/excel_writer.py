# Модуль для вставки данных из pandas DataFrame в таблицу Excel
import xlwings as xw  # Библиотека для работы с Excel
import pandas as pd   # Библиотека для работы с данными

def paste_to_excel(sheet_name: str, table_name: str, df: pd.DataFrame):
    """
    Вставляет данные из DataFrame в существующую таблицу Excel.
    
    Args:
        sheet_name (str): Имя листа Excel
        table_name (str): Имя таблицы Excel
        df (pd.DataFrame): DataFrame с данными для вставки
    """
    # Получаем активную книгу Excel
    wb = xw.Book.caller()
    app = wb.app
    
    # Отключаем обновление экрана и автоматические вычисления для ускорения работы
    app.screen_updating = False
    app.calculation = 'manual'
    
    # Получаем объекты листа и таблицы
    sheet = wb.sheets[sheet_name]
    table = sheet.api.ListObjects(table_name)
    
    # Очищаем существующие данные в таблице, если они есть
    if table.DataBodyRange:
        table.DataBodyRange.ClearContents()
    
    # Определяем начальную позицию для вставки (строка после заголовка, первый столбец таблицы)
    start_row = table.HeaderRowRange.Row + 1
    start_col = table.Range.Column
    
    # Создаем диапазон для новых данных и вставляем их
    # Заменяем NaN на пустые строки для корректного отображения
    data_range = sheet.range((start_row, start_col)).resize(len(df), len(df.columns))
    data_range.value = df.fillna('').values.tolist()
    
    # Изменяем размер таблицы, чтобы включить все новые данные
    # +1 в размере учитывает строку заголовка
    new_range = sheet.range((table.HeaderRowRange.Row, start_col)).resize(len(df) + 1, len(df.columns))
    table.Resize(new_range.api)
    
    # Возвращаем настройки Excel в исходное состояние
    app.calculation = 'automatic'
    app.screen_updating = True


def paste_to_excel_smart(sheet_name: str, table_name: str, df: pd.DataFrame):
    """
    Старый аналог процедуры paste_to_excel
    Данная версия корректно работает с smart-таблицами размещенных одна под другой
    """
    # Получаем активную книгу Excel, вызвавшую скрипт
    wb = xw.Book.caller()

    # Получаем нужный лист по имени
    sheet = wb.sheets[sheet_name]

    # Получаем таблицу Excel (ListObject) по имени
    table = sheet.api.ListObjects(table_name)

    # Определяем текущее количество строк в таблице
    current_row_count = table.ListRows.Count

    # Количество строк и столбцов в переданном DataFrame
    new_row_count = df.shape[0]
    col_count = df.shape[1]

    # Если в DataFrame больше строк, чем в таблице —
    # добавляем недостающие строки в таблицу Excel
    if new_row_count > current_row_count:
        for _ in range(new_row_count - current_row_count):
            table.ListRows.Add()

    # Определяем начальную ячейку для вставки данных —
    # это первая ячейка под заголовками таблицы
    dest_range = sheet.range((table.HeaderRowRange.Row + 1, table.Range.Column))

    # Определяем диапазон нужного размера (по числу строк и столбцов DataFrame)
    dest_range = dest_range.resize(new_row_count, col_count)

    # Вставляем значения из DataFrame как список списков (list of lists)
    dest_range.value = df.values.tolist()

    # Если в DataFrame меньше строк, чем в текущей таблице —
    # удаляем лишние строки таблицы снизу
    if new_row_count < current_row_count:
        for _ in range(current_row_count - new_row_count):
            table.ListRows(new_row_count + 1).Delete()

