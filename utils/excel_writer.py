# Вставка данных из DataFrame в таблицу Excel
import xlwings as xw
import pandas as pd

def paste_to_excel(sheet_name: str, table_name: str, df: pd.DataFrame):
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
