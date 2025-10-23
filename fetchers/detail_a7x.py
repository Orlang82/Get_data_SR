# -*- coding: utf-8 -*-
"""
Скрипт для обработки данных из файла DA7X.
Читает путь из таблицы параметров, открывает файл,
фильтрует данные по счетам, начинающимся с "142",
и вставляет результат в таблицу tA7_Details.
"""
import sys
import os

# Добавляем корневую папку проекта в sys.path для возможности импорта модулей
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

# Настройка кодировки для корректного отображения кириллицы в консоли Windows
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

import xlwings as xw
import pandas as pd
from utils.excel_writer import paste_to_excel


def get_path_from_params() -> str:
    """
    Получает путь к файлу DA7X из таблицы параметров.

    Returns:
        str: Путь к файлу Excel
    """
    # Получаем активную книгу Excel
    wb = xw.Book.caller()

    # Получаем лист с параметрами
    sheet = wb.sheets['sys']

    # Получаем таблицу tParam
    table = sheet.api.ListObjects("tParam")
    table_range = table.Range

    # Читаем данные из таблицы в DataFrame
    df = sheet.range(table_range.Address).options(pd.DataFrame, header=1, index=False).value

    # Находим строку с параметром Path_DA7X
    path_row = df[df['Параметр'] == 'Path_DA7X']

    if path_row.empty:
        raise ValueError("Параметр 'Path_DA7X' не найден в таблице tParam")

    # Получаем значение пути
    path = path_row['Значение'].iloc[0]

    if pd.isna(path) or path == '':
        raise ValueError("Путь для параметра 'Path_DA7X' не указан")

    return path


def fetch_data_from_da7x() -> pd.DataFrame:
    """
    Читает данные из файла DA7X и фильтрует по счетам, начинающимся с "142".

    Returns:
        pd.DataFrame: Отфильтрованные данные
    """
    # Получаем путь к файлу
    file_path = get_path_from_params()

    print(f"Открытие файла: {file_path}")

    # Открываем файл Excel
    try:
        wb_source = xw.Book(file_path)
    except Exception as e:
        raise Exception(f"Не удалось открыть файл '{file_path}': {e}")

    try:
        # Читаем все данные из первого листа
        # (можно изменить на конкретное имя листа, если известно)
        sheet_source = wb_source.sheets[0]

        # Получаем все используемые ячейки
        used_range = sheet_source.used_range

        # Читаем данные в DataFrame
        df = sheet_source.range(used_range.address).options(pd.DataFrame, header=1, index=False).value

        print(f"Прочитано {len(df)} строк из файла")

        # Проверяем наличие столбца "R020 " (с пробелом)
        if 'R020 ' not in df.columns:
            # Пробуем без пробела
            if 'R020' in df.columns:
                column_name = 'R020'
            else:
                raise ValueError(f"Столбец 'R020 ' не найден в файле. Доступные столбцы: {df.columns.tolist()}")
        else:
            column_name = 'R020 '

        # Фильтруем данные: оставляем только строки, где значение в столбце R020 начинается с "142"
        # Преобразуем в строку и проверяем начало
        df_filtered = df[df[column_name].astype(str).str.startswith('142')].copy()

        print(f"После фильтрации осталось {len(df_filtered)} строк (счета, начинающиеся с '142')")

        return df_filtered

    finally:
        # Закрываем файл-источник без сохранения
        wb_source.close()
        print("Файл-источник закрыт")


def paste_to_excel_a7x_details():
    """
    Основная функция: получает данные из файла DA7X и вставляет их в таблицу tA7_Details.
    """
    try:
        # Получаем данные
        df = fetch_data_from_da7x()

        # Выводим информацию о данных для отладки
        print(f"\nИнформация о данных для вставки:")
        print(f"  Количество строк: {df.shape[0]}")
        print(f"  Количество столбцов: {df.shape[1]}")
        print(f"  Названия столбцов: {df.columns.tolist()}")

        # Вставляем данные в таблицу tA7_Details на листе #A7
        # Используем paste_to_excel
        paste_to_excel("#A7", "tA7_Details", df)

        print(f"✓ Данные успешно вставлены в таблицу 'tA7_Details' ({len(df)} строк)")

    except Exception as e:
        print(f"!!! Ошибка при обработке данных: {e}")
        raise


if __name__ == "__main__":
    # Для тестирования: укажите путь к вашей рабочей книге
    file_path = r"r:\Подразделения\РИСК-менеджмент\Внутренние\3 - РИСК ЛИКВИДНОСТИ\1 - БАЛАНС\23-10-2025\Balance_Bank_v6_23-10-25.xlsm"
    xw.Book(file_path).set_mock_caller()
    paste_to_excel_a7x_details()
