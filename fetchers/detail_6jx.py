# -*- coding: utf-8 -*-
import xlwings as xw
import pandas as pd
import sys
import os
import tempfile

# --- Ваши импорты ---
from db.oracle import query
from utils.path_utils import get_sql_path
from utils.date_utils import get_previous_working_day

sys.stdout.reconfigure(encoding='utf-8')

def clear_and_paste(sheet_name: str, table_name: str, df_to_paste: pd.DataFrame):
    """Ваша функция для вставки данных (без изменений)."""
    try:
        wb = xw.Book.caller()
        sheet = wb.sheets[sheet_name]
        table = sheet.tables[table_name]
        if table.data_body_range is not None:
            table.data_body_range.clear_contents()
        if not df_to_paste.empty:
            sheet.range(table.header_row_range.address).offset(1, 0).value = df_to_paste.values
            table.resize(table.header_row_range.expand('down'))
        print(f"Данные ({len(df_to_paste)} строк) успешно вставлены в таблицу '{table_name}'.")
    except Exception as e:
        print(f"!!! Ошибка при вставке в Excel: {e}")

def get_initial_data_pairs() -> pd.DataFrame:
    """Получает DataFrame с парами для обработки (без изменений)."""
    print("Шаг 1: Чтение данных из 'F6JX_Details'...")
    wb = xw.Book.caller()
    sheet = wb.sheets['F6JX_Details']
    table_range = sheet.api.ListObjects("F6JX_Details").Range
    df = sheet.range(table_range.Address).options(pd.DataFrame, header=1, index=False).value
    
    df_filtered = df[df['Reserve_Up'] != 0].copy()
    if df_filtered.empty:
        return pd.DataFrame()
        
    df_filtered['ID рахунку'] = df_filtered['ID рахунку'].dropna().astype(int)
    df_filtered['ID договору'] = df_filtered['ID договору'].dropna().astype(int)
    
    df_pairs = df_filtered[['ID рахунку', 'ID договору']].drop_duplicates()
    print(f"Найдено {len(df_pairs)} уникальных пар (счет, договор) для обработки.")
    return df_pairs

def fetch_chunk_data(id_acc_chunk_str: str, id_ctr_chunk_str: str) -> pd.DataFrame:
    """Выполняет SQL-запрос для одного "чанка", убирая ';' (без изменений)."""
    sql_path = get_sql_path("SR_6JX_Reserve_template.sql")
    with open(sql_path, encoding="utf-8") as f:
        sql_template = f.read().strip().rstrip(';')

    date_param_str = get_previous_working_day().strftime("%d.%m.%Y")
    
    sql = sql_template.replace(":data_id_ctr", id_ctr_chunk_str)
    sql = sql.replace(":data_id_acc", id_acc_chunk_str)
    sql = sql.replace(":date_param", f"'{date_param_str}'")
    
    return query(sql)

def paste_to_excel_6jx_reserve():
    """Основная функция с обновленной логикой агрегации."""
    df_pairs = get_initial_data_pairs()
    
    if df_pairs.empty:
        clear_and_paste("F6JX_Details", "F6JX_Reserve", pd.DataFrame())
        return

    chunk_size = 25
    all_results = []
    
    print(f"Обработка будет вестись частями по {chunk_size} записей...")
    
    # --- Блок получения данных по "чанкам" (без изменений) ---
    for i, chunk_df in enumerate([df_pairs[i:i + chunk_size] for i in range(0, len(df_pairs), chunk_size)]):
        print(f"Обработка чанка {i + 1}...")
        id_acc_chunk = chunk_df['ID рахунку'].astype(str).tolist()
        id_ctr_chunk = chunk_df['ID договору'].astype(str).tolist()
        id_acc_chunk_str = ",".join([f"'{x}'" for x in id_acc_chunk])
        id_ctr_chunk_str = ",".join([f"'{x}'" for x in id_ctr_chunk])
        
        try:
            df_chunk_db = fetch_chunk_data(id_acc_chunk_str, id_ctr_chunk_str)
            if not df_chunk_db.empty:
                all_results.append(df_chunk_db)
        except Exception as e:
            print(f"!!! Ошибка при обработке чанка {i + 1}: {e}")
            continue

    # --- КЛЮЧЕВОЕ ИЗМЕНЕНИЕ: Обработка и агрегация данных ---
    if not all_results:
        final_df = pd.DataFrame()
    else:
        print("Объединение и агрегация результатов...")
        aggregated_df = pd.concat(all_results, ignore_index=True)
        
        # Создаем маску для определения типа операции. 
        # .str.upper() делает сравнение нечувствительным к регистру ('RESERVE', 'reserve', etc.)
        is_reserve_mask = aggregated_df['ACCOUNTING_TYPE'].str.upper() == 'RESERVE'
        
        # Создаем колонки RESERVE и BODY, распределяя суммы по условию
        aggregated_df['RESERVE'] = aggregated_df.where(is_reserve_mask, 0)['SUM_UAH']
        aggregated_df['BODY'] = aggregated_df.where(~is_reserve_mask, 0)['SUM_UAH']
        
        # Группируем по ключам и суммируем новые колонки.
        # as_index=False сразу создает плоский DataFrame без необходимости вызывать .reset_index()
        final_df = aggregated_df.groupby(
            ['ACCOUNT_ID', 'CONTRACT_ID', 'ACCOUNT_NUMBER', 'CODE'], 
            as_index=False
        )[['RESERVE', 'BODY']].sum()

    # Вставка финального результата в Excel (без изменений)
    clear_and_paste("F6JX_Details", "F6JX_Reserve", final_df)


if __name__ == "__main__":
    file_path = r"r:\Подразделения\РИСК-менеджмент\Внутренние\3 - РИСК ЛИКВИДНОСТИ\1 - БАЛАНС\12-09-2025\Balance_Bank_v5_12-09-25.xlsm"
    xw.Book(file_path).set_mock_caller()
    paste_to_excel_6jx_reserve()

