from fetchers.secur_doc import paste_to_excel_secur_doc
from fetchers.grp_9000 import paste_to_excel_9000grp
from fetchers.dz_spot import paste_to_excel_dz_spot
from fetchers.balance_nrk import paste_to_excel_balance_nrk
from fetchers.dz_spot_diff import paste_to_excel_diff_spot
from mail.forecast_nrk import generate_forecast_email
from fetchers.diff_acc import paste_to_excel_diff_acc
from fetchers.fz_ccf_6jx import paste_to_excel_fz_ccf_6jx
from fetchers.doc_acc import paste_to_excel_doc_acc
from charts.chart_es import paste_plot_var_es
from fetchers.compens_579 import paste_to_excel_comp_579
from fetchers.detail_6jx import paste_to_excel_6jx_reserve
from fetchers.banks_42x import paste_to_excel_banks_42x
from fetchers.rc_component import paste_to_excel_rc_comp
from fetchers.detail_a7x import paste_to_excel_a7x_details
from charts.chart_as_v2 import insert_image_to_excel
from charts.chart_es_trade import paste_plot_var_es_trade
from charts.chart_as_trade import insert_chart_as_trade
from db.entry_db_6kx import process_single_6kx_file

# Основные вызовы (вызываются из Excel через xlwings)
def run_secur_doc():
    """Запускает вставку данных по security documents в Excel."""
    paste_to_excel_secur_doc()

def run_grp_9000():
    """Запускает вставку данных по группе 9000 в Excel."""
    paste_to_excel_9000grp()

def run_dz_spot():
    """Запускает вставку данных по dz spot в Excel."""
    paste_to_excel_dz_spot()

def run_balance_nrk():
    """Запускает вставку данных по балансу NRK в Excel."""
    paste_to_excel_balance_nrk()

def run_diff_spot():
    """Запускает вставку данных по diff spot в Excel."""
    paste_to_excel_diff_spot()

def run_forecast_mail():
    """Генерирует и отправляет прогнозное письмо."""
    generate_forecast_email()

def run_diff_acc():
    """Запускает вставку данных по diff accounts в Excel."""
    paste_to_excel_diff_acc()

def run_fz_ccf_6jx():
    """Запускает вставку данных по FZ CCF 6JX в Excel."""
    paste_to_excel_fz_ccf_6jx()

def run_doc_acc():
    """Запускает вставку данных по doc accounts в Excel."""
    paste_to_excel_doc_acc()

def run_plot_var_es():
    """Создает и вставляет график VAR ES в Excel."""
    paste_plot_var_es()

def run_compens_579():
    """Запускает вставку данных по компенсации 579 в Excel."""
    paste_to_excel_comp_579()

def run_6jx_reserve():
    """Запускает вставку данных по 6JX reserve в Excel."""
    paste_to_excel_6jx_reserve()

def run_42x_banks():
    """Запускает вставку данных для 42X по банкам в Excel."""
    paste_to_excel_banks_42x()

def run_rc_comp():
    """Запускает вставку данных для компонентов РК в Excel."""
    paste_to_excel_rc_comp()

def run_plot_as():
    """Создает и вставляет графики AS ES в Excel."""
    insert_image_to_excel()

def run_a7x_details():
    """Запускает вставку данных из файла DA7X в Excel."""
    paste_to_excel_a7x_details()

def run_plot_es_trade():
    """Создает и вставляет графики ES Trade в Excel."""
    paste_plot_var_es_trade()

def run_plot_as_trade():
    """Создает и вставляет графики AS Trade в Excel."""
    insert_chart_as_trade()

def run_single_6kx_file():
    """Обрабатывает один файл 6KX для добавление в БД."""
    process_single_6kx_file()