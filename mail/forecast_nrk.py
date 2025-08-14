import xlwings as xw
import win32com.client as win32
import win32clipboard
import datetime
import time
import os
import re

def get_table_as_html(sheet, table_name):
    # Копируем диапазон как HTML (через COM-метод!)
    table = sheet.api.ListObjects(table_name)
    rng = table.Range
    rng.Copy()
    time.sleep(0.3)
    win32clipboard.OpenClipboard()
    try:
        html_data = win32clipboard.GetClipboardData(win32clipboard.RegisterClipboardFormat("HTML Format"))
    finally:
        win32clipboard.CloseClipboard()
    match = re.search(b"(<table.*?</table>)", html_data, re.DOTALL)
    if match:
        html_table = match.group(1).decode("utf-8")
    else:
        raise Exception("Не удалось извлечь HTML-таблицу из буфера обмена")
    return html_table

def generate_forecast_email():
    wb = xw.Book.caller()
    sheet = wb.sheets['Нрк_TEST']
    table_html = get_table_as_html(sheet, "tDZ_Spot")
    today = datetime.datetime.now().strftime('%d.%m.%Y')
    script_dir = os.path.dirname(os.path.abspath(__file__))
    html_path = os.path.join(script_dir, "temp_forecast.html")
    recipient = "youremail@company.com"  # Или подтяните из Excel
    subject = "Тема письма"              # Или подтяните из Excel

    # 1. Загружаем HTML шаблон
    with open(html_path, "r", encoding="windows-1251") as f:
        html_template = f.read()

    # 2. Делаем замены (учитывая ваши маркеры: ##today## и ##tDZ_Spot##)
    html_body = html_template
    html_body = html_body.replace("##today##", today)
    html_body = html_body.replace("##tDZ_Spot##", table_html)

    # 3. Создаём письмо через Outlook и вставляем HTML
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)  # Новый e-mail (olMailItem)
    mail.HTMLBody = html_body
    mail.Subject = subject
    mail.To = recipient
    mail.Display()  # mail.Send() если хотите сразу отправлять

if __name__ == "__main__":
    xw.Book("ВАША_КНИГА.xlsx").set_mock_caller()
    generate_forecast_email()
