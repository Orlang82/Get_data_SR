import xlwings as xw
import win32com.client as win32
from datetime import datetime
import os


def range_to_html_with_formatting(xl_range):
    """
    Преобразует диапазон Excel в HTML-таблицу с расширенным сохранением форматирования, включая границы.
    """
    html = '<table border="1" cellspacing="0" cellpadding="3" style="border-collapse: collapse;">'
    for row in xl_range.rows:
        html += '<tr>'
        for cell in row:
            value = cell.value if cell.value is not None else ''
            style = ''

            # Цвет текста
            if cell.api.Font.Color != 0:
                style += f'color: rgb{_bgr_to_rgb(cell.api.Font.Color)};'

            # Цвет заливки
            if cell.api.Interior.Color != 0:
                style += f'background-color: rgb{_bgr_to_rgb(cell.api.Interior.Color)};'

            # Жирный
            if cell.api.Font.Bold:
                style += 'font-weight: bold;'

            # Курсив
            if cell.api.Font.Italic:
                style += 'font-style: italic;'

            # Размер шрифта
            font_size = cell.api.Font.Size
            if font_size:
                style += f'font-size: {font_size}px;'

            # Имя шрифта
            font_name = cell.api.Font.Name
            if font_name:
                style += f'font-family: {font_name};'

            # Выравнивание по горизонтали
            align = cell.api.HorizontalAlignment
            if align == -4131:  # xlLeft
                style += 'text-align: left;'
            elif align == -4108:  # xlCenter
                style += 'text-align: center;'
            elif align == -4152:  # xlRight
                style += 'text-align: right;'

            # Границы (если есть)
            borders = cell.api.Borders
            border_styles = []
            for i in range(1, 5):  # 1-4 основные стороны ячейки
                border = borders(i)
                if border.LineStyle != 0:  # Если граница задана
                    border_styles.append('1px solid black')
            if border_styles:
                style += 'border: 1px solid black;'

            html += f'<td style="{style}">{value}</td>'
        html += '</tr>'
    html += '</table>'
    return html


def _bgr_to_rgb(bgr_color):
    """Преобразует цвет из формата BGR (как возвращает Excel) в формат RGB для CSS."""
    blue = bgr_color // 65536
    green = (bgr_color // 256) % 256
    red = bgr_color % 256
    return (red, green, blue)


def generate_forecast_email(excel_path=None):
    """
    Формирует и отображает письмо прогноза на основе данных из Excel.
    """

    wb = xw.Book.caller() if excel_path is None else xw.Book(excel_path)
    sheet = wb.sheets['Нрк_TEST']

    tables = {
        'tDZ_Spot': '',
        'tDZ_Spot_Diff': '',
    }

    for name in tables:
        table = None
        for tbl in sheet.tables:
            if tbl.name == name:
                table = tbl
                break
        if table is None:
            raise ValueError(f'Table {name} not found.')
        tables[name] = range_to_html_with_formatting(table.range)

    cell_impact = sheet.range('P:P').api.Find('Зміни за гр. 20% та 50%')
    cell_top15 = sheet.range('P:P').api.Find('ТОП 15 збільшення 100 % гр.')

    if cell_impact is None or cell_top15 is None:
        raise ValueError('Required ranges not found for "Impact" or "TOP 15".')

    impact_range = sheet.range((cell_impact.Row, cell_impact.Column)).resize(10, 5)
    top15_range = sheet.range((cell_top15.Row, cell_top15.Column)).resize(18, 9)

    html_impact = range_to_html_with_formatting(impact_range)
    html_top15 = range_to_html_with_formatting(top15_range)

    base_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(base_dir, 'temp_forecast.html')

    with open(template_path, 'r', encoding='utf-8') as f:
        html_body = f.read()

    today = datetime.today().strftime('%d.%m.%Y')
    html_body = html_body.replace('##today##', today)
    html_body = html_body.replace('##tDZ_Spot##', tables['tDZ_Spot'])
    html_body = html_body.replace('##tDZ_Spot_Diff##', tables['tDZ_Spot_Diff'])
    html_body = html_body.replace('##Impact##', html_impact)
    html_body = html_body.replace('##Top15##', html_top15)

    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.Subject = 'Прогноз ликвидности'
    mail.HTMLBody = html_body
    mail.Display()
