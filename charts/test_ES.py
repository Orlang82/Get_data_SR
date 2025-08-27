import xlwings as xw
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import os
from scipy.stats import gaussian_kde

def get_loss_data(sheet_name='Scenario_Level_ES', col='CN', start_row=7):
    wb = xw.Book.caller()
    ws = wb.sheets[sheet_name]
    last_row = ws.range(f'{col}1048576').end('up').row
    data = ws.range(f'{col}{start_row}:{col}{last_row}').options(np.array, dtype=float).value
    if data.ndim > 1:
        data = data.flatten()
    data = data[~np.isnan(data)]
    data = data[data != 0]
    return data

def get_named_value(wb, name):
    """Получить значение из именованной ячейки"""
    return float(wb.names[name].refers_to_range.value)

def plot_distribution(data, var, es, img_path='distribution.png'):
    sns.set_theme(style="whitegrid")
    plt.figure(figsize=(8, 5))
    
    sns.histplot(data, bins=30, kde=True, stat="density", color="skyblue", edgecolor="black", label="Гистограмма")
    sns.kdeplot(data, color="blue", linewidth=2, label="Щільність")
    
    # VaR линия
    plt.axvline(var, color="red", linestyle="--", label=f"VaR = {var:,.2f}")
    # ES линия
    plt.axvline(es, color="orange", linestyle="--", label=f"ES = {es:,.2f}")
    # Закрашиваем хвост правее VaR
    plt.fill_betweenx(
        y=[0, plt.gca().get_ylim()[1]],
        x1=var, x2=plt.gca().get_xlim()[1],
        color='red', alpha=0.01, label="Хвост (L > VaR)"
    )
    
    def plot_density_var_es(data, var, es, img_path='distribution.png'):
        sns.set_theme(style="whitegrid")
    plt.figure(figsize=(9, 5))

    # Ограничиваем область по X (центрируем колокол)
    x_min = max(-1.2e6, np.min(data))
    x_max = np.max([np.max(data), es]) * 1.7
    x_grid = np.linspace(x_min, x_max, 1000)

    # KDE плотность
    kde = gaussian_kde(data)
    y_kde = kde(x_grid)

    # Индексы до и после VaR
    idx_var = np.searchsorted(x_grid, var)
    idx_es = np.searchsorted(x_grid, es)

    # --- Закраска до VaR (голубая область) ---
    plt.fill_between(x_grid[:idx_var], 0, y_kde[:idx_var], color='skyblue', alpha=0.6, label='P(Loss ≤ VaR)')
    # --- Закраска после VaR (розовая область) ---
    plt.fill_between(x_grid[idx_var:], 0, y_kde[idx_var:], color='pink', alpha=0.5, label='P(Loss > VaR)')

    # Кривая плотности
    plt.plot(x_grid, y_kde, color='blue', linewidth=2, label="Щільність")

    # Линия VaR
    plt.axvline(var, color="red", linestyle="--", linewidth=2, label=f"VaR 99% = {var:,.2f}")
    # Линия ES
    plt.axvline(es, color="orange", linestyle="--", linewidth=2, label=f"ES = {es:,.2f}")

    # Найти максимум плотности
    y_max = np.max(y_kde)
    x_span = x_grid[-1] - x_grid[0]

    # Подпись для VaR
    plt.text(var - 0.01 * x_span, y_max * 0.95, f'VaR\n{var:,.0f}',
         color='red', fontsize=12, fontweight='bold', va='bottom', ha='right')
    # Подпись для ES
    plt.text(es + 0.01 * x_span, y_max * 0.80, f'ES\n{es:,.0f}',
         color='orange', fontsize=12, fontweight='bold', va='bottom', ha='left')
    
    # # Аннотации
    # plt.annotate(f'VaR\n{var:,.0f}', xy=(var, kde(var)), xytext=(var, kde(var)*1.1),
    #              textcoords='data', ha='center', color='red', fontsize=10, fontweight='bold',
    #              arrowprops=dict(facecolor='red', arrowstyle='-|>', lw=1.5))
    # plt.annotate(f'ES\n{es:,.0f}', xy=(es, kde(es)), xytext=(es, kde(es)*1.1),
    #              textcoords='data', ha='center', color='orange', fontsize=10, fontweight='bold',
    #              arrowprops=dict(facecolor='orange', arrowstyle='-|>', lw=1.5))

    plt.xlabel('Збитки (Loss)')
    plt.ylabel('Щільність імовірності')
    plt.title('Розподіл збитків з VaR та ES')
    plt.legend()
    plt.xlim(left=-1.2e6, right=x_max)
    plt.tight_layout()
    plt.savefig(img_path)
    plt.close()
    return img_path

def insert_image_to_excel(img_path, sheet_name='Backtesting', cell='R12', width=500, height=350):
    wb = xw.Book.caller()
    ws = wb.sheets[sheet_name]
    for pic in ws.pictures:
        if pic.name == 'VaR_ES_Distribution':
            pic.delete()
    pic = ws.pictures.add(img_path, name='VaR_ES_Distribution', update=True,
                          left=ws.range(cell).left, top=ws.range(cell).top)
    pic.width = width
    pic.height = height

def paste_plot_var_es():
    sheet = 'LN_ES'
    col = 'BG'
    cell_img = 'BN8'
    tmp_img = os.path.join(os.path.expanduser("~"), 'distribution.png')
    
    wb = xw.Book.caller()
    data = get_loss_data(sheet, col)
    var = get_named_value(wb, 'Value_VaR')
    es = get_named_value(wb, 'Value_ES')
    img_path = plot_distribution(data, var, es, tmp_img)
    insert_image_to_excel(img_path, sheet, cell_img)

if __name__ == "__main__":
    xw.Book('ИМЯ_ТВОЕГО_ФАЙЛА.xlsx').set_mock_caller()
    paste_plot_var_es()
