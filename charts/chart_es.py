import xlwings as xw
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import os
import sys
from scipy.stats import gaussian_kde

# Для корректного вывода в консоли Windows
if sys.platform == "win32":
    os.system("chcp 65001 > NUL")
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# =============================================================================
# КОНФИГУРАЦИЯ - ВСЕ ПАРАМЕТРЫ ЗАДАЮТСЯ ЗДЕСЬ
# =============================================================================
class Config:
    """Централизованная конфигурация всех параметров"""
    
    # Параметры данных
    DATA_SHEET = 'Scenario_Level_ES'  # Лист с данными о потерях
    DATA_COLUMN = 'AV'                # Колонка с данными о потерях
    DATA_START_ROW = 7                # Начальная строка данных
    
    # Именованные ячейки для VaR и ES
    VAR_NAMED_CELL = 'Value_VaR'   # Именованная ячейка с VaR
    ES_NAMED_CELL = 'Value_ES'     # Именованная ячейка с ES
    
    # Параметры вставки изображения
    IMAGE_SHEET = 'Верифікація моделі ES'    # Лист для вставки изображения (можете изменить)
    IMAGE_CELL = 'D7'             # Ячейка для вставки изображения
    IMAGE_WIDTH = 500              # Ширина изображения в пикселях
    IMAGE_HEIGHT = 160             # Высота изображения в пикселях
    IMAGE_NAME = 'VaR_ES_Distribution'  # Имя изображения в Excel
    
    # Параметры файла изображения
    TEMP_IMAGE_NAME = 'distribution.png'  # Имя временного файла
    IMAGE_DPI = 200                # Качество изображения
    
    # Параметры графика
    FIGURE_SIZE = (19.5, 5)           # Размер графика (ширина, высота)
    X_MIN_LIMIT = -1.0e6           # Минимальная граница по X
    BINS_COUNT = 30                # Количество столбцов гистограммы (если используется)

# =============================================================================
# ФУНКЦИИ РАБОТЫ С ДАННЫМИ
# =============================================================================

def get_loss_data():
    """
    Получить данные о потерях из Excel согласно конфигурации.
    Все параметры берутся из Config.
    """
    wb = xw.Book.caller()
    ws = wb.sheets[Config.DATA_SHEET]
    
    # Найти последнюю строку с данными
    last_row = ws.range(f'{Config.DATA_COLUMN}1048576').end('up').row
    
    # Получаем данные без принудительного преобразования в float
    raw_data = ws.range(f'{Config.DATA_COLUMN}{Config.DATA_START_ROW}:{Config.DATA_COLUMN}{last_row}').value
    
    # Если данные одномерные, превращаем в список
    if not isinstance(raw_data, (list, tuple)):
        raw_data = [raw_data]
    elif any(isinstance(item, (list, tuple)) for item in raw_data):
        # Если данные двумерные, делаем их плоскими
        flat_data = []
        for item in raw_data:
            if isinstance(item, (list, tuple)):
                flat_data.extend(item)
            else:
                flat_data.append(item)
        raw_data = flat_data
    
    # Фильтруем только числовые значения
    numeric_data = []
    for item in raw_data:
        if item is not None:
            try:
                if isinstance(item, (int, float)):
                    if not np.isnan(item) and item != 0:
                        numeric_data.append(float(item))
                elif isinstance(item, str):
                    # Пытаемся преобразовать строку в число
                    try:
                        num_val = float(item)
                        if not np.isnan(num_val) and num_val != 0:
                            numeric_data.append(num_val)
                    except ValueError:
                        # Пропускаем строки, которые нельзя преобразовать в число
                        continue
            except (ValueError, TypeError):
                continue
    
    print(f"Прочитано данных: {len(raw_data)} ячеек, из них числовых: {len(numeric_data)}")
    return np.array(numeric_data)

def get_var_es_values():
    """
    Получить значения VaR и ES из именованных ячеек согласно конфигурации.
    """
    wb = xw.Book.caller()
    var = float(wb.names[Config.VAR_NAMED_CELL].refers_to_range.value)
    es = float(wb.names[Config.ES_NAMED_CELL].refers_to_range.value)
    return var, es

# =============================================================================
# ФУНКЦИИ ПОСТРОЕНИЯ ГРАФИКА
# =============================================================================

def create_distribution_plot(data, var, es):
    """
    Создать график распределения с VaR и ES.
    Параметры графика берутся из Config.
    """
    if len(data) == 0:
        raise ValueError("Нет данных для построения графика")
    
    # Настройка стиля
    sns.set_theme(style="whitegrid")
    plt.figure(figsize=Config.FIGURE_SIZE)

    # Определение границ графика
    x_min = max(Config.X_MIN_LIMIT, np.min(data))
    x_max = max(np.max(data), es) * 1
    x_grid = np.linspace(x_min, x_max, 1000)

    # Построение KDE
    kde = gaussian_kde(data)
    y_kde = kde(x_grid)

    # Индекс для VaR
    idx_var = np.searchsorted(x_grid, var)

    # Закраска областей
    plt.fill_between(x_grid[:idx_var], 0, y_kde[:idx_var], 
                     color='skyblue', alpha=0.6, label='P(Loss ≤ VaR)')
    plt.fill_between(x_grid[idx_var:], 0, y_kde[idx_var:], 
                     color='pink', alpha=0.5, label='P(Loss > VaR)')

    # Кривая плотности
    plt.plot(x_grid, y_kde, color='blue', linewidth=3, label="Щільність")

    # Вертикальные линии
    plt.axvline(var, color="red", linestyle="--", linewidth=2, 
                label=f"VaR 99% = {var:,.2f}")
    plt.axvline(es, color="orange", linestyle="--", linewidth=2, 
                label=f"ES = {es:,.2f}")

    # Подписи
    y_max = np.max(y_kde)
    x_span = x_grid[-1] - x_grid[0]
    
    plt.text(var - 0.01 * x_span, y_max * 0.95, f'VaR\n{var:,.0f}',
             color='red', fontsize=16, fontweight='bold', va='bottom', ha='right')
    plt.text(es + 0.01 * x_span, y_max * 0.80, f'ES\n{es:,.0f}',
             color='orange', fontsize=16, fontweight='bold', va='bottom', ha='left')

    # Оформление
    plt.xlabel('Збитки (Loss)', fontsize=14)
    plt.ylabel('Щільність імовірності', fontsize=14)
    #plt.title('Розподіл збитків з VaR та ES', fontsize=18)
    plt.legend(fontsize=14)
    plt.xlim(left=x_min, right=x_max)
    plt.tight_layout()
    
    return plt

def save_plot(plt_obj, filepath):
    """Сохранить график в файл с параметрами из конфигурации"""
    plt_obj.savefig(filepath, dpi=Config.IMAGE_DPI, bbox_inches='tight')
    plt_obj.close()
    return filepath

# =============================================================================
# ФУНКЦИИ РАБОТЫ С EXCEL
# =============================================================================

def insert_image_to_excel(img_path):
    """
    Вставить изображение в Excel согласно конфигурации.
    Все параметры берутся из Config.
    """
    wb = xw.Book.caller()
    ws = wb.sheets[Config.IMAGE_SHEET]
    
    # Удаляем предыдущее изображение
    for pic in ws.pictures:
        if pic.name == Config.IMAGE_NAME:
            pic.delete()
    
    # Вставляем новое изображение
    pic = ws.pictures.add(img_path, name=Config.IMAGE_NAME, update=True,
                          left=ws.range(Config.IMAGE_CELL).left, 
                          top=ws.range(Config.IMAGE_CELL).top)
    pic.width = Config.IMAGE_WIDTH
    pic.height = Config.IMAGE_HEIGHT

# =============================================================================
# ОСНОВНАЯ ФУНКЦИЯ
# =============================================================================

def create_var_es_plot():
    """
    Основная функция для создания и вставки графика VaR/ES.
    Все параметры берутся из класса Config.
    """
    try:
        print("=== СОЗДАНИЕ ГРАФИКА VaR/ES ===")
        print(f"Источник данных: лист '{Config.DATA_SHEET}', колонка '{Config.DATA_COLUMN}'")
        
        # Получение данных
        print("1. Получение данных о потерях...")
        data = get_loss_data()
        
        if len(data) == 0:
            raise ValueError("Не найдено числовых данных для построения графика")
        
        # Получение VaR и ES
        print("2. Получение значений VaR и ES...")
        var, es = get_var_es_values()
        print(f"   VaR = {var:,.2f}")
        print(f"   ES = {es:,.2f}")
        
        # Создание графика
        print("3. Построение графика...")
        plt_obj = create_distribution_plot(data, var, es)
        
        # Сохранение в файл
        print("4. Сохранение графика...")
        temp_path = os.path.join(os.path.expanduser("~"), Config.TEMP_IMAGE_NAME)
        img_path = save_plot(plt_obj, temp_path)
        
        # Вставка в Excel
        print("5. Вставка в Excel...")
        print(f"   Лист: '{Config.IMAGE_SHEET}', ячейка: '{Config.IMAGE_CELL}'")
        insert_image_to_excel(img_path)
        
        print("✅ График успешно создан и вставлен!")
        
    except Exception as e:
        print(f"❌ Ошибка: {str(e)}")
        raise

# =============================================================================
# ОТЛАДОЧНЫЕ ФУНКЦИИ
# =============================================================================

def debug_configuration():
    """Показать текущую конфигурацию"""
    print("=== ТЕКУЩАЯ КОНФИГУРАЦИЯ ===")
    print(f"Данные:")
    print(f"  Лист: {Config.DATA_SHEET}")
    print(f"  Колонка: {Config.DATA_COLUMN}")
    print(f"  Начальная строка: {Config.DATA_START_ROW}")
    print(f"VaR/ES:")
    print(f"  VaR ячейка: {Config.VAR_NAMED_CELL}")
    print(f"  ES ячейка: {Config.ES_NAMED_CELL}")
    print(f"Изображение:")
    print(f"  Лист: {Config.IMAGE_SHEET}")
    print(f"  Ячейка: {Config.IMAGE_CELL}")
    print(f"  Размер: {Config.IMAGE_WIDTH}x{Config.IMAGE_HEIGHT}")

if __name__ == "__main__":
    try:
        debug_configuration()
        print()
        create_var_es_plot()
    except Exception as e:
        print(f"Критическая ошибка: {e}")

# =============================================================================
# ФУНКЦИЯ ДЛЯ ОБРАТНОЙ СОВМЕСТИМОСТИ
# =============================================================================

def paste_plot_var_es():
    """
    Функция для обратной совместимости.
    Просто вызывает новую основную функцию.
    """
    create_var_es_plot()