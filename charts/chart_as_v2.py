# =============================================================================
# ДВОЙНОЙ СПИДОМЕТР ДЛЯ EXCEL - AS_Z2_1d И AS_Z2_10d
# =============================================================================

import xlwings as xw
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import os

# =============================================================================
# КОНФИГУРАЦИЯ - ВСЕ ПАРАМЕТРЫ ЗАДАЮТСЯ ЗДЕСЬ
# =============================================================================

class Config:
    """Централизованная конфигурация двойного спидометра AS (Acerbi-Szekely Z2)"""
    
    # === ИСТОЧНИКИ ДАННЫХ ===
    SOURCE_CELL_1D = 'AS_Z2_1d'              # Первая именованная ячейка (1 день)
    SOURCE_CELL_10D = 'AS_Z2_10d'            # Вторая именованная ячейка (10 дней)
    
    # === РАЗМЕЩЕНИЕ В EXCEL ===
    TARGET_SHEET = 'Верифікація моделі ES'   # Куда вставлять диаграмму лист
    TARGET_CELL = 'D66'                      # Куда вставлять диаграмму
    IMAGE_NAME = 'DoubleSpeedometerChart'    # Имя изображения в Excel
    IMAGE_WIDTH = 250                        # Увеличенная ширина изображения в Excel
    IMAGE_HEIGHT = 140                       # Увеличенная высота изображения в Excel
    
    # === ЗАГОЛОВКИ ===
    TITLE_1D = 'Acerbi-Szekely Z2 (1 день)'     # Заголовок первого спидометра
    TITLE_10D = 'Acerbi-Szekely Z2 (10 днів)'   # Заголовок второго спидометра
    
    # === ГРАНИЦЫ ЦВЕТОВЫХ ЗОН ===
    GREEN_MIN = 0.75                          # Начало зеленой зоны
    GREEN_MAX = 1.25                          # Конец зеленой зоны
    YELLOW1_MIN = 0.60                        # Начало первой желтой зоны
    YELLOW1_MAX = 0.75                        # Конец первой желтой зоны
    YELLOW2_MIN = 1.25                        # Начало второй желтой зоны
    YELLOW2_MAX = 1.40                        # Конец второй желтой зоны
    RED1_MIN = 0.4                            # Начало первой красной зоны
    RED1_MAX = 0.60                           # Конец первой красной зоны
    RED2_MIN = 1.40                           # Начало второй красной зоны
    RED2_MAX = 1.6                            # Конец второй красной зоны
    
    # === ПАРАМЕТРЫ ШКАЛЫ ===
    SCALE_MIN = 0.4                           # Минимум шкалы
    SCALE_MAX = 1.6                           # Максимум шкалы
    
    # === НАСТРОЙКИ ВНЕШНЕГО ВИДА ===
    FIGURE_WIDTH = 10                         # Увеличенная ширина для двух спидометров
    FIGURE_HEIGHT = 6                         # Увеличенная высота
    DPI = 150                                 # Качество изображения
    
    # === НАСТРОЙКИ РАЗМЕЩЕНИЯ ДВУХ СПИДОМЕТРОВ ===
    SPEEDOMETER_SPACING = 2.1                 # Расстояние между центрами спидометров
    
    # === ПАРАМЕТРЫ ОБРЕЗКИ ПУСТОГО ПРОСТРАНСТВА ===
    BOTTOM_MARGIN = -0.2                      # Нижняя граница (было -0.8, стало -0.4)
    TOP_MARGIN = 1.5                          # Верхняя граница
    SIDE_MARGIN = 1.1                         # Боковые отступы
    
    # === ПАРАМЕТРЫ ЦЕНТРАЛЬНОГО КРУГА ===
    CENTER_CIRCLE_RADIUS = 0.18               # Радиус центрального круга
    CENTER_FONT_SIZE = 14                     # Размер шрифта в центре
    TITLE_FONT_SIZE = 16                      # Размер шрифта заголовков
    
    # === ЦВЕТА ЗОН ===
    RED_COLOR = '#FF4444'
    YELLOW_COLOR = '#FFD700'
    GREEN_COLOR = '#00B050'


# =============================================================================
# ОСНОВНЫЕ ФУНКЦИИ
# =============================================================================

def create_single_speedometer(ax, current_value, title, x_offset=0):
    '''Создает один спидометр на заданном subplot'''
    
    # Функция для перевода значения в угол
    def value_to_angle(value):
        return 180 - 180 * (value - Config.SCALE_MIN) / (Config.SCALE_MAX - Config.SCALE_MIN)
    
    # Рисуем цветовые зоны
    zones = [
        (Config.RED1_MIN, Config.RED1_MAX, Config.RED_COLOR),
        (Config.YELLOW1_MIN, Config.YELLOW1_MAX, Config.YELLOW_COLOR),
        (Config.GREEN_MIN, Config.GREEN_MAX, Config.GREEN_COLOR),
        (Config.YELLOW2_MIN, Config.YELLOW2_MAX, Config.YELLOW_COLOR),
        (Config.RED2_MIN, Config.RED2_MAX, Config.RED_COLOR)
    ]
    
    for min_val, max_val, color in zones:
        wedge = patches.Wedge(center=(x_offset, 0), r=1, 
                             theta1=value_to_angle(max_val), 
                             theta2=value_to_angle(min_val), 
                             facecolor=color, alpha=0.7)
        ax.add_patch(wedge)
    
    # Рисуем деления шкалы
    for i in range(9):
        val = Config.SCALE_MIN + (Config.SCALE_MAX - Config.SCALE_MIN) * i / 8
        angle = np.deg2rad(value_to_angle(val))
        
        # Линия деления
        r_start, r_end = 0.85, 0.95
        x_start = x_offset + r_start * np.cos(angle)
        y_start = r_start * np.sin(angle)
        x_end = x_offset + r_end * np.cos(angle)
        y_end = r_end * np.sin(angle)
        ax.plot([x_start, x_end], [y_start, y_end], 'k-', lw=2)
        
        # Подпись (показываем только на каждом втором делении для читаемости)
        if i % 2 == 0:
            label_r = 0.75
            x_label = x_offset + label_r * np.cos(angle)
            y_label = label_r * np.sin(angle)
            ax.text(x_label, y_label, f'{val:.1f}', 
                   ha='center', va='center', fontsize=10, fontweight='bold')
    
    # Стрелка
    clamped_value = np.clip(current_value, Config.SCALE_MIN, Config.SCALE_MAX)
    arrow_angle_rad = np.deg2rad(value_to_angle(clamped_value))
    arrow_x = x_offset + 0.65 * np.cos(arrow_angle_rad)
    arrow_y = 0.65 * np.sin(arrow_angle_rad)
    ax.arrow(x_offset, 0, arrow_x - x_offset, arrow_y, width=0.02, head_width=0.08, 
            head_length=0.1, fc='black', ec='black')
    
    # Центральный круг с значением
    center_circle = patches.Circle((x_offset, 0), Config.CENTER_CIRCLE_RADIUS, 
                                  facecolor='white', edgecolor='black', linewidth=2)
    ax.add_patch(center_circle)
    ax.text(x_offset, 0, f'{current_value:.2f}', ha='center', va='center', 
           fontsize=Config.CENTER_FONT_SIZE, weight='bold')
    
    # Заголовок спидометра
    ax.text(x_offset, 1.2, title, ha='center', va='center', 
           fontsize=Config.TITLE_FONT_SIZE, weight='bold')


def create_double_speedometer_plot(value_1d, value_10d):
    '''Создает и сохраняет двойной спидометр'''
    
    # Создаем фигуру с уменьшенной высотой
    fig, ax = plt.subplots(figsize=(Config.FIGURE_WIDTH, Config.FIGURE_HEIGHT), 
                          subplot_kw={'aspect': 'equal'})
    ax.axis('off')
    
    # Позиции для двух спидометров
    left_position = -Config.SPEEDOMETER_SPACING / 2
    right_position = Config.SPEEDOMETER_SPACING / 2
    
    # Создаем два спидометра
    create_single_speedometer(ax, value_1d, Config.TITLE_1D, left_position)
    create_single_speedometer(ax, value_10d, Config.TITLE_10D, right_position)
    
    # ОБРЕЗАЕМ ГРАНИЦЫ - убираем пустое пространство снизу
    ax.set_xlim(left_position - Config.SIDE_MARGIN, right_position + Config.SIDE_MARGIN)
    ax.set_ylim(Config.BOTTOM_MARGIN, Config.TOP_MARGIN)  # Используем параметры из констант
    
    # Общий заголовок
    # fig.suptitle('Спидометры AS (Expected Shortfall)', fontsize=18, fontweight='bold', y=0.92)
    
    # Сохраняем с плотной обрезкой
    temp_path = os.path.join(os.path.expanduser("~"), 'double_speedometer.png')
    fig.savefig(temp_path, dpi=Config.DPI, bbox_inches='tight', pad_inches=0.1)  # Минимальные отступы
    plt.close(fig)
    
    return temp_path


def insert_image_to_excel():
    '''Единственная функция - делает все сразу'''
    
    # Получаем данные из обеих именованных ячеек
    wb = xw.Book.caller()
    value_1d = wb.names[Config.SOURCE_CELL_1D].refers_to_range.value
    value_10d = wb.names[Config.SOURCE_CELL_10D].refers_to_range.value
    
    # Создаем двойной спидометр и получаем путь к файлу
    img_path = create_double_speedometer_plot(value_1d, value_10d)
    
    # Удаляем старое изображение
    ws = wb.sheets[Config.TARGET_SHEET]
    for pic in ws.pictures:
        if pic.name == Config.IMAGE_NAME:
            pic.delete()
    
    # Вставляем новое изображение
    pic = ws.pictures.add(img_path, name=Config.IMAGE_NAME, update=True,
                         left=ws.range(Config.TARGET_CELL).left,
                         top=ws.range(Config.TARGET_CELL).top)
    
    pic.width = Config.IMAGE_WIDTH
    pic.height = Config.IMAGE_HEIGHT
    
    print(f"✅ Двойной спидометр создан! AS_1d: {value_1d:.3f}, AS_10d: {value_10d:.3f}")


# Для совместимости
main = insert_image_to_excel

if __name__ == '__main__':
    insert_image_to_excel()