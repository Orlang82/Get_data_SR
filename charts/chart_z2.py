import xlwings as xw
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import io
from PIL import Image
import os
from pathlib import Path

# =============================================================================
# КОНСТАНТЫ И НАСТРОЙКИ
# =============================================================================

# Цветовые зоны согласно техническому заданию
GREEN_ZONE = (0.75, 1.25)      # Зелёная зона: хорошие значения
YELLOW_ZONE_1 = (0.60, 0.75)   # Жёлтая зона: нижний диапазон
YELLOW_ZONE_2 = (1.25, 1.40)   # Жёлтая зона: верхний диапазон
RED_ZONE_1 = (0, 0.60)         # Красная зона: критически низкие значения
RED_ZONE_2 = (1.40, 2.0)       # Красная зона: критически высокие значения

# Пределы шкалы
MIN_SCALE = 0.0
MAX_SCALE = 2.0

# Настройки для matplotlib (поддержка русских шрифтов)
plt.rcParams['font.family'] = ['DejaVu Sans', 'Arial', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False


# =============================================================================
# ОСНОВНЫЕ ФУНКЦИИ
# =============================================================================

def create_speedometer_figure(current_value):
    """
    Создает диаграмму спидометра в виде полукруга с цветовыми зонами.
    
    Args:
        current_value (float): Текущее значение для отображения
        
    Returns:
        matplotlib.figure.Figure: Объект фигуры matplotlib
    """
    # Создаем фигуру с соотношением сторон для полукруга
    fig, ax = plt.subplots(figsize=(8, 5), subplot_kw={'aspect': 'equal'})
    ax.axis('off')  # Убираем оси
    
    # Функция для преобразования значения в угол (180° = 0, 0° = 2)
    def value_to_angle(value):
        """Преобразует значение шкалы в угол для полукруга"""
        return 180 - 180 * (value - MIN_SCALE) / (MAX_SCALE - MIN_SCALE)
    
    # =============================================================================
    # РИСУЕМ ЦВЕТОВЫЕ ЗОНЫ
    # =============================================================================
    
    # Красные зоны (критические)
    red1 = patches.Wedge(
        center=(0, 0), r=1, 
        theta1=value_to_angle(RED_ZONE_1[1]), 
        theta2=value_to_angle(RED_ZONE_1[0]), 
        facecolor='#FF4444', alpha=0.7, linewidth=1, edgecolor='white'
    )
    red2 = patches.Wedge(
        center=(0, 0), r=1, 
        theta1=value_to_angle(RED_ZONE_2[0]), 
        theta2=value_to_angle(RED_ZONE_2[1]), 
        facecolor='#FF4444', alpha=0.7, linewidth=1, edgecolor='white'
    )
    ax.add_patch(red1)
    ax.add_patch(red2)
    
    # Жёлтые зоны (предупредительные)
    yellow1 = patches.Wedge(
        center=(0, 0), r=1, 
        theta1=value_to_angle(YELLOW_ZONE_1[1]), 
        theta2=value_to_angle(YELLOW_ZONE_1[0]), 
        facecolor='#FFD700', alpha=0.7, linewidth=1, edgecolor='white'
    )
    yellow2 = patches.Wedge(
        center=(0, 0), r=1, 
        theta1=value_to_angle(YELLOW_ZONE_2[1]), 
        theta2=value_to_angle(YELLOW_ZONE_2[0]), 
        facecolor='#FFD700', alpha=0.7, linewidth=1, edgecolor='white'
    )
    ax.add_patch(yellow1)
    ax.add_patch(yellow2)
    
    # Зелёная зона (оптимальная)
    green = patches.Wedge(
        center=(0, 0), r=1, 
        theta1=value_to_angle(GREEN_ZONE[1]), 
        theta2=value_to_angle(GREEN_ZONE[0]), 
        facecolor='#44AA44', alpha=0.7, linewidth=1, edgecolor='white'
    )
    ax.add_patch(green)
    
    # =============================================================================
    # СОЗДАЕМ ШКАЛУ С ДЕЛЕНИЯМИ
    # =============================================================================
    
    # Количество делений на шкале
    num_major_ticks = 9   # Основные деления (0, 0.25, 0.5, ..., 2.0)
    num_minor_ticks = 21  # Мелкие деления
    
    # Рисуем мелкие деления
    for i in range(num_minor_ticks):
        val = MIN_SCALE + (MAX_SCALE - MIN_SCALE) * i / (num_minor_ticks - 1)
        angle = np.deg2rad(value_to_angle(val))
        
        # Координаты для мелких делений
        r_start = 0.90
        r_end = 0.95
        x_start = r_start * np.cos(angle)
        y_start = r_start * np.sin(angle)
        x_end = r_end * np.cos(angle)
        y_end = r_end * np.sin(angle)
        
        ax.plot([x_start, x_end], [y_start, y_end], 
               color='black', lw=1, solid_capstyle='round')
    
    # Рисуем основные деления с подписями
    for i in range(num_major_ticks):
        val = MIN_SCALE + (MAX_SCALE - MIN_SCALE) * i / (num_major_ticks - 1)
        angle = np.deg2rad(value_to_angle(val))
        
        # Координаты для основных делений
        r_start = 0.85
        r_end = 0.95
        x_start = r_start * np.cos(angle)
        y_start = r_start * np.sin(angle)
        x_end = r_end * np.cos(angle)
        y_end = r_end * np.sin(angle)
        
        # Рисуем основное деление
        ax.plot([x_start, x_end], [y_start, y_end], 
               color='black', lw=2, solid_capstyle='round')
        
        # Добавляем числовую подпись
        label_r = 0.75
        x_label = label_r * np.cos(angle)
        y_label = label_r * np.sin(angle)
        ax.text(x_label, y_label, f'{val:.2f}', 
               ha='center', va='center', fontsize=11, 
               fontweight='bold', color='black')
    
    # =============================================================================
    # ДОБАВЛЯЕМ ПОДПИСИ ЗОН
    # =============================================================================
    
    ax.text(-0.7, -0.2, 'КРИТИЧЕСКАЯ\nЗОНА', color='#CC0000', 
           fontsize=9, weight='bold', ha='center', va='center')
    ax.text(-0.25, -0.35, 'ПРЕДУПРЕДИТЕЛЬНАЯ\nЗОНА', color='#B8860B', 
           fontsize=9, weight='bold', ha='center', va='center')
    ax.text(0.25, -0.35, 'ОПТИМАЛЬНАЯ\nЗОНА', color='#228B22', 
           fontsize=9, weight='bold', ha='center', va='center')
    ax.text(0.7, -0.2, 'КРИТИЧЕСКАЯ\nЗОНА', color='#CC0000', 
           fontsize=9, weight='bold', ha='center', va='center')
    
    # =============================================================================
    # РИСУЕМ СТРЕЛКУ-УКАЗАТЕЛЬ
    # =============================================================================
    
    # Ограничиваем значение в пределах шкалы
    clamped_value = np.clip(current_value, MIN_SCALE, MAX_SCALE)
    current_angle = value_to_angle(clamped_value)
    arrow_angle_rad = np.deg2rad(current_angle)
    
    # Координаты конца stрелки
    arrow_length = 0.65
    arrow_x = arrow_length * np.cos(arrow_angle_rad)
    arrow_y = arrow_length * np.sin(arrow_angle_rad)
    
    # Рисуем стрелку
    ax.arrow(0, 0, arrow_x, arrow_y, 
            width=0.02, head_width=0.08, head_length=0.1, 
            fc='black', ec='black', length_includes_head=True)
    
    # =============================================================================
    # ДОБАВЛЯЕМ ЦЕНТРАЛЬНЫЙ КРУГ С ЗНАЧЕНИЕМ
    # =============================================================================
    
    # Центральный круг
    center_circle = patches.Circle((0, 0), 0.12, 
                                  facecolor='white', edgecolor='black', 
                                  linewidth=2, zorder=10)
    ax.add_patch(center_circle)
    
    # Отображаем текущее значение в центре
    ax.text(0, 0, f'{current_value:.3f}', 
           ha='center', va='center', fontsize=16, 
           weight='bold', color='black', zorder=11)
    
    # =============================================================================
    # ДОБАВЛЯЕМ ЗАГОЛОВОК И ФИНАЛЬНЫЕ НАСТРОЙКИ
    # =============================================================================
    
    # Заголовок
    ax.text(0, 1.15, 'СПИДОМЕТР ПОКАЗАТЕЛЕЙ', 
           ha='center', va='center', fontsize=16, 
           weight='bold', color='black')
    
    # Устанавливаем пределы отображения
    ax.set_xlim(-1.3, 1.3)
    ax.set_ylim(-0.6, 1.3)
    
    return fig


def save_figure_to_image_buffer(fig):
    """
    Сохраняет matplotlib фигуру в буфер памяти как PNG изображение.
    
    Args:
        fig: matplotlib figure объект
        
    Returns:
        io.BytesIO: Буфер с изображением в формате PNG
    """
    buf = io.BytesIO()
    fig.savefig(buf, format='png', bbox_inches='tight', 
               dpi=150, facecolor='white', edgecolor='none')
    buf.seek(0)
    return buf


def get_zone_status(value):
    """
    Определяет статус значения по зонам.
    
    Args:
        value (float): Значение для анализа
        
    Returns:
        str: Статус зоны ('GREEN', 'YELLOW', 'RED')
    """
    if GREEN_ZONE[0] <= value <= GREEN_ZONE[1]:
        return 'GREEN'
    elif (YELLOW_ZONE_1[0] <= value <= YELLOW_ZONE_1[1] or 
          YELLOW_ZONE_2[0] <= value <= YELLOW_ZONE_2[1]):
        return 'YELLOW'
    else:
        return 'RED'


def update_speedometer():
    """
    Основная функция для обновления спидометра в Excel.
    Читает значение из ячейки Z2 и создает/обновляет диаграмму.
    """
    try:
        # Получаем активную книгу Excel
        wb = xw.Book.caller()
        sheet = wb.sheets[0]  # Используем первый лист
        
        # =============================================================================
        # ЧТЕНИЕ И ВАЛИДАЦИЯ ДАННЫХ
        # =============================================================================
        
        # Читаем значение из ячейки Z2
        cell_value = sheet.range('Z2').value
        
        # Проверяем, что ячейка не пустая
        if cell_value is None:
            raise ValueError('Ячейка Z2 пуста. Введите числовое значение.')
        
        # Преобразуем в число
        try:
            current_value = float(cell_value)
        except (ValueError, TypeError):
            raise ValueError(f'Некорректное значение в ячейке Z2: "{cell_value}". Ожидается число.')
        
        # Проверяем диапазон (предупреждение, но не блокируем)
        if not (MIN_SCALE <= current_value <= MAX_SCALE):
            print(f'Предупреждение: Значение {current_value} выходит за пределы шкалы [{MIN_SCALE}, {MAX_SCALE}]')
        
        # =============================================================================
        # СОЗДАНИЕ И СОХРАНЕНИЕ ДИАГРАММЫ
        # =============================================================================
        
        # Создаем диаграмму спидометра
        fig = create_speedometer_figure(current_value)
        
        # Сохраняем фигуру в буфер памяти
        image_buffer = save_figure_to_image_buffer(fig)
        
        # Открываем изображение через PIL
        img = Image.open(image_buffer)
        
        # =============================================================================
        # ВСТАВКА ИЗОБРАЖЕНИЯ В EXCEL
        # =============================================================================
        
        # Удаляем предыдущие изображения спидометра
        for pic in sheet.pictures:
            if 'Speedometer' in pic.name:
                pic.delete()
        
        # Вставляем новое изображение в ячейку B2
        target_range = sheet.range('B2')
        sheet.pictures.add(
            img, 
            left=target_range.left, 
            top=target_range.top, 
            name='SpeedometerChart',
            update=True
        )
        
        # =============================================================================
        # ДОПОЛНИТЕЛЬНАЯ ИНФОРМАЦИЯ В EXCEL
        # =============================================================================
        
        # Записываем статус в соседнюю ячейку
        zone_status = get_zone_status(current_value)
        status_messages = {
            'GREEN': 'ОПТИМАЛЬНО',
            'YELLOW': 'ПРЕДУПРЕЖДЕНИЕ', 
            'RED': 'КРИТИЧНО'
        }
        
        status_colors = {
            'GREEN': (0, 128, 0),    # Зелёный
            'YELLOW': (255, 165, 0), # Оранжевый
            'RED': (255, 0, 0)       # Красный
        }
        
        # Записываем статус в ячейку A2
        status_cell = sheet.range('A2')
        status_cell.value = f'Статус: {status_messages[zone_status]}'
        status_cell.font.color = status_colors[zone_status]
        status_cell.font.bold = True
        
        # Записываем текущее время обновления в A1
        from datetime import datetime
        sheet.range('A1').value = f'Обновлено: {datetime.now().strftime("%d.%m.%Y %H:%M:%S")}'
        
        # Закрываем фигуру для освобождения памяти
        plt.close(fig)
        
        print(f'Спидометр успешно обновлен. Значение: {current_value}, Статус: {status_messages[zone_status]}')
        
    except Exception as e:
        # =============================================================================
        # ОБРАБОТКА ОШИБОК
        # =============================================================================
        error_message = f'Ошибка при создании спидометра: {str(e)}'
        print(error_message)
        
        try:
            # Пытаемся записать ошибку в Excel
            wb = xw.Book.caller()
            sheet = wb.sheets[0]
            error_cell = sheet.range('A1')
            error_cell.value = error_message
            error_cell.font.color = (255, 0, 0)  # Красный цвет
            error_cell.font.bold = True
        except:
            # Если не удается записать в Excel, выводим в консоль
            pass


# =============================================================================
# ФУНКЦИИ ДЛЯ АВТОМАТИЧЕСКОГО ОБНОВЛЕНИЯ
# =============================================================================

def setup_auto_update():
    """
    Настраивает автоматическое обновление спидометра при изменении ячейки Z2.
    """
    try:
        wb = xw.Book.caller()
        sheet = wb.sheets[0]
        
        # Добавляем обработчик изменения ячейки (требует VBA макрос)
        vba_code = '''
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address = "$Z$2" Then
        Application.Run "update_speedometer"
    End If
End Sub
        '''
        
        print("Для автоматического обновления добавьте следующий VBA код в лист:")
        print(vba_code)
        
    except Exception as e:
        print(f'Ошибка настройки автообновления: {e}')


# =============================================================================
# ОСНОВНАЯ ФУНКЦИЯ (ТОЧКА ВХОДА)
# =============================================================================

def main():
    """
    Главная функция - точка входа для xlwings.
    Вызывается при нажатии кнопки "Run main" в xlwings или через RunPython.
    """
    update_speedometer()


# =============================================================================
# ЗАПУСК КАК ОТДЕЛЬНОГО СКРИПТА
# =============================================================================

if __name__ == '__main__':
    """
    Блок для запуска скрипта независимо от Excel.
    Полезно для тестирования и отладки.
    """
    try:
        # Проверяем, запущен ли скрипт из Excel
        try:
            wb = xw.Book.caller()
            main()
        except:
            # Если не из Excel, то создаем тестовый сценарий
            print("Скрипт запущен вне Excel. Создаем тестовую визуализацию...")
            
            # Тестовое значение
            test_value = 1.1
            
            # Создаем и показываем диаграмму
            fig = create_speedometer_figure(test_value)
            plt.show()
            
            # Опционально сохраняем в файл
            save_path = Path.cwd() / 'speedometer_test.png'
            fig.savefig(save_path, dpi=150, bbox_inches='tight')
            print(f'Тестовая диаграмма сохранена в: {save_path}')
            
            plt.close(fig)
            
    except Exception as e:
        print(f'Ошибка выполнения: {e}')
