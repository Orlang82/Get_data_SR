"""
chart_7s_mrrr.py
~~~~~~~~~~~~~~~~
График «Динаміка мінімального розміру ринкового ризику».

Источник данных : таблица tDB_History_2 на листе sys (Excel / xlwings)
Назначение      : столбчатая гистограмма МРРР + три линии рисков
Вывод           : изображение PNG, вставляется в лист To_Report ячейка C17
"""

import os
import sys
import logging

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import xlwings as xw

# Корректный вывод кириллицы в консоли Windows
if sys.platform == "win32":
    os.system("chcp 65001 > NUL")
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")


# =============================================================================
# ЛОГИРОВАНИЕ
# =============================================================================

ENABLE_LOGGING = True


def _setup_logger() -> logging.Logger:
    """Настраивает файловый логгер; при ENABLE_LOGGING=False возвращает заглушку."""
    if not ENABLE_LOGGING:
        return logging.getLogger("chart_7s_mrrr_disabled")

    logger = logging.getLogger("chart_7s_mrrr")
    if logger.handlers:          # защита от дублирования при перезагрузке модуля
        return logger

    logger.setLevel(logging.INFO)
    log_dir = os.path.abspath(
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "logs")
    )
    os.makedirs(log_dir, exist_ok=True)

    handler = logging.FileHandler(
        os.path.join(log_dir, "chart_7s_mrrr.log"), encoding="utf-8"
    )
    handler.setFormatter(
        logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    )
    logger.addHandler(handler)
    return logger


logger = _setup_logger()


# =============================================================================
# КОНФИГУРАЦИЯ
# =============================================================================

class Config:
    """
    Централизованные настройки графика «Динаміка мінімального розміру ринкового ризику».
    
    Класс содержит все параметры, необходимые для:
    - чтения данных из Excel
    - построения визуализации (matplotlib)
    - сохранения и вставки изображения обратно в Excel
    
    Все константы сгруппированы по функциональным блокам для удобства навигации.
    """
    
    # ==========================================================================
    # ИСТОЧНИК ДАННЫХ
    # ==========================================================================
    # Параметры подключения к таблице Excel через xlwings
    
    DATA_SHEET  = "sys"           # Лист Excel с исходными данными
    DATA_TABLE  = "tDB_History_2" # Имя таблицы (ListObject) на листе
    COL_DATE    = "Дата"          # Столбец с датами наблюдений
    COL_MRRR    = "МРРР"          # Столбец с значениями МРРР (основные столбцы)
    COL_VAL     = "Валютний ризик"    # Столбец с значениями валютного риска
    COL_PCT     = "Процентний ризик"  # Столбец с значениями процентного риска
    COL_TOVAR   = "Товарний ризик"    # Столбец с значениями товарного риска
    N_LAST      = 12                # Количество последних наблюдений для отображения

    # ==========================================================================
    # ВЫВОД ИЗОБРАЖЕНИЯ В EXCEL
    # ==========================================================================
    # Параметры размещения готового PNG-файла в рабочей книге
    
    IMAGE_SHEET  = "To_Report"         # Лист Excel для вставки графика
    IMAGE_CELL   = "C18"               # Ячейка-якорь для позиционирования
    IMAGE_NAME   = "Chart_MarketRisk_7S" # Имя объекта Picture в Excel (для удаления/замены)
    TEMP_IMAGE   = "market_risk_7s.png"  # Имя временного PNG-файла в домашней директории
    IMAGE_WIDTH  = 720                   # Ширина изображения в Excel (пиксели)
    IMAGE_HEIGHT = 200                   # Высота изображения в Excel (пиксели)
    IMAGE_DPI    = 150                   # Разрешение при сохранении PNG (dots per inch)

    # ==========================================================================
    # ФИГУРА (MATPLOTLIB)
    # ==========================================================================
    # Общие параметры холста и фона
    
    FIGURE_SIZE = (8.0, 3.5)  # Размер фигуры (дюймы): (ширина, высота)
    BG_COLOR    = "#FFFFFF"   # Цвет фона фигуры (hex)

    # ==========================================================================
    # СТОЛБЦЫ МРРР
    # ==========================================================================
    # Визуальные параметры основных столбцов гистограммы
    
    BAR_COLOR = "#1F3864"  # Цвет заливки столбцов (тёмно-синий)
    BAR_WIDTH = 0.75        # Ширина одного столбца (в долях единицы по оси X)
    BAR_ALPHA = 1.0        # Прозрачность столбцов (1.0 = полностью непрозрачный)

    # ==========================================================================
    # ЛИНИИ РИСКОВ
    # ==========================================================================
    # Параметры отображения трёх линий: Валютний, Процентний, Товарний ризик
    
    LINE_VAL_COLOR   = "#00B050"   # Цвет линии валютного риска (зелёный)
    LINE_PCT_COLOR   = "#F85639"   # Цвет линии процентного риска (жёлтый)
    LINE_TOVAR_COLOR = "#ED7D31"   # Цвет линии товарного риска (оранжевый)
    LINE_WIDTH       = 2.0         # Толщина всех линий (пункты)
    MARKER_SIZE      = 4           # Размер маркеров на точках данных

    # Стили маркеров для каждой линии ('o' = круг, 's' = квадрат, '^' = треугольник вверх)
    MARKER_STYLE_VAL   = "o"       # Маркер валютного риска
    MARKER_STYLE_PCT   = "o"       # Маркер процентного риска
    MARKER_STYLE_TOVAR = "o"       # Маркер товарного риска

    # ==========================================================================
    # ПОДПИСИ ДАННЫХ
    # ==========================================================================
    # Параметры текстовых подписей значений на графике
    
    # Общие настройки шрифта для всех подписей
    LABEL_FONTWEIGHT      = "normal"  # Насыщенность шрифта ('normal', 'bold')
    LABEL_FONTSTYLE       = "normal"  # Начертание ('normal', 'italic')

    # Размер шрифта задаётся индивидуально для каждого ряда подписей
    LABEL_FONTSIZE_BAR   = 8   # Подписи столбцов МРРР
    LABEL_FONTSIZE_VAL   = 6   # Подписи Валютний ризик
    LABEL_FONTSIZE_PCT   = 6   # Подписи Процентний ризик
    LABEL_FONTSIZE_TOVAR = 6   # Подписи Товарний ризик

    LABEL_FONTWEIGHT_MRRR = "bold"    # Насыщенность шрифта для подписей МРРР
    
    # Цвета подписей (подбираются для контраста с фоном линии/столбца)
    LABEL_COLOR_LINES = "#1F3864"     # Подписи линий по умолчанию (тёмно-синий)
    LABEL_COLOR_VAL   = "#FFFFFF"     # Подписи на зелёной линии (белый)
    LABEL_COLOR_PCT   = "#FFFFFF"     # Подписи на жёлтой линии (тёмно-синий)
    LABEL_COLOR_TOVAR = "#FFFFFF"     # Подписи на оранжевой линии (белый)

    # Отступы подписей от базовой точки (доля от y_max основной оси)
    LABEL_OFFSET_Y     = 0.04  # Отступ для подписей МРРР
    LABEL_OFFSET_VAL   = 0.02   # Отступ для подписей валютного риска
    LABEL_OFFSET_PCT   = -0.10   # Отступ для подписей процентного риска 
    LABEL_OFFSET_TOVAR = 0.02   # Отступ для подписей товарного риска

    # Положение подписей относительно точки данных
    # 'top' — над точкой, 'bottom' — под точкой
    LABEL_POSITION_VAL   = "bottom"     # Положение подписей валютного риска
    LABEL_POSITION_PCT   = "bottom"  # Положение подписей процентного риска
    LABEL_POSITION_TOVAR = "bottom"     # Положение подписей товарного риска

    # ==========================================================================
    # ЗАГОЛОВОК
    # ==========================================================================
    # Параметры заголовка графика
    
    TITLE_TEXT       = "Динаміка мінімального розміру ринкового ризику"
    TITLE_FONTSIZE   = 12          # Размер шрифта заголовка (пункты)
    TITLE_FONTWEIGHT = "bold"      # Насыщенность шрифта заголовка
    TITLE_COLOR      = "#1F3864"   # Цвет текста заголовка
    TITLE_PAD        = 7          # Отступ заголовка от диаграммы (пункты)

    # ==========================================================================
    # ОСЬ X
    # ==========================================================================
    # Параметры отображения оси времени (даты)
    
    XAXIS_DATE_FORMAT = "%d.%m.%y"  # Формат дат на оси X (strftime)
    XAXIS_FONTSIZE    = 7.5           # Размер шрифта подписей дат (пункты)
    XAXIS_FONTWEIGHT  = "bold"      # Насыщенность шрифта подписей дат
    XAXIS_FONTSTYLE   = "normal"    # Начертание шрифта подписей дат
    XAXIS_COLOR       = "#1F3864"   # Цвет оси и подписей (тёмно-синий)
    XAXIS_LABEL_PAD   = 6           # Отступ подписей дат от оси X (пункты)

    # ==========================================================================
    # ОСИ Y
    # ==========================================================================
    # Параметры основной и вторичной осей Y
    
    YAXIS_VISIBLE         = False   # Показать/скрыть основную ось Y (левую)
    YAXIS2_VISIBLE        = False   # Показать/скрыть вторичную ось Y (правую)
    YAXIS2_TOP_MULTIPLIER = 9     # Множитель для верхней границы вторичной оси
                                    # (y_max_tovar × множитель = верхняя граница)

    # ==========================================================================
    # ЛЕГЕНДА
    # ==========================================================================
    # Параметры легенды (ключ к графикам)

    LEGEND_FONTSIZE   = 8             # Размер шрифта легенды (пункты)
    LEGEND_LOC        = "lower center"  # Положение легенды ('lower center', 'upper right'...)
    LEGEND_NCOL       = 4             # Количество колонок в легенде
    LEGEND_FRAMEON    = False         # Показывать ли рамку вокруг легенды
    LEGEND_Y_OFFSET   = -0.20         # Вертикальный отступ легенды от оси X
                                      # (отрицательный — ниже оси, положительный — выше)

    # ==========================================================================
    # РАЗРЫВ СТОЛБЦОВ (BREAK)
    # ==========================================================================
    # Механизм визуального разрыва для высоких столбцов МРРР.
    # 
    # Логика работы:
    # 1. Если максимальное значение МРРР превышает MRRR_BREAK_MULTIPLIER × медиану,
    #    активируется режим разрыва.
    # 2. Столбцы выше y_cap (MRRR_CAP_MULTIPLIER × медиана) обрезаются или сжимаются.
    # 3. На месте разрыва рисуется символ «зигзаг» (две ломаные линии).
    # 4. Подписи над разорванными столбцами смещаются выше символа разрыва.
    
    MRRR_BREAK_MULTIPLIER = 2           # Порог активации разрыва:
                                        # если max(MRRR) > медиана × множитель
    
    MRRR_CAP_MULTIPLIER   = 1.8           # Ограничение высоты столбцов:
                                        # y_cap = медиана × множитель
    
    BREAK_PROPORTIONAL    = True        # Режим сжатия высоких столбцов:
                                        # True  — сохранять пропорции между высокими столбцами
                                        #         (столбец 28 > столбца 20 на графике)
                                        # False — все превышающие столбцы обрезаются по y_cap
    
    BREAK_LINE_HEIGHT     = 0.030       # Высота одного символа разрыва (зигзага)
                                        # в долях от y_cap
    
    BREAK_LINE_WIDTH      = 1.2         # Ширина символа разрыва в долях от BAR_WIDTH
    
    BREAK_COLOR           = "#FFFFFF"   # Цвет линий символа разрыва (белый)
    
    BREAK_LINEWIDTH       = 5.0         # Толщина линий символа разрыва (пункты)
    
    BREAK_LABEL_Y_MULT    = 2.5         # Множитель для отступа подписи над разрывом:
                                        # label_y = break_top × (1 + BREAK_LINE_HEIGHT × множитель)

    BREAK_SYMBOL_Y_MULTS  = (10.5,)       # Позиции зигзагов по Y относительно break_top:
                                        # (7.5,) — один зигзаг; (10.5, 6.0) — два зигзага


# =============================================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# =============================================================================

def format_label(value_grn: float) -> str:
    """Переводит значение в грн. → подпись в тыс. грн. с пробелом-разделителем."""
    if pd.isna(value_grn):
        return ""
    return f"{round(value_grn / 1000):,.0f}".replace(",", " ")


def draw_break_symbol(ax, x_pos: float, bar_top: float) -> None:
    """
    Рисует два зигзага (символ разрыва) поверх обрезанного столбца.

    Параметры
    ---------
    ax      : основная ось matplotlib
    x_pos   : X-координата центра столбца
    bar_top : Y-координата верхней грани обрезанного столбца
    """
    cfg = Config
    h = bar_top * cfg.BREAK_LINE_HEIGHT          # высота одного зигзага
    w = cfg.BAR_WIDTH * cfg.BREAK_LINE_WIDTH / 2  # полуширина

    # Два зигзага: верхний и нижний
    for y_mult in cfg.BREAK_SYMBOL_Y_MULTS:
        y0 = bar_top - h * y_mult
        xs = [x_pos - w, x_pos - w / 3, x_pos + w / 3, x_pos + w]
        ys = [y0 + h,    y0,             y0 + h,          y0       ]
        ax.plot(
            xs, ys,
            color=cfg.BREAK_COLOR,
            linewidth=cfg.BREAK_LINEWIDTH,
            solid_capstyle="round",
            zorder=15,
        )


# =============================================================================
# ПОЛУЧЕНИЕ ДАННЫХ
# =============================================================================

def get_chart_data() -> pd.DataFrame:
    """
    Читает таблицу tDB_History_2 с листа sys через xlwings.
    Возвращает последние N_LAST строк с приведёнными типами.
    """
    cfg = Config
    wb    = xw.Book.caller()
    sheet = wb.sheets[cfg.DATA_SHEET]

    try:
        table = sheet.api.ListObjects(cfg.DATA_TABLE)
    except Exception as e:
        raise ValueError(
            f"Таблиця '{cfg.DATA_TABLE}' не знайдена на аркуші '{cfg.DATA_SHEET}': {e}"
        )

    df = sheet.range(table.Range.Address).options(
        pd.DataFrame, header=1, index=False
    ).value

    # Приведение типов
    df[cfg.COL_DATE] = pd.to_datetime(df[cfg.COL_DATE])
    for col in [cfg.COL_MRRR, cfg.COL_VAL, cfg.COL_PCT, cfg.COL_TOVAR]:
        df[col] = pd.to_numeric(df[col], errors="coerce").astype(float)

    df = df.tail(cfg.N_LAST).reset_index(drop=True)

    # Диагностический вывод диапазонов
    for col in [cfg.COL_MRRR, cfg.COL_VAL, cfg.COL_PCT, cfg.COL_TOVAR]:
        print(f"   {col}: min={df[col].min():.0f}, max={df[col].max():.0f}")

    period = (
        f"{df[cfg.COL_DATE].iloc[0].strftime('%d.%m.%Y')}"
        f" — {df[cfg.COL_DATE].iloc[-1].strftime('%d.%m.%Y')}"
    )
    print(f"   Прочитано {len(df)} рядків: {period}")
    logger.info(f"Прочитано {len(df)} строк: {period}")

    return df


# =============================================================================
# ПОСТРОЕНИЕ ГРАФИКА
# =============================================================================

def _compute_bar_heights(mrrr: np.ndarray, use_break: bool, y_cap: float):
    """
    Вычисляет визуальные высоты столбцов с учётом разрыва.

    Если BREAK_PROPORTIONAL=True — столбцы выше y_cap сжимаются пропорционально,
    сохраняя взаимные соотношения (столбец 28 > столбца 20 на графике).
    Если False — все превышающие столбцы обрезаются точно по y_cap.

    Возвращает (bar_heights, break_tops, k_proportional) — высоты столбцов и
    коэффициент сжатия (или None, если не используется пропорциональный режим).
    """
    if not use_break:
        return mrrr.copy(), mrrr.copy(), None

    if Config.BREAK_PROPORTIONAL:
        y_max_render = y_cap * 1.3                              # предел для самого высокого
        mrrr_max     = np.max(mrrr)
        # линейное сжатие: y_cap + (v - y_cap) * k = y_max_render при v = mrrr_max
        k            = (y_max_render - y_cap) / (mrrr_max - y_cap) if mrrr_max > y_cap else 1.0
        bar_heights  = np.where(mrrr > y_cap, y_cap + (mrrr - y_cap) * k, mrrr)
        break_tops   = bar_heights.copy()
    else:
        bar_heights = np.where(mrrr > y_cap, y_cap, mrrr)
        break_tops  = np.full_like(mrrr, y_cap)
        k = None

    return bar_heights, break_tops, k


def _draw_line_with_labels(ax, x, values, line_kw: dict, label_kw: dict,
                           position: str, offset: float, y_limit: float,
                           values_label=None, y_label_base=None) -> None:
    """
    Рисует линию с маркерами и подписями данных.

    Параметры
    ---------
    ax            : ось matplotlib
    x             : координаты по X
    values        : массив значений для РИСОВАНИЯ (NaN пропускаются автоматически)
    line_kw       : kwargs для ax.plot()
    label_kw      : kwargs для ax.text() (fontsize, fontweight, color …)
    position      : 'top' или 'bottom'
    offset        : отступ подписи от базовой точки
    y_limit       : верхняя граница оси — точки выше не подписываются
    values_label  : (опц.) оригинальные значения для текста подписей
                    Нужно при «обрезке» линии: рисуем по y_cap, подписываем оригиналом.
    y_label_base  : (опц.) Y-координаты базовой точки для размещения подписей.
                    Если None — используются значения из values (draw_val).
                    Нужно в режиме разрыва для PCT: подписи над верхушкой столбца,
                    а не над обрезанной линией внутри бара.
    """
    ax.plot(x, values, **line_kw)

    labels_src = values_label if values_label is not None else values

    for i, draw_val in enumerate(values):
        orig_val = labels_src[i]
        # Пропускаем NaN и точки вне видимой области
        if pd.isna(orig_val) or draw_val > y_limit:
            continue

        base_y  = y_label_base[i] if y_label_base is not None else draw_val
        label_y = base_y - offset if position == "bottom" else base_y + offset
        va      = "top"            if position == "bottom" else "bottom"
        ax.text(x[i], label_y, format_label(orig_val),
                ha="center", va=va, zorder=20, **label_kw)


def build_chart(df: pd.DataFrame):
    """
    Строит полную matplotlib-фигуру по данным df.

    Последовательность отрисовки:
      1. Столбцы МРРР (+ символы разрыва)
      2. Линии на основной оси (Валютний, Процентний)
      3. Линия на вторичной оси (Товарний)
      4. Подписи данных
      5. Настройка осей, заголовка, легенды
    """
    cfg = Config

    dates    = df[cfg.COL_DATE]
    mrrr     = df[cfg.COL_MRRR].values
    val_risk = df[cfg.COL_VAL].values
    pct_risk = df[cfg.COL_PCT].values
    tovar    = df[cfg.COL_TOVAR].values
    x        = np.arange(len(dates))

    # --- Логика разрыва ---
    median_mrrr = np.median(mrrr)
    use_break   = bool(np.any(mrrr > cfg.MRRR_BREAK_MULTIPLIER * median_mrrr))
    y_cap       = median_mrrr * cfg.MRRR_CAP_MULTIPLIER if use_break else None

    bar_heights, break_tops, k_proportional = _compute_bar_heights(mrrr, use_break, y_cap)

    # --- Инициализация фигуры ---
    fig, ax = plt.subplots(figsize=cfg.FIGURE_SIZE)
    fig.patch.set_facecolor(cfg.BG_COLOR)
    ax.set_facecolor(cfg.BG_COLOR)
    ax2 = ax.twinx()                    # вторичная ось для Товарний ризик

    # --- Масштаб основной оси Y ---
    # В режиме разрыва линии тоже ограничиваем y_cap, чтобы выброс не сжимал график
    def _safe_max(arr):
        """Максимум массива без NaN; при разрыве — ограничен сжатыми значениями."""
        clean = np.where(np.isnan(arr), 0, arr)
        if use_break and k_proportional is not None:
            # Пропорциональное сжатие
            compressed = np.where(clean > y_cap, y_cap + (clean - y_cap) * k_proportional, clean)
            return float(np.max(compressed))
        elif use_break:
            # Простая обрезка
            return float(np.max(np.minimum(clean, y_cap)))
        else:
            return float(np.max(clean))

    y_main_max = max(
        float(np.max(bar_heights)),
        _safe_max(val_risk),
        _safe_max(pct_risk),
    )
    y_main_top    = y_main_max * 1.25   # запас для подписей сверху
    label_offset  = y_main_top * cfg.LABEL_OFFSET_Y

    # --- 1. Столбцы МРРР ---
    ax.bar(
        x, bar_heights,
        width=cfg.BAR_WIDTH, color=cfg.BAR_COLOR,
        alpha=cfg.BAR_ALPHA, zorder=2,
        label=cfg.COL_MRRR,
    )

    # --- 2. Символы разрыва и подписи столбцов ---
    for i, val in enumerate(mrrr):
        is_broken = use_break and val > y_cap

        if is_broken:
            draw_break_symbol(ax, x[i], break_tops[i])
            label_y = break_tops[i] * (1 + cfg.BREAK_LINE_HEIGHT * cfg.BREAK_LABEL_Y_MULT)
        else:
            label_y = bar_heights[i] + label_offset

        ax.text(
            x[i], label_y, format_label(val),
            ha="center", va="bottom", zorder=20,
            fontsize=cfg.LABEL_FONTSIZE_BAR,
            fontweight=cfg.LABEL_FONTWEIGHT_MRRR,
            fontstyle=cfg.LABEL_FONTSTYLE,
            color=cfg.LABEL_COLOR_LINES,
        )

    # --- 3. Линии на основной оси ---
    label_common = dict(
        fontweight=cfg.LABEL_FONTWEIGHT,
        fontstyle=cfg.LABEL_FONTSTYLE,
    )

    # Валютний ризик: в режиме разрыва тоже применяем пропорциональное сжатие
    if use_break and k_proportional is not None:
        val_draw = np.where(
            np.isnan(val_risk),
            np.nan,
            np.where(val_risk > y_cap, y_cap + (val_risk - y_cap) * k_proportional, val_risk)
        )
    elif use_break:
        val_draw = np.where(np.isnan(val_risk), np.nan, np.minimum(val_risk, y_cap))
    else:
        val_draw = val_risk
    _draw_line_with_labels(
        ax, x, val_draw,
        line_kw=dict(
            color=cfg.LINE_VAL_COLOR, linewidth=cfg.LINE_WIDTH,
            marker=cfg.MARKER_STYLE_VAL, markersize=cfg.MARKER_SIZE,
            label=cfg.COL_VAL, zorder=10,
        ),
        label_kw=dict(color=cfg.LABEL_COLOR_VAL, fontsize=cfg.LABEL_FONTSIZE_VAL, **label_common),
        position=cfg.LABEL_POSITION_VAL,
        offset=y_main_top * cfg.LABEL_OFFSET_VAL,
        y_limit=y_main_top,
        values_label=val_risk,          # подписи — оригинальные значения
    )

    # Процентний ризик: значения 01.02.26 и 01.03.26 (~19-21M) превышают y_cap (~6.7M).
    # Для РИСОВАНИЯ обрезаем до y_cap (линия остаётся в видимой области).
    # Для ПОДПИСЕЙ используем оригинальные значения.
    # В режиме разрыва линия оказывается ВНУТРИ обрезанного столбца (bar до y_cap*1.3),
    # поэтому подписи выводим ВЫШЕ верхушки столбца (y_label_base=break_heights),
    # а не под обрезанной линией, где тёмный текст сливался бы с тёмным баром.
    if use_break and k_proportional is not None:
        # Пропорциональное сжатие для сохранения соотношений между значениями
        pct_draw = np.where(
            np.isnan(pct_risk), 
            np.nan, 
            np.where(pct_risk > y_cap, y_cap + (pct_risk - y_cap) * k_proportional, pct_risk)
        )
    elif use_break:
        # Простая обрезка по y_cap
        pct_draw = np.where(np.isnan(pct_risk), np.nan, np.minimum(pct_risk, y_cap))
    else:
        pct_draw = pct_risk
    
    _draw_line_with_labels(
        ax, x, pct_draw,
        line_kw=dict(
            color=cfg.LINE_PCT_COLOR, linewidth=cfg.LINE_WIDTH,
            marker=cfg.MARKER_STYLE_PCT, markersize=cfg.MARKER_SIZE,
            label=cfg.COL_PCT, zorder=10,
        ),
        label_kw=dict(color=cfg.LABEL_COLOR_PCT, fontsize=cfg.LABEL_FONTSIZE_PCT, **label_common),
        position="top" if use_break else cfg.LABEL_POSITION_PCT,
        offset=y_main_top * cfg.LABEL_OFFSET_PCT,
        y_limit=y_main_top,
        values_label=pct_risk,          # подписи — оригинальные значения
        y_label_base=break_tops if use_break else None,    # база — верх столбца
    )

    # --- 4. Линия Товарний ризик на вторичной оси ---
    tovar_max = float(np.nanmax(tovar))
    tovar_top = tovar_max * cfg.YAXIS2_TOP_MULTIPLIER

    _draw_line_with_labels(
        ax2, x, tovar,
        line_kw=dict(
            color=cfg.LINE_TOVAR_COLOR, linewidth=cfg.LINE_WIDTH,
            marker=cfg.MARKER_STYLE_TOVAR, markersize=cfg.MARKER_SIZE,
            label=cfg.COL_TOVAR, zorder=10,
        ),
        label_kw=dict(color=cfg.LABEL_COLOR_TOVAR, fontsize=cfg.LABEL_FONTSIZE_TOVAR, **label_common),
        position=cfg.LABEL_POSITION_TOVAR,
        offset=tovar_top * cfg.LABEL_OFFSET_TOVAR,
        y_limit=tovar_top,
    )

    # --- 5. Настройка осей ---
    ax.set_ylim(0, y_main_top)
    ax.yaxis.set_visible(cfg.YAXIS_VISIBLE)
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.spines["bottom"].set_visible(True)
    ax.spines["bottom"].set_color(cfg.XAXIS_COLOR)

    ax2.set_ylim(0, tovar_top)
    ax2.yaxis.set_visible(cfg.YAXIS2_VISIBLE)
    for spine in ax2.spines.values():
        spine.set_visible(False)

    ax.set_xticks(x)
    ax.set_xticklabels(
        [d.strftime(cfg.XAXIS_DATE_FORMAT) for d in dates],
        fontsize=cfg.XAXIS_FONTSIZE,
        fontweight=cfg.XAXIS_FONTWEIGHT,
        fontstyle=cfg.XAXIS_FONTSTYLE,
        color=cfg.XAXIS_COLOR,
    )
    ax.tick_params(axis="x", which="both", length=0, pad=cfg.XAXIS_LABEL_PAD)

    # --- 6. Заголовок ---
    ax.set_title(
        cfg.TITLE_TEXT,
        fontsize=cfg.TITLE_FONTSIZE,
        fontweight=cfg.TITLE_FONTWEIGHT,
        color=cfg.TITLE_COLOR,
        pad=cfg.TITLE_PAD,
        loc="left",
    )

    # --- 7. Легенда (объединяем обе оси) ---
    handles1, labels1 = ax.get_legend_handles_labels()
    handles2, labels2 = ax2.get_legend_handles_labels()
    ax.legend(
        handles1 + handles2,
        labels1  + labels2,
        fontsize=cfg.LEGEND_FONTSIZE,
        loc=cfg.LEGEND_LOC,
        ncol=cfg.LEGEND_NCOL,
        frameon=cfg.LEGEND_FRAMEON,
        bbox_to_anchor=(0.5, cfg.LEGEND_Y_OFFSET),
    )

    plt.tight_layout()
    return fig


# =============================================================================
# СОХРАНЕНИЕ И ВСТАВКА В EXCEL
# =============================================================================

def save_chart(fig) -> str:
    """Сохраняет фигуру в домашнюю директорию, возвращает путь к файлу."""
    path = os.path.join(os.path.expanduser("~"), Config.TEMP_IMAGE)
    fig.savefig(
        path,
        dpi=Config.IMAGE_DPI,
        bbox_inches="tight",
        facecolor=Config.BG_COLOR,
    )
    plt.close(fig)
    return path


def insert_chart_to_excel(img_path: str) -> None:
    """Удаляет старый объект IMAGE_NAME и вставляет новое изображение в IMAGE_CELL."""
    cfg = Config
    wb  = xw.Book.caller()
    ws  = wb.sheets[cfg.IMAGE_SHEET]

    # Удаляем предыдущее изображение (если есть)
    for pic in ws.pictures:
        if pic.name == cfg.IMAGE_NAME:
            pic.delete()

    pic = ws.pictures.add(
        img_path,
        name=cfg.IMAGE_NAME,
        update=True,
        left=ws.range(cfg.IMAGE_CELL).left,
        top=ws.range(cfg.IMAGE_CELL).top,
    )
    pic.width  = cfg.IMAGE_WIDTH
    pic.height = cfg.IMAGE_HEIGHT


# =============================================================================
# ПУБЛИЧНАЯ ТОЧКА ВХОДА
# =============================================================================

def create_market_risk_chart() -> None:
    """
    Основная публичная функция.
    Вызывается из Excel через xlwings / main.py.
    """
    try:
        print("=== ГРАФІК: ДИНАМІКА МІНІМАЛЬНОГО РОЗМІРУ РИНКОВОГО РИЗИКУ ===")

        print("1. Читання даних...")
        df = get_chart_data()

        print("2. Побудова графіка...")
        fig = build_chart(df)

        print("3. Збереження у тимчасовий файл...")
        img_path = save_chart(fig)
        print(f"   Файл: {img_path}")

        print(
            f"4. Вставка в Excel "
            f"(аркуш '{Config.IMAGE_SHEET}', комірка '{Config.IMAGE_CELL}')..."
        )
        insert_chart_to_excel(img_path)

        print("Готово! Графік успішно створено та вставлено.")
        logger.info("График рыночного риска успешно создан и вставлен.")

    except Exception as e:
        print(f"ПОМИЛКА: {e}")
        logger.error(f"Ошибка при создании графика: {e}", exc_info=True)
        raise
