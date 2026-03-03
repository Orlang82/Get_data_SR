import xlwings as xw
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
import sys
import logging

# Для корректного вывода в консоли Windows
if sys.platform == "win32":
    os.system("chcp 65001 > NUL")
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# =============================================================================
# ЛОГИРОВАНИЕ
# =============================================================================
ENABLE_LOGGING = True


def _setup_logger():
    if not ENABLE_LOGGING:
        return logging.getLogger("chart_7s_disabled")
    logger = logging.getLogger("chart_7s")
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    log_dir = os.path.abspath(os.path.join(script_dir, '..', 'logs'))
    os.makedirs(log_dir, exist_ok=True)
    handler = logging.FileHandler(os.path.join(log_dir, 'chart_7s.log'), encoding='utf-8')
    handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logger.addHandler(handler)
    return logger


logger = _setup_logger()


# =============================================================================
# КОНФИГУРАЦИЯ - ВСЕ ПАРАМЕТРЫ ЗАДАЮТСЯ ЗДЕСЬ
# =============================================================================
class Config:
    # --- Источник данных ---
    DATA_SHEET       = 'sys'                # Лист с данными
    DATA_TABLE       = 'tDB_History_2'      # Имя Excel-таблицы
    COL_DATE         = 'Дата'              # Колонка с датой
    COL_MRRR         = 'МРРР'             # Колонка МРРР (столбчатая гистограмма)
    COL_VAL          = 'Валютний ризик'    # Колонка валютного риска
    COL_PCT          = 'Процентний ризик'  # Колонка процентного риска
    COL_TOVAR        = 'Товарний ризик'    # Колонка товарного риска (вторичная ось)
    N_LAST           = 12                  # Количество последних наблюдений

    # --- Вставка изображения в Excel ---
    IMAGE_SHEET      = 'To_Report'         # Лист для вставки
    IMAGE_CELL       = 'C17'              # Ячейка привязки
    IMAGE_NAME       = 'Chart_MarketRisk_7S'  # Имя объекта в Excel
    TEMP_IMAGE_NAME  = 'market_risk_7s.png'   # Временный файл (в home dir)
    IMAGE_WIDTH      = 720               # Ширина изображения в Excel (пиксели)
    IMAGE_HEIGHT     = 200                # Высота изображения в Excel (пиксели)
    IMAGE_DPI        = 150                # DPI сохранения

    # --- Размер и стиль фигуры ---
    FIGURE_SIZE      = (8.0, 4)         # Размер matplotlib-фигуры (дюймы)
    BG_COLOR         = '#FFFFFF'          # Цвет фона фигуры и осей

    # --- Параметры столбчатой гистограммы (МРРР) ---
    BAR_COLOR        = '#1F3864'          # Цвет столбцов (тёмно-синий)
    BAR_WIDTH        = 0.7              # Ширина столбца (0..1)
    BAR_ALPHA        = 1.0               # Прозрачность столбцов

    # --- Линии ---
    LINE_VAL_COLOR   = '#00B050'          # Цвет линии Валютний ризик (зелёный)
    LINE_PCT_COLOR   = '#FFD700'          # Цвет линии Процентний ризик (жёлтый/золотой)
    LINE_TOVAR_COLOR = '#ED7D31'          # Цвет линии Товарний ризик (оранжевый)
    LINE_WIDTH       = 2.0               # Толщина линий
    MARKER_SIZE      = 4                 # Размер маркеров на линиях
    MARKER_STYLE_VAL   = 'o'            # Маркер Валютний ризик
    MARKER_STYLE_PCT   = 'o'            # Маркер Процентний ризик
    MARKER_STYLE_TOVAR = 'o'            # Маркер Товарний ризик

    # --- Подписи данных ---
    LABEL_FONTSIZE   = 6                 # Размер шрифта подписей данных
    LABEL_FONTWEIGHT = 'normal'          # Жирность: 'normal', 'bold'
    LABEL_FONTSTYLE  = 'normal'          # Стиль: 'normal', 'italic', 'oblique'
    LABEL_COLOR_BAR  = '#FFFFFF'         # Цвет подписей на столбцах (белый, внутри бара)
    LABEL_COLOR_LINES = '#1F3864'        # Цвет подписей на линиях
    LABEL_FONTWEIGHT_MRRR = 'bold'       # Жирность подписей для МРРР
    LABEL_OFFSET_Y   = 0.1            # Вертикальный отступ подписи от точки (доля от y_max)

    # --- Заголовок графика ---
    TITLE_TEXT       = 'Динаміка мінімального розміру ринкового ризику'
    TITLE_FONTSIZE   = 10
    TITLE_FONTWEIGHT = 'bold'
    TITLE_COLOR      = '#1F3864'

    # --- Ось X ---
    XAXIS_DATE_FORMAT = '%d.%m.%y'       # Формат дат на оси X (01.03.25)
    XAXIS_FONTSIZE    = 6                # Размер шрифта меток оси X
    XAXIS_FONTWEIGHT  = 'bold'           # Жирность: 'normal', 'bold'
    XAXIS_FONTSTYLE   = 'normal'         # Стиль: 'normal', 'italic', 'oblique'
    XAXIS_COLOR       = '#1F3864'

    # --- Ось Y (основная и вторичная) ---
    YAXIS_VISIBLE         = False        # Основная ось Y — не отображать
    YAXIS2_VISIBLE        = False        # Вторичная ось Y — не отображать
    YAXIS2_TOP_MULTIPLIER = 4.5          # Верхняя граница вторичной оси = max(Товарний) × множитель;
                                         # увеличьте, чтобы опустить линию вниз (напр. 2.0, 3.0)

    # --- Легенда ---
    LEGEND_FONTSIZE   = 6                # Размер шрифта легенды
    LEGEND_LOC        = 'lower center'   # Расположение легенды
    LEGEND_NCOL       = 4                # Количество колонок легенды
    LEGEND_FRAMEON    = False            # Рамка легенды

    # --- Разрыв столбцов (break) ---
    MRRR_BREAK_MULTIPLIER = 3          # Порог для активации разрыва (× медиана) 2,5
    MRRR_CAP_MULTIPLIER   = 2.5          # Ограничение высоты столбца (× медиана) 1,8
    BREAK_LINE_HEIGHT     = 0.030        # Высота символа разрыва (доля от y_cap) 0,018
    BREAK_LINE_WIDTH      = 0.75         # Ширина символа разрыва (доля от bar_width)
    BREAK_COLOR           = '#FFFFFF'    # Цвет символа разрыва
    BREAK_LINEWIDTH       = 4.0          # Толщина линий символа разрыва
    BREAK_LABEL_Y_MULT    = 3.5          # Множитель отступа подписи над разорванным столбцом
    BREAK_SYMBOL_Y_MULTS  = (7.5, 6.0)   # Множители положения символа разрыва по оси Y



# =============================================================================
# ФУНКЦИИ
# =============================================================================

def get_chart_data() -> pd.DataFrame:
    """Читает таблицу tDB_History_2 с листа sys, возвращает последние N_LAST строк."""
    wb = xw.Book.caller()
    sheet = wb.sheets[Config.DATA_SHEET]

    try:
        table = sheet.api.ListObjects(Config.DATA_TABLE)
    except Exception as e:
        raise ValueError(
            f"Таблица '{Config.DATA_TABLE}' не найдена на листе '{Config.DATA_SHEET}': {e}"
        )

    rng_addr = table.Range.Address
    df = sheet.range(rng_addr).options(pd.DataFrame, header=1, index=False).value

    # Приведение типов
    df[Config.COL_DATE] = pd.to_datetime(df[Config.COL_DATE])
    for col in [Config.COL_MRRR, Config.COL_VAL, Config.COL_PCT, Config.COL_TOVAR]:
        df[col] = pd.to_numeric(df[col], errors='coerce').astype(float)

    # Последние N_LAST строк
    df = df.tail(Config.N_LAST).reset_index(drop=True)

    logger.info(
        f"Прочитано {len(df)} строк: "
        f"{df[Config.COL_DATE].iloc[0].date()} — {df[Config.COL_DATE].iloc[-1].date()}"
    )
    print(
        f"   Прочитано {len(df)} строк: "
        f"{df[Config.COL_DATE].iloc[0].strftime('%d.%m.%Y')} — "
        f"{df[Config.COL_DATE].iloc[-1].strftime('%d.%m.%Y')}"
    )
    return df


def format_label(value_grn: float) -> str:
    """Форматирует значение в грн. → подпись в тыс. грн. с пробелом-разделителем тысяч."""
    if pd.isna(value_grn):
        return ""
    value_k = round(value_grn / 1000)
    return f"{value_k:,.0f}".replace(",", " ")


def draw_break_symbol(ax, x: float, y_cap: float):
    """Рисует символ разрыва (два зигзага) поверх обрезанного столбца."""
    h = y_cap * Config.BREAK_LINE_HEIGHT
    w = Config.BAR_WIDTH * Config.BREAK_LINE_WIDTH

    # Два зигзага: верхний и нижний
    for offset_y in [y_cap - h * Config.BREAK_SYMBOL_Y_MULTS[0], y_cap - h * Config.BREAK_SYMBOL_Y_MULTS[1]]:
        xs = [x - w, x - w / 3, x + w / 3, x + w]
        ys = [offset_y + h, offset_y, offset_y + h, offset_y]
        ax.plot(
            xs, ys,
            color=Config.BREAK_COLOR,
            linewidth=Config.BREAK_LINEWIDTH,
            solid_capstyle='round',
            zorder=5,
        )


def build_chart(df: pd.DataFrame):
    """Строит полную фигуру matplotlib по данным DataFrame."""
    cfg = Config

    dates    = df[cfg.COL_DATE]
    mrrr     = df[cfg.COL_MRRR].values
    val_risk = df[cfg.COL_VAL].values
    pct_risk = df[cfg.COL_PCT].values
    tovar    = df[cfg.COL_TOVAR].values
    x        = np.arange(len(dates))

    # --- Логика разрыва (break) ---
    median_mrrr = np.median(mrrr)
    use_break   = bool(np.any(mrrr > cfg.MRRR_BREAK_MULTIPLIER * median_mrrr))
    y_cap       = median_mrrr * cfg.MRRR_CAP_MULTIPLIER if use_break else None

    # --- Фигура ---
    fig, ax = plt.subplots(figsize=cfg.FIGURE_SIZE)
    fig.patch.set_facecolor(cfg.BG_COLOR)
    ax.set_facecolor(cfg.BG_COLOR)
    ax2 = ax.twinx()  # Вторичная ось для Товарний ризик

    # --- Столбцы МРРР ---
    bar_heights = np.where(use_break & (mrrr > y_cap), y_cap, mrrr) if use_break else mrrr.copy()
    ax.bar(
        x, bar_heights,
        width=cfg.BAR_WIDTH, color=cfg.BAR_COLOR,
        alpha=cfg.BAR_ALPHA, zorder=2,
        label=cfg.COL_MRRR,
    )

    # Символы разрыва над обрезанными столбцами
    if use_break:
        for i, val in enumerate(mrrr):
            if val > y_cap:
                draw_break_symbol(ax, x[i], y_cap)

    # --- Пределы основной оси Y ---
    # В режиме разрыва линии тоже ограничиваем порогом y_cap,
    # чтобы один выброс не раздувал шкалу оси
    if use_break:
        val_for_scale = np.where(np.isnan(val_risk), 0, np.minimum(val_risk, y_cap))
        pct_for_scale = np.where(np.isnan(pct_risk), 0, np.minimum(pct_risk, y_cap))
    else:
        val_for_scale = np.where(np.isnan(val_risk), 0, val_risk)
        pct_for_scale = np.where(np.isnan(pct_risk), 0, pct_risk)
    y_main_max = max(
        float(np.max(bar_heights)),
        float(np.max(val_for_scale)),
        float(np.max(pct_for_scale)),
    )
    y_main_top = y_main_max * 1.25          # Запас для подписей сверху
    label_offset = y_main_top * cfg.LABEL_OFFSET_Y

    # --- Подписи данных на столбцах МРРР ---
    for i, val in enumerate(mrrr):
        lbl = format_label(val)
        if use_break and val > y_cap:
            # Над символом разрыва (выше y_cap)
            label_y = y_cap * (1 + cfg.BREAK_LINE_HEIGHT * cfg.BREAK_LABEL_Y_MULT)
            ax.text(
                x[i], label_y, lbl,
                ha='center', va='bottom',
                fontsize=cfg.LABEL_FONTSIZE, fontweight=cfg.LABEL_FONTWEIGHT_MRRR, fontstyle=cfg.LABEL_FONTSTYLE,
                color=cfg.LABEL_COLOR_LINES, zorder=7,
            )
        else:
            # Над столбцом
            ax.text(
                x[i], bar_heights[i] + label_offset, lbl,
                ha='center', va='bottom',
                fontsize=cfg.LABEL_FONTSIZE, fontweight=cfg.LABEL_FONTWEIGHT_MRRR, fontstyle=cfg.LABEL_FONTSTYLE,
                color=cfg.LABEL_COLOR_LINES, zorder=3,
            )

    # --- Линии на основной оси: Валютний та Процентний ризик ---
    ax.plot(
        x, val_risk,
        color=cfg.LINE_VAL_COLOR, linewidth=cfg.LINE_WIDTH,
        marker=cfg.MARKER_STYLE_VAL, markersize=cfg.MARKER_SIZE,
        label=cfg.COL_VAL, zorder=3,
    )
    ax.plot(
        x, pct_risk,
        color=cfg.LINE_PCT_COLOR, linewidth=cfg.LINE_WIDTH,
        marker=cfg.MARKER_STYLE_PCT, markersize=cfg.MARKER_SIZE,
        label=cfg.COL_PCT, zorder=3,
    )

    # Подписи для линий основной оси (пропускаем точки выше видимой области)
    for i in range(len(x)):
        if not pd.isna(val_risk[i]) and val_risk[i] <= y_main_top:
            ax.text(
                x[i], val_risk[i] - label_offset, format_label(val_risk[i]),
                ha='center', va='top',
                fontsize=cfg.LABEL_FONTSIZE, fontweight=cfg.LABEL_FONTWEIGHT, fontstyle=cfg.LABEL_FONTSTYLE,
                color=cfg.LABEL_COLOR_LINES, zorder=4,
            )
        if not pd.isna(pct_risk[i]) and pct_risk[i] <= y_main_top:
            ax.text(
                x[i], pct_risk[i] + label_offset, format_label(pct_risk[i]),
                ha='center', va='bottom',
                fontsize=cfg.LABEL_FONTSIZE, fontweight=cfg.LABEL_FONTWEIGHT, fontstyle=cfg.LABEL_FONTSTYLE,
                color=cfg.LABEL_COLOR_LINES, zorder=4,
            )

    # --- Линия на вторичной оси: Товарний ризик ---
    ax2.plot(
        x, tovar,
        color=cfg.LINE_TOVAR_COLOR, linewidth=cfg.LINE_WIDTH,
        marker=cfg.MARKER_STYLE_TOVAR, markersize=cfg.MARKER_SIZE,
        label=cfg.COL_TOVAR, zorder=3,
    )

    # Подписи для вторичной оси (координаты в системе ax2)
    tovar_max   = float(np.nanmax(tovar))
    tovar_top   = tovar_max * cfg.YAXIS2_TOP_MULTIPLIER
    label_off2  = tovar_top * cfg.LABEL_OFFSET_Y
    for i in range(len(x)):
        if not pd.isna(tovar[i]):
            ax2.text(
                x[i], tovar[i] - label_off2, format_label(tovar[i]),
                ha='center', va='top',
                fontsize=cfg.LABEL_FONTSIZE, fontweight=cfg.LABEL_FONTWEIGHT,
                color=cfg.LABEL_COLOR_LINES, zorder=4,
            )

    # --- Настройка осей ---
    ax.set_ylim(0, y_main_top)
    ax.yaxis.set_visible(cfg.YAXIS_VISIBLE)
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.spines['bottom'].set_visible(True)
    ax.spines['bottom'].set_color(cfg.XAXIS_COLOR)

    ax2.set_ylim(0, tovar_top)
    ax2.yaxis.set_visible(cfg.YAXIS2_VISIBLE)
    for spine in ax2.spines.values():
        spine.set_visible(False)

    # --- Ось X ---
    ax.set_xticks(x)
    ax.set_xticklabels(
        [d.strftime(cfg.XAXIS_DATE_FORMAT) for d in dates],
        fontsize=cfg.XAXIS_FONTSIZE, fontweight=cfg.XAXIS_FONTWEIGHT,
        fontstyle=cfg.XAXIS_FONTSTYLE, color=cfg.XAXIS_COLOR,
    )
    ax.tick_params(axis='x', which='both', length=0)

    # --- Заголовок ---
    ax.set_title(
        cfg.TITLE_TEXT,
        fontsize=cfg.TITLE_FONTSIZE,
        fontweight=cfg.TITLE_FONTWEIGHT,
        color=cfg.TITLE_COLOR,
        pad=12,
    )

    # --- Легенда (объединяем обе оси) ---
    handles1, labels1 = ax.get_legend_handles_labels()
    handles2, labels2 = ax2.get_legend_handles_labels()
    ax.legend(
        handles1 + handles2, labels1 + labels2,
        fontsize=cfg.LEGEND_FONTSIZE,
        loc=cfg.LEGEND_LOC,
        ncol=cfg.LEGEND_NCOL,
        frameon=cfg.LEGEND_FRAMEON,
        bbox_to_anchor=(0.5, -0.30), # отступ легенды от оси х
    )

    plt.tight_layout()
    return fig


def save_chart(fig) -> str:
    """Сохраняет фигуру во временный файл, возвращает путь к файлу."""
    temp_path = os.path.join(os.path.expanduser("~"), Config.TEMP_IMAGE_NAME)
    fig.savefig(
        temp_path,
        dpi=Config.IMAGE_DPI,
        bbox_inches='tight',
        facecolor=Config.BG_COLOR,
    )
    plt.close(fig)
    return temp_path


def insert_chart_to_excel(img_path: str):
    """Удаляет старое изображение IMAGE_NAME и вставляет новое в ячейку IMAGE_CELL."""
    wb = xw.Book.caller()
    ws = wb.sheets[Config.IMAGE_SHEET]

    # Удаляем предыдущее изображение (если есть)
    for pic in ws.pictures:
        if pic.name == Config.IMAGE_NAME:
            pic.delete()

    # Вставляем новое изображение
    pic = ws.pictures.add(
        img_path,
        name=Config.IMAGE_NAME,
        update=True,
        left=ws.range(Config.IMAGE_CELL).left,
        top=ws.range(Config.IMAGE_CELL).top,
    )
    pic.width  = Config.IMAGE_WIDTH
    pic.height = Config.IMAGE_HEIGHT


def create_market_risk_chart():
    """
    Основная публичная функция: создаёт и вставляет в Excel
    график «Динаміка мінімального розміру ринкового ризику».
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
            f"(лист '{Config.IMAGE_SHEET}', комірка '{Config.IMAGE_CELL}')..."
        )
        insert_chart_to_excel(img_path)

        print("Готово! Графік успішно створено та вставлено.")
        logger.info("График рыночного риска успешно создан и вставлен.")

    except Exception as e:
        print(f"ПОМИЛКА: {e}")
        logger.error(f"Ошибка при создании графика: {e}", exc_info=True)
        raise
