# QWEN.md — Контекст проекта Get_data_SR

## Обзор проекта

**Get_data_SR** — система управления ликвидностью на Python для интеграции с Excel через xlwings. Проект получает данные из базы данных Oracle (схема SR_BANK), обрабатывает их и вставляет в Excel-таблицы для регуляторной отчётности. Также включает модули построения графиков для визуализации данных (VAR/ES, структура активов/капитала, рыночные риски).

### Ключевые возможности

- **20+ fetchers-модулей** для получения данных по различным статьям отчётности (баланс, резервы, счета, документы, forex-сделки)
- **6 chart-модулей** для построения графиков распределения потерь, динамики рисков, структуры активов
- **Интеграция с Excel** через xlwings UDF и макросы
- **Работа с Oracle DB** через oracledb в Thin Mode
- **SQLite-хранилище** для обработки файлов 6KX

---

## Структура проекта

```
Get_data_SR/
├── main.py                 # Точки входа для вызова из Excel (run_* функции)
├── udf_modules.py          # UDF-функции для Excel формул
├── fetchers/               # Модули получения данных
│   ├── balance_nrk.py      # Баланс НРК
│   ├── detail_6sx.py       # Перечень счетов 6S (с фильтрацией)
│   ├── pay_6sx.py          # Документы для 6S (зависит от detail_6sx)
│   ├── forex_6sx.py        # Forex-сделки 6S (зависит от pay_6sx)
│   ├── interest_7sx.py     # Процентный риск 7S
│   ├── dz_spot.py          # DZ SPOT данные
│   ├── grp_9000.py         # Группа 9000
│   └── ... (20+ модулей)
├── charts/                 # Модули построения графиков
│   ├── chart_es.py         # График распределения VaR/ES
│   ├── chart_as_v2.py      # Структура активов
│   ├── chart_7s_mrrr.py    # Динамика минимального размера рыночного риска
│   └── ...
├── db/                     # Работа с базами данных
│   ├── oracle.py           # Выполнение SQL-запросов к Oracle
│   ├── connect_db_oracle.py # Подключение к Oracle (credentials из ~/.conda/db_ac.json)
│   ├── entry_db_6kx.py     # Обработка单个 файла 6KX в SQLite
│   └── batch_entry_db_6kx.py # Пакетная обработка 6KX
├── sql/                    # SQL-шаблоны с параметрами
│   ├── SR_6SX_ACCOUNT_template.sql
│   ├── SR_BALANCE_NRK_template.sql
│   └── ... (20+ шаблонов)
├── utils/                  # Утилиты
│   ├── excel_writer.py     # Вставка DataFrame в Excel-таблицы
│   ├── date_utils.py       # Расчёт рабочих дней, forecast-режим
│   ├── path_utils.py       # Резолвинг путей к SQL-шаблонам
│   └── parser_forex.py     # Парсинг номеров forex-сделок
├── request/                # Standalone-скрипты (запуск без Excel)
├── logs/                   # Лог-файлы
└── CLAUDE.md               # Подробная документация по паттернам
```

---

## Запуск и тестирование

### Предварительные требования

1. **Credentials для Oracle**: создать файл `~\.conda\db_ac.json`:
```json
{
  "user": "your_username",
  "password": "your_password",
  "dsn": "host:port/service_name"
}
```

2. **Зависимости**: pandas, xlwings, oracledb, matplotlib, seaborn, scipy

### Запуск скриптов

**Fetchers и charts** требуют активную Excel-книгу и вызываются через макрос:
```vba
RunPython "import main; main.run_secur_doc()"
RunPython "import main; main.run_chart_7s()"
```

**Standalone-скрипты** в папке `request/` запускаются напрямую:
```bash
python request/script_name.py
```

---

## Архитектурные паттерны

### 1. Стандартный flow fetcher-модуля

```python
from db.oracle import query
from utils.date_utils import get_previous_working_day, forecast_date
from utils.excel_writer import paste_to_excel  # или paste_to_excel_smart
from utils.path_utils import get_sql_path

def fetch_to_<name>():
    sql_path = get_sql_path("SR_<NAME>_template.sql")
    sql = open(sql_path, encoding="utf-8").read().strip().rstrip(";")
    date_param = forecast_date() or get_previous_working_day()
    return query(sql, {"date_param": date_param})

def paste_to_excel_<name>():
    df = fetch_to_<name>()
    paste_to_excel("<SheetName>", "<TableName>", df)
```

### 2. Добавление нового fetcher-модуля

1. Создать SQL-шаблон в `sql/SR_<NAME>_template.sql` с параметрами `:param_name`
2. Создать модуль в `fetchers/<name>.py` по стандартному паттерну
3. Добавить импорт и функцию в `main.py`:
```python
from fetchers.<name> import paste_to_excel_<name>

def run_<name>():
    """Описание."""
    paste_to_excel_<name>()
```

### 3. Цепочки зависимостей между fetchers

Некоторые fetchers переиспользуют данные других:
```
detail_6sx → pay_6sx → forex_6sx
```
Пример:
```python
from fetchers.detail_6sx import fetch_6sx_data

def fetch_pay_6sx_data():
    acc_calc, _ = fetch_6sx_data()  # переиспользуем результат
    for _, row in acc_calc.iterrows():
        df = query(sql, {"data_acc": row['ACCOUNT_NUMBER'], ...})
        results.append(df)
    return pd.concat(results, ignore_index=True)
```

### 4. Условная логика с приоритетами

При маркировке записей условия применяются в порядке приоритета:
```python
df['mark'] = None
# 1. Высший приоритет
df.loc[condition1, 'mark'] = 'pre_excluded'
# 2. Средний приоритет — только если ещё не помечено
df.loc[condition2 & df['mark'].isna(), 'mark'] = 'exclude'
# 3. Низкий приоритет — только если всё ещё не помечено
df.loc[condition3 & df['mark'].isna(), 'mark'] = 'exclude'
```

### 5. Структура chart-модуля

Все chart-модули следуют единой структуре:

1. **Config class** — все параметры в одном месте (цвета, размеры, имена)
2. **get_chart_data()** — чтение данных из Excel
3. **build_chart(df)** — построение matplotlib-фигуры
4. **save_chart(fig)** — сохранение во временный PNG в `~`
5. **insert_chart_to_excel(img_path)** — вставка в Excel с заменой старого изображения
6. **create_*_chart()** — публичная функция, вызывающая 2→3→4→5

Пример Config:
```python
class Config:
    DATA_SHEET = 'sys'
    DATA_TABLE = 'tDB_History_2'
    IMAGE_SHEET = 'To_Report'
    IMAGE_CELL = 'C17'
    IMAGE_NAME = 'Chart_MarketRisk_7S'
    TEMP_IMAGE_NAME = 'market_risk_7s.png'
    IMAGE_WIDTH = 720
    IMAGE_HEIGHT = 200
    # ... остальные параметры
```

---

## Работа с данными

### Oracle DB

- **oracledb в Thin Mode** — Oracle Client не требуется
- **Named parameters**: `:date_param`, `:data_acc`, `:data_cur`
- **Формат даты**: `'DD.MM.YYYY'`

#### Динамический IN-клаус

Для передачи списка значений в `IN (...)`:
```python
numbers = ["c11132", "c11133", "954521482"]
placeholders = ", ".join(f":v{i}" for i in range(len(numbers)))
sql = sql.replace(":data_number", placeholders)
params = {f"v{i}": v for i, v in enumerate(numbers)}
df = query(sql, params)
```

### Excel Integration

#### Чтение параметров

**Именованные ячейки:**
```python
wb = xw.Book.caller()
rdate = wb.names['RDATE'].refers_to_range.value  # datetime или None
```

**Параметры из таблицы tParam:**
```python
table = sheet.api.ListObjects("tParam")
df = sheet.range(table.Range.Address).options(pd.DataFrame, header=1, index=False).value
path_row = df[df['Параметр'] == 'Path_DA7X']
path = path_row.iloc[0]['Значение']
```

#### Вставка данных в Excel

**paste_to_excel()** — стандартная стратегия:
- Очищает тело таблицы, меняет размер, вставляет данные
- Заменяет NaN на пустые строки
- Отключает screen_updating и calculation на время операции

**paste_to_excel_smart()** — для таблиц, размещённых вертикально:
- Добавляет/удаляет строки по одной
- Не трогает соседние таблицы
- Не заменяет NaN (могут появиться None в Excel)

#### Продвинутое форматирование через COM API

```python
table = sheet.api.ListObjects("TableName")
start_row = table.HeaderRowRange.Row + 1
start_col = table.Range.Column

for i, (idx, row) in enumerate(df.iterrows()):
    excel_row = start_row + i  # использовать i, а не idx!
    row_range = sheet.range((excel_row, start_col)).resize(1, num_columns)
    
    if row['mark'] == 'pre_excluded':
        row_range.api.Font.Color = 0xA6A6A6  # BGR формат!
        row_range.api.Font.Strikethrough = True
    elif row['mark'] == 'exclude':
        row_range.api.Font.Color = 0x0000FF  # красный в BGR
        row_range.api.Font.Bold = True
```

**Важно:** Excel использует **BGR** формат цвета, не RGB: `0xBBGGRR`

---

## Логирование

Стандартный паттерн для fetchers:
```python
ENABLE_LOGGING = False  # включить/выключить на уровне модуля

def _setup_logger():
    if not ENABLE_LOGGING:
        return logging.getLogger("module_name_disabled")
    logger = logging.getLogger("module_name")
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    log_dir = os.path.abspath(os.path.join(script_dir, '..', 'logs'))
    os.makedirs(log_dir, exist_ok=True)
    handler = logging.FileHandler(os.path.join(log_dir, 'module_name.log'), encoding='utf-8')
    handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logger.addHandler(handler)
    return logger

logger = _setup_logger()
```

---

## Обработка ошибок

**try-finally для восстановления настроек Excel:**
```python
app.screen_updating = False
app.calculation = 'manual'
try:
    # операции с Excel
finally:
    app.calculation = 'automatic'
    app.screen_updating = True
```

**Важно:** `paste_to_excel()` не использует try-finally внутри — caller отвечает за восстановление настроек.

---

## Конвенции

- **Язык комментариев**: русский
- **Кодировка файлов**: UTF-8 (обязательно для кириллицы в SQL и логах)
- **Именование функций**: `run_<action>()` для main.py, `fetch_to_<name>()` для получения данных, `paste_to_excel_<name>()` для вставки
- **SQL-шаблоны**: суффикс `_template.sql`, параметры через `:param_name`
- **Config в chart-модулях**: все магические числа выносить в класс Config

---

## Отладка и разработка

### Forecast-режим

Система поддерживает два режима:
- **Normal**: используется предыдущий рабочий день (`get_previous_working_day()`)
- **Forecast**: дата из именованной ячейки `ForecastDate` (если не None)

`forecast_date()` возвращает **raw datetime** из Excel (не строку).

### Консольный вывод для chart-модулей

Для корректного отображения кириллицы в консоли Windows:
```python
if sys.platform == "win32":
    os.system("chcp 65001 > NUL")
    sys.stdout.reconfigure(encoding='utf-8')
```

Минимальный console-лог обязателен для chart-модулей:
```python
print("=== СОЗДАНИЕ ГРАФИКА ===")
print("1. Получение данных...")
print("2. Построение...")
print("✅ График создан!")
```

---

## Полезные ссылки

- **CLAUDE.md** — подробная документация по паттернам и примерам кода
- **TASKS_chart_7s_new.md** — детальная спецификация графика динамики рыночного риска
