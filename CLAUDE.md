# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Setup

### Database Credentials

Oracle connection reads credentials from `~\.conda\db_ac.json`. Create this file before running any data fetcher:

```json
{
  "user": "your_username",
  "password": "your_password",
  "dsn": "host:port/service_name"
}
```

Connection uses `oracledb` in Thin Mode (no Oracle Client required).

### Running Scripts

All fetchers require an active Excel workbook as caller. Scripts **cannot run standalone** — they must be invoked from Excel via:
```vba
RunPython "import main; main.run_<module_name>()"
```

To test a fetcher in isolation, open the target `.xlsm` workbook in Excel and trigger the macro from there.

## Project Overview

This is a Python-based liquidity risk management system that integrates with Excel via xlwings. The project fetches data from an Oracle database (SR_BANK schema), processes it, and inserts it into Excel tables for analysis. It also includes charting capabilities for visualizing asset/equity structure (AS/ES) data.

## Architecture

### Core Data Flow Pattern

All data fetchers follow a consistent three-step pattern:

1. **SQL Template**: Read parameterized SQL query from `sql/` directory
2. **Database Query**: Execute query via `db/oracle.query()` with date parameters
3. **Excel Export**: Write resulting DataFrame to named Excel table via `utils/excel_writer.paste_to_excel()`

Example:
```python
# fetchers/balance_nrk.py
def fetch_to_balance_nrk():
    sql_path = get_sql_path("SR_BALANCE_NRK_template.sql")
    sql = open(sql_path).read().strip().rstrip(";")
    date_param = forecast_date() or get_previous_working_day()
    return query(sql, {"date_param": date_param})

def paste_to_excel_balance_nrk():
    df = fetch_to_balance_nrk()
    paste_to_excel("Нрк_TEST", "DB_test_NRK", df)
```

### Directory Structure

- **`main.py`**: Exposes `run_*()` functions callable from Excel macros via xlwings
- **`fetchers/`**: Data fetching modules that query Oracle and prepare DataFrames
- **`db/`**: Database utilities
  - `oracle.py`: Oracle query execution (returns pandas DataFrames)
  - `connect_db_oracle.py`: Oracle connection management
  - `entry_db_6kx.py`: Process and insert 6KX files into SQLite database
  - `batch_entry_db_6kx.py`: Batch processing of multiple 6KX files
- **`sql/`**: SQL query templates with named parameters (e.g., `:date_param`)
- **`utils/`**: Shared utilities
  - `excel_writer.py`: Two strategies for writing DataFrames to Excel tables
  - `date_utils.py`: Working day calculations and forecast date handling
  - `path_utils.py`: Path resolution for SQL templates
  - `parser_forex.py`: Text parser for extracting forex deal numbers from DESCRIPTION fields (numbers starting with `c`/`с` or `9`; normalizes Cyrillic/Ukrainian `с` to ASCII `c`)
- **`charts/`**: Chart generation modules (AS/ES analysis, trading charts)
- **`request/`**: External data requests (e.g., fair price OVDP from web sources)

### Excel Integration via xlwings

Functions in `main.py` are exposed to Excel with the pattern:
```python
def run_<module_name>():
    """Description of what this does."""
    paste_to_excel_<module_name>()
```

These are called from Excel VBA macros using: `RunPython "import main; main.run_<module_name>()"`

### Date Handling Logic

The system supports two modes:
- **Normal mode**: Uses previous working day (`get_previous_working_day()`)
- **Forecast mode**: Uses date from Excel named range `ForecastDate` if set

Most fetchers check `forecast_date()` first, falling back to previous working day if None. `forecast_date()` returns a raw Excel `datetime` object (not a formatted string) or `None` if the cell is empty.

### Excel Table Writing Strategies

Two functions in `utils/excel_writer.py`:

1. **`paste_to_excel()`**: Default strategy
   - Clears table, resizes, and inserts new data
   - Turns off screen updating and calculations during operation
   - Replaces `NaN` with empty strings (`fillna('')`) before writing
   - Preferred for most use cases

2. **`paste_to_excel_smart()`**: For vertically stacked tables
   - Adds/removes rows individually to avoid disrupting adjacent tables
   - Use when multiple Excel tables are placed one under another
   - Does **not** replace `NaN` — passes raw `df.values.tolist()`, so `None` may appear in Excel cells if DataFrame contains nulls

### Database Workflows

**Oracle (SR_BANK schema)**:
- Primary data source for regulatory reporting data
- All SQL templates use named parameters (`:date_param`, etc.)
- Connection managed via `db/connect_db_oracle.py`

**SQLite (6KX files)**:
- Used for processing and storing 6KX regulatory report files
- `entry_db_6kx.py`: Process single file from Excel (via named table `tPathF6KX`)
- `batch_entry_db_6kx.py`: Process multiple files from directory
- Logs to `logs/entry_db_6kx.log`

## Working with the Code

### Adding a New Data Fetcher

1. Create SQL template in `sql/SR_<NAME>_template.sql` with named parameters
2. Create fetcher in `fetchers/<name>.py`:
   ```python
   from db.oracle import query
   from utils.date_utils import get_previous_working_day, forecast_date
   from utils.excel_writer import paste_to_excel  # or paste_to_excel_smart
   from utils.path_utils import get_sql_path

   def fetch_to_<name>():
       sql_path = get_sql_path("SR_<NAME>_template.sql")
       sql = open(sql_path, encoding="utf-8").read().strip().rstrip(";")
       date_param = forecast_date() or get_previous_working_day()
       return query(sql, {"date_param": date_param})

   def paste_to_excel_<name>():
       df = fetch_to_<name>()
       # Use paste_to_excel_smart if tables are vertically stacked
       paste_to_excel("<SheetName>", "<TableName>", df)
   ```
3. Add to `main.py`:
   ```python
   from fetchers.<name> import paste_to_excel_<name>

   def run_<name>():
       """Description."""
       paste_to_excel_<name>()
   ```

**Note:** Use `paste_to_excel_smart` instead of `paste_to_excel` when multiple tables are placed vertically on the same sheet to avoid disrupting adjacent tables. Unlike `paste_to_excel`, `paste_to_excel_smart` does **not** disable `screen_updating` or `calculation` during execution.

### Modifying SQL Queries

- All SQL templates are in `sql/` directory with `.sql` extension
- Use Oracle SQL syntax with named parameters (`:param_name`)
- Common parameters: `:date_param` (format: 'DD.MM.YYYY')
- Query the SR_BANK schema tables (ACCOUNT, ACCOUNT_SNAPSHOT, DOCUMENT, CURRENCY, etc.)
- Some queries accept per-row parameters (e.g., `:data_acc`, `:data_cur`) that are passed in a loop from Python when iterating over a DataFrame

#### Dynamic IN-clause for Multi-value Queries

`oracledb` does not support binding a Python list to a single `:param` in an `IN (...)` clause. Build the placeholders dynamically and replace before executing:

```python
numbers = ["c11132", "c11133", "954521482"]
placeholders = ", ".join(f":v{i}" for i in range(len(numbers)))
sql = sql.replace(":data_number", placeholders)   # :data_number in template
params = {f"v{i}": v for i, v in enumerate(numbers)}
df = query(sql, params)
```

Use a single placeholder name (e.g., `:data_number`) in the SQL template to mark where the list should be injected.

### Reading Report Date Directly vs. Forecast Date

Most fetchers use `forecast_date() or get_previous_working_day()`. However, fetchers that need a specific **report date** (not forecast) read `RDATE` from the Excel named range directly:

```python
wb = xw.Book.caller()
rdate = wb.names['RDATE'].refers_to_range.value  # Returns datetime or string
```

Use this pattern when the operation is tied to a specific reporting date regardless of forecast mode (e.g., `detail_6sx.py`).

### Working with Charts

Chart modules in `charts/` generate matplotlib/plotly visualizations and insert them into Excel:
- `chart_as*.py`: Asset structure analysis charts
- `chart_es*.py`: Equity structure and VAR analysis charts
- All use xlwings to insert images/charts into specific Excel ranges

### Advanced Excel Operations

#### Excel COM API Formatting

For advanced Excel formatting (fonts, colors, styles), access Excel COM objects via `.api` property:

```python
# Access Excel table via COM
table = sheet.api.ListObjects("TableName")
start_row = table.HeaderRowRange.Row + 1
start_col = table.Range.Column

# Format specific row
row_range = sheet.range((excel_row, start_col)).resize(1, num_columns)
row_range.api.Font.Color = 0x0000FF  # Red in BGR format
row_range.api.Font.Bold = True
row_range.api.Font.Strikethrough = True
```

**Important notes:**
- Excel uses **BGR color format**, not RGB: `0xBBGGRR` (e.g., red = `0x0000FF`, gray = `0xA6A6A6`)
- When iterating over filtered DataFrames to apply formatting, use `enumerate()` to get sequential row numbers:
  ```python
  for i, (idx, row) in enumerate(df.iterrows()):
      excel_row = start_row + i  # Use i, not idx
  ```
  The `idx` from `iterrows()` preserves original DataFrame indices, which may not be sequential after filtering.

#### Reading Excel Named Ranges

To read named cells/ranges from Excel:
```python
wb = xw.Book.caller()
value = wb.names['NAMED_CELL'].refers_to_range.value
```

Common named ranges in this project:
- `RDATE`: Report date on "menu" sheet
- `ForecastDate`: Forecast date for forecast mode

### Fetchers with Conditional Logic

Some fetchers (e.g., `detail_6sx.py`) perform multi-step data processing:

1. Execute multiple SQL queries
2. Apply conditional logic to mark/filter records
3. Write to multiple Excel tables
4. Apply conditional formatting

**Critical pattern for conditional logic:**
- Apply conditions in priority order, checking `mark.isna()` for subsequent conditions
- Example from `detail_6sx.py`:
  ```python
  df['mark'] = None
  # 1. Highest priority: pre-excluded accounts
  df.loc[condition1, 'mark'] = 'pre_excluded'
  # 2. Lower priority: only mark if not already marked
  df.loc[condition2 & (df['mark'].isna()), 'mark'] = 'exclude'
  # 3. Lowest priority: only mark if still not marked
  df.loc[condition3 & (df['mark'].isna()), 'mark'] = 'exclude'
  ```

### Fetchers That Depend on Other Fetchers

Some fetchers reuse data prepared by another fetcher rather than querying Oracle directly. Example: `pay_6sx.py` calls `fetch_6sx_data()` from `detail_6sx` to obtain the filtered account list, then runs per-row SQL queries against that list.

```python
from fetchers.detail_6sx import fetch_6sx_data

def fetch_pay_6sx_data():
    acc_calc, _ = fetch_6sx_data()   # переиспользуем результат другого фетчера
    for _, row in acc_calc.iterrows():
        df = query(sql, {"data_acc": row['ACCOUNT_NUMBER'], ...})
        results.append(df)
    return pd.concat(results, ignore_index=True)
```

This pattern is used when the output of one SQL pipeline feeds into a second query as a parameter list. Dependencies can be multi-level:

```
detail_6sx → pay_6sx → forex_6sx
```

Calling `run_forex_6sx()` implicitly re-runs both `detail_6sx` and `pay_6sx` queries.

## Common Patterns

### UDF (User-Defined Function) for Excel

`udf_modules.py` contains xlwings UDFs callable directly from Excel formulas:
```python
@xw.func
def py_RoundLR(data, threshold):
    # Returns 0 if absolute value doesn't exceed threshold
```

### Logging

Database entry modules use Python logging to `logs/` directory with UTF-8 encoding for Cyrillic characters.

**Pattern for optional logging:**
```python
ENABLE_LOGGING = False  # Toggle at module level

def _setup_logger():
    if not ENABLE_LOGGING:
        return logging.getLogger("module_name_disabled")
    # ... full logger setup

logger = _setup_logger()
```

This allows easy enable/disable of logging for performance without removing logging code.

### Error Handling in Excel Operations

Always wrap xlwings operations in try-finally blocks to restore Excel settings:
```python
app.screen_updating = False
app.calculation = 'manual'
try:
    # Excel operations
finally:
    app.calculation = 'automatic'
    app.screen_updating = True
```

## Conventions

- Code comments are written in **Russian**
- All file I/O uses `encoding="utf-8"` (required for Cyrillic in SQL files and logs)
- In-progress development tasks are tracked in `TASKS.md` at the project root
