# 📊 Система автоматизированной выгрузки данных и отчетности из SR-bank

⛏ **Назначение проекта**
Данный проект автоматизирует процесс получения данных из базы данных **Oracle**,
обработки результатов и формирования Excel-отчетов.

Проект интегрируется с **Excel через xlwings** и может вызываться прямо из книги Excel.

---

## 🏗 Архитектура проекта

```
Get_data_SR/
├── main.py                  # Главный вход: набор функций-запускателей
│
├── db/                      # Модули для подключения к БД Oracle
│   ├── connect_db_oracle.py # Логика соединения
│   └── oracle.py            # Вспомогательные методы работы с Oracle
│
├── fetchers/                # Модули-загрузчики для разных отчетов
│   ├── balance_nrk.py       # Прогнозный баланс для НРК
│   ├── banks_42x.py         # Межбанковские операции (форма 42X)
│   ├── compens_579.py       # Компенсация по форме 579
│   ├── detail_6jx.py        # Детализация резервов 6JX
│   ├── diff_acc.py          # Динамика остатков по счетам
│   ├── doc_acc.py           # Документы по заданному счету
│   ├── dz_spot.py           # Сделки по ДЗ и СПОТ
│   ├── dz_spot_diff.py      # Динамика по ДЗ и СПОТ
│   ├── fz_ccf_6jx.py        # Отчет по фин.обязат. (9 кл.) для 6JX
│   ├── grp_9000.py          # Выгрузка 900% и 9129 для НРК
│   ├── rc_component.py      # Компоненты регуляторного капитала
│   └── secur_doc.py         # Ценные бумаги — сделки
│
├── charts/                  # Модули для создания графиков и визуализаций
│   ├── chart_as.py          # Спидометр Acerbi-Szekely (старая версия)
│   ├── chart_as_v2.py       # Двойной спидометр Acerbi-Szekely Z2 (1d/10d)
│   └── chart_es.py          # График распределения потерь VaR/ES
│
├── sql/                     # SQL-шаблоны для запросов
│   ├── SR_6JX_FZ_CCF_template.sql      # Фин.обязат. (9 кл.) с CCF для 6JX
│   ├── SR_6JX_Reserve_template.sql     # Резервы 6JX
│   ├── SR_BALANCE_NRK_template.sql     # Баланс для НРК
│   ├── SR_BANKS_42X_template.sql       # Межбанковские операции (форма 42X)
│   ├── SR_CHECK_9000_template.sql      # 900% и 9129 для НРК
│   ├── SR_CHECK_DZ_SPOT_template.sql   # Сделки по ДЗ и СПОТ
│   ├── SR_COMPENSATION_579_template.sql # Компенсация по форме 579
│   ├── SR_DIFF_ACC_template.sql        # Динамика остатков по счетам
│   ├── SR_DIFF_DZ_SPOT_template.sql    # Динамика по ДЗ и СПОТ
│   ├── SR_DOC_ACC_template.sql         # Документы по счету
│   ├── SR_RC_component_template.sql    # Компоненты РК
│   └── SR_SECUR_DOC_template.sql       # Сделки ЦБ
│
├── utils/                   # Вспомогательные модули
│   ├── date_utils.py        # Работа с датами (рабочие дни, прогноз)
│   ├── excel_writer.py      # Экспорт в Excel через xlwings
│   └── path_utils.py        # Пути и директории
│
├── mail/                    # Модули для email-рассылки
│   ├── forecast_nrk.py      # Генерация и отправка прогнозных писем НРК
│   └── forecast_nrk_OLD.py  # Устаревшая версия (архив)
│
└── watcher/                 # Модули для мониторинга файлов на сетевых дисках
    ├── check_copy_v1.py     # Базовая версия мониторинга
    ├── check_copy_v2_claude.py # Расширенная версия мониторинга
    ├── config/              # Конфигурация мониторинга
    │   ├── __init__.py
    │   ├── settings.py      # Основные настройки
    │   └── watch_rules.py   # Правила отслеживания файлов
    └── core/                # Ядро системы мониторинга
        ├── __init__.py
        ├── file_handler.py  # Обработчик файловых событий
        ├── utils.py         # Вспомогательные функции
        └── watcher.py       # Основной класс наблюдателя
```

---

## ⚙ Установка

Установка и исользование реализуется через **conda** v25.5.1

Операции после установки **conda**:

1. **Отключить проверку SSL через терминал conda**
```bash
conda config --set ssl_verify false
#Проверка:
conda config --show ssl_verify
```

2. **Добавление прокси для HTTP**
```bash
conda config --set proxy_servers.http http://inetsvc.radabank.com.ua:8080
conda config --set proxy_servers.https http://inetsvc.radabank.com.ua:8080
#Проверка итоговых настроек
conda config --show proxy_servers
```

3. **Настройка основного канала conda-forge:**
```bash
conda config --add channels conda-forge
conda config --set channel_priority strict
#Проверяем текущие каналы:
conda config --show channels
```

4. **Проверка корректности установки Python:**
```bash
python --version
conda --version
conda list
```

5. **Установка Spyder с полным набором расширений:**
```bash
conda install spyder spyder-kernels qtpy pyqt
```

6. **Установка Jupyter Notebook, JupyterLab и расширений**
```bash
conda install notebook jupyterlab jupyterlab-language-pack-ru-RU jupyterlab_widgets ipywidgets
jupyter_contrib_nbextensions
jupyter contrib nbextension install --user
jupyter nbextensions_configurator enable --user
pip install spyder-notebook
```

7. **Установка xlwings через conda**
```bash
conda install xlwings
conda list xlwings
setx XLWINGS_LICENSE_KEY noncommercial
xlwings addin install
```

```python
#Тестовый скрипт
import xlwings as xw
app = xw.App(visible=True)
wb = app.books.add()
wb.sheets[0]['A1'].value = "Тест xlwings — всe работает!"
```

8. **Установить зависимости**
В проекте используются следующие обязательные внешние зависимости:
- `oracledb` — подключение к базе данных Oracle
- `pandas` — обработка табличных данных
- `pywin32` — доступ к широкому спектру функций Windows API и компонентам COM (Component Object Model) из Python

*Опционально:*
- `numpy` — поддержку многомерных массивов и матриц, а также широкий набор высокоуровневых математических функций
- `matplotlib` — для визуализации данных

```bash
conda install pandas oracledb pywin32 numpy matplotlib
```

9. **Настройка подключения к Oracle**
В модуле `db/connect_db_oracle.py` авторизация к базе Oracle реализована через конфигурационный JSON-файл, расположенный у пользователя по пути:
`~\.conda\db_ac.json`

📂 Формат **db_ac.json**

```json
{
    "user": "ВАШ_ЛОГИН",
    "password": "ВАШ_ПАРОЛЬ",
    "dsn": "RB-RDB1.radabank.com.ua:1521/srbank"
}
```
🔑 *Особенности:*
Файл должен быть создан и обновлён самим пользователем.

Путь формируется через `os.path.expanduser`, что гарантирует правильное определение домашней директории независимо от системы.

Авторизация происходит в Thin Mode Oracle, что не требует установки полной Oracle Client.

📌 *Логика:*
Определяется путь к `~\.conda\db_ac.json`.

Загружаются учетные данные (user, password, dsn) в словарь.

Выполняется подключение через `oracledb.connect(...)`.

При успешном подключении выводится сообщение "✅ Connected using JSON config in .conda".

---

## 🚀 Запуск

### 📌 Через Python
```bash
python main.py
```
или вызов отдельных функций:
```python
from main import run_secur_doc, run_balance_nrk

run_balance_nrk()
```

### 📌 Через Excel (xlwings)
- Открыть книгу Excel, привязанную к проекту
- Вызывать макросы `run_*` напрямую

---

## 🔍 Описание основных функций (main.py)

| Функция                | Описание |
|------------------------|----------|
| `run_secur_doc()`      | Выгрузка и экспорт данных по ценным бумагам |
| `run_grp_9000()`       | Формирование отчета по группе 9000 (900% и 9129 для НРК) |
| `run_dz_spot()`        | Данные по сделкам ДЗ и СПОТ |
| `run_balance_nrk()`    | Прогнозный баланс НРК |
| `run_diff_spot()`      | Динамика по сделкам ДЗ и СПОТ |
| `run_forecast_mail()`  | Генерация и отправка email с прогнозом НРК |
| `run_diff_acc()`       | Динамика остатков по счетам |
| `run_fz_ccf_6jx()`     | Отчет по финансовым обязательствам (9 кл.) с CCF для 6JX |
| `run_doc_acc()`        | Документы по заданному счету |
| `run_plot_var_es()`    | Построение графика распределения потерь VaR/ES |
| `run_compens_579()`    | Выгрузка данных по компенсации 579 |
| `run_6jx_reserve()`    | Выгрузка детализации резервов 6JX |
| `run_42x_banks()`      | Выгрузка межбанковских операций (форма 42X) |
| `run_rc_comp()`        | Выгрузка компонентов регуляторного капитала |
| `run_plot_as()`        | Построение двойного спидометра Acerbi-Szekely Z2 (1d/10d) |

---

## 🔄 Логика работы

### 📊 Выгрузка данных из БД
```
┌─────────────────┐
│  Excel / Python │
└───────┬─────────┘
        │ вызов run_*
        ▼
┌─────────────────────────────┐
│ main.py                     │
│  - диспетчер функций        │
└───────┬─────────────────────┘
        ▼
┌─────────────────────────────┐
│ fetchers/<report>.py        │
│  - загружает SQL-шаблон     │
│  - подставляет параметры    │
│  - выполняет запрос к Oracle│
│  - получает DataFrame       │
└───────┬─────────────────────┘
        ▼
┌─────────────────────────────┐
│ utils/excel_writer.py       │
│  - записывает данные в      │
│    именованную таблицу Excel│
│    через xlwings            │
└─────────────────────────────┘
```

### 📧 Отправка email-отчетов
```
┌─────────────────┐
│  Excel / Python │
└───────┬─────────┘
        │ вызов run_forecast_mail()
        ▼
┌─────────────────────────────┐
│ mail/forecast_nrk.py        │
│  - собирает данные баланса  │
│  - формирует письмо         │
│  - отправляет email         │
└─────────────────────────────┘
```

### 📈 Построение графиков
```
┌─────────────────┐
│  Excel / Python │
└───────┬─────────┘
        │ вызов run_plot_*
        ▼
┌─────────────────────────────┐
│ charts/<chart>.py           │
│  - читает данные из Excel   │
│  - строит график matplotlib │
│  - сохраняет в temp файл    │
│  - вставляет в Excel        │
└─────────────────────────────┘
```

### 👁 Мониторинг файлов
```
┌─────────────────────────────┐
│ watcher/check_copy_v2.py    │
│  - отслеживает сетевые папки│
│  - копирует новые файлы     │
│  - применяет правила фильтрации│
└─────────────────────────────┘
```

---

## 🛠 Технические детали

- **Язык:** Python 3.12
- **БД:** Oracle 19с
- **MS Office Excel:** 2021
- **Формат отчетов:** Excel (.xlsx / .xlsm)
- **Интеграция:** xlwings 0.33.15
---
