import argparse
import logging
import sqlite3
import sys
from datetime import datetime
from pathlib import Path
from typing import Iterable, List, Optional

import pandas as pd

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")


DEFAULT_SOURCE_DIR = Path(
    r"r:\Подразделения\РИСК-менеджмент\Внутренние\3 - РИСК ЛИКВИДНОСТИ\DB_LCR\6K"
)
DEFAULT_DB_PATH = Path(
    r"r:\Подразделения\РИСК-менеджмент\Внутренние\3 - РИСК ЛИКВИДНОСТИ\DB_LCR\liquidity_data.db"
)
EXPECTED_COLUMNS = ["REC_NO", "EKP", "R030", "T100"]
SOURCE_EXTENSIONS = {".xls", ".xlsx", ".xlsm"}
# Соответствие кодов EKP колонкам в таблице LCR_Combined
LCR_SOURCE_CODES = {
    "A6K081": "LCRвв",
    "A6K082": "LCRів",
}


def configure_logger(verbose: bool, log_to_file: bool) -> logging.Logger:
    """Создает логгер, пишущий в консоль и при необходимости в logs/entry_db_6kx_batch.log. Возвращает единый logging.Logger для всех функций загрузчика."""
    logger = logging.getLogger("entry_db_6kx_batch")
    logger.setLevel(logging.DEBUG if verbose else logging.INFO)

    if not logger.handlers:
        formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

        stream_handler = logging.StreamHandler(sys.stdout)
        stream_handler.setFormatter(formatter)
        logger.addHandler(stream_handler)

        if log_to_file:
            log_dir = Path(__file__).resolve().parent.parent / "logs"
            log_dir.mkdir(parents=True, exist_ok=True)
            file_handler = logging.FileHandler(
                log_dir / "entry_db_6kx_batch.log", encoding="utf-8"
            )
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)

    return logger


def discover_excel_files(
    source_dir: Path, pattern: str, recursive: bool
) -> List[Path]:
    """Сканирует каталог (при желании рекурсивно) и возвращает Excel-файлы по маске. Результат отсортирован, что облегчает воспроизводимость загрузок."""
    iterator: Iterable[Path]
    if recursive:
        iterator = source_dir.rglob(pattern)
    else:
        iterator = source_dir.glob(pattern)

    files = [
        path
        for path in iterator
        if path.is_file() and path.suffix.lower() in SOURCE_EXTENSIONS
    ]
    files.sort()
    return files


def extract_report_date(file_path: Path) -> str:
    """Извлекает дату отчета (ГГГГ-ММ-ДД) из имени файла формата 6K_DDMMYYYY.xlsx. Бросает ValueError, если дата отсутствует или записана некорректно."""
    parts = file_path.stem.split("_")
    if len(parts) < 2:
        raise ValueError("Файл не содержит дату в имени (ожидается 6K_DDMMYYYY.xlsx)")

    digits = "".join(ch for ch in parts[1] if ch.isdigit())
    if len(digits) < 8:
        raise ValueError("Не могу прочитать дату (формат DDMMYYYY) из имени файла")

    date_obj = datetime.strptime(digits[:8], "%d%m%Y")
    return date_obj.strftime("%Y-%m-%d")


def read_source_dataframe(file_path: Path) -> pd.DataFrame:
    """Считывает Excel-файл, пропуская служебные строки и приводя столбцы к строковому типу. Повторяет набор параметров одиночного скрипта для единообразного поведения."""
    return pd.read_excel(file_path, skiprows=8, dtype=str)


def validate_dataframe(df: pd.DataFrame) -> Optional[str]:
    """Проверяет наличие обязательных колонок и непустых значений EKP; возвращает описание ошибки. При успешной проверке возвращает None и данные можно использовать далее."""
    missing_columns = [col for col in EXPECTED_COLUMNS if col not in df.columns]
    if missing_columns:
        return f"Нет обязательных колонок: {', '.join(missing_columns)}"

    if df.empty or df["EKP"].isna().all():
        return "В файле нет данных по колонке EKP"

    return None


def calculate_r031(value: str) -> str:
    """Повторяет логику расчета R031 из исходного Excel-скрипта (980->NV, #->#, иначе FCY). Результат напрямую используется при построении набора DB_6KX."""
    if str(value) == "980":
        return "NV"
    if str(value) == "#":
        return "#"
    return "FCY"


def build_combined_dataframe(df: pd.DataFrame, file_date: str) -> pd.DataFrame:
    """Формирует набор строк для DB_6KX: добавляет дату и вычисляет служебную колонку R031. Возвращает DataFrame с итоговым набором колонок, готовым к записи в SQLite."""
    subset = df[EXPECTED_COLUMNS].copy()
    subset["Date"] = file_date
    subset["R031"] = subset["R030"].apply(calculate_r031)
    columns = ["Date", "REC_NO", "EKP", "R030", "R031", "T100"]
    return subset[columns]


def normalize_numeric(value) -> Optional[float]:
    """Переводит текстовые представления чисел (с пробелами, запятыми, NBSP) к float. Если преобразование невозможно, возвращает None для безопасной обработки."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    text = str(value).strip().replace(" ", "").replace("\u00A0", "")
    if not text:
        return None
    text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def build_lcr_row(df_combined: pd.DataFrame, file_date: str, logger: logging.Logger):
    """Готовит единственную строку для LCR_Combined на основе EKP A6K081/A6K082 и нормированного T100. Возвращает словарь, который далее конвертируется в DataFrame перед записью."""
    lcr_row = {"Date": file_date, "LCRвв": None, "LCRів": None, "Min_NRM": 1.00, "Target": 1.10}

    for ekp_code, column_name in LCR_SOURCE_CODES.items():
        ekp_slice = df_combined[df_combined["EKP"] == ekp_code]
        if ekp_slice.empty:
            continue
        numeric_value = normalize_numeric(ekp_slice.iloc[0]["T100"])
        if numeric_value is None:
            logger.warning(
                "Не удалось преобразовать T100 для EKP %s (файл %s)", ekp_code, file_date
            )
            continue
        lcr_row[column_name] = numeric_value / 100

    return lcr_row


def check_required_tables(db_path: Path, logger: logging.Logger) -> bool:
    """Проверяет наличие таблиц DB_6KX и LCR_Combined в базе перед записью данных. Ведет подробный лог и прекращает обработку при отсутствии структуры."""
    required_tables = ["DB_6KX", "LCR_Combined"]
    try:
        with sqlite3.connect(db_path) as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            )
            existing = {row[0] for row in cursor.fetchall()}
    except Exception as exc:
        logger.error("Не удалось проверить таблицы в БД: %s", exc)
        return False

    missing = [table for table in required_tables if table not in existing]
    if missing:
        logger.error("В БД отсутствуют таблицы: %s", ", ".join(missing))
        return False

    logger.info("Требуемые таблицы найдены: %s", ", ".join(required_tables))
    return True


def is_date_already_loaded(conn: sqlite3.Connection, file_date: str) -> bool:
    """Проверяет, существует ли дата в LCR_Combined, чтобы избежать повторной загрузки. Используется опцией --skip-existing."""
    cursor = conn.cursor()
    cursor.execute(
        "SELECT 1 FROM LCR_Combined WHERE Date = ? LIMIT 1",
        (file_date,),
    )
    return cursor.fetchone() is not None


def process_file(
    file_path: Path,
    file_date: str,
    conn: sqlite3.Connection,
    logger: logging.Logger,
    dry_run: bool,
) -> bool:
    """Обрабатывает один Excel-файл: читает, проверяет и при необходимости записывает данные в таблицы DB_6KX и LCR_Combined. Учитывает флаг dry-run и детально логирует каждый этап."""
    logger.info("Обработка файла %s (дата %s)", file_path.name, file_date)
    try:
        df = read_source_dataframe(file_path)
    except Exception as exc:
        logger.error("Ошибка при чтении файла %s: %s", file_path, exc)
        return False

    validation_error = validate_dataframe(df)
    if validation_error:
        logger.error("Пропускаю %s: %s", file_path, validation_error)
        return False

    df_combined = build_combined_dataframe(df, file_date)
    lcr_row = build_lcr_row(df_combined, file_date, logger)

    if dry_run:
        logger.info("Режим dry-run: запись в БД пропущена для %s", file_path.name)
        return True

    try:
        df_combined.to_sql("DB_6KX", conn, if_exists="append", index=False)
        pd.DataFrame([lcr_row]).to_sql("LCR_Combined", conn, if_exists="append", index=False)
    except Exception as exc:
        logger.error("Ошибка при записи данных из %s: %s", file_path, exc)
        return False

    logger.info("Файл %s успешно загружен", file_path.name)
    return True


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Массовая загрузка файлов 6KX в базу данных liquidity_data.db"
    )
    parser.add_argument(
        "--source",
        "-s",
        type=Path,
        default=DEFAULT_SOURCE_DIR,
        help=f"Каталог с отчётами 6KX (по умолчанию {DEFAULT_SOURCE_DIR})",
    )
    parser.add_argument(
        "--db",
        "-d",
        type=Path,
        default=DEFAULT_DB_PATH,
        help=f"Путь к SQLite базе (по умолчанию {DEFAULT_DB_PATH})",
    )
    parser.add_argument(
        "--pattern",
        "-p",
        default="6КХ_*.xls*",
        help="Маска поиска файлов (по умолчанию 6КХ_*.xls*)",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Сканировать вложенные каталоги",
    )
    parser.add_argument(
        "--skip-existing",
        action="store_true",
        help="Пропускать файлы, если дата уже есть в LCR_Combined",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Читать файлы, но не выполнять запись в БД",
    )
    parser.add_argument(
        "--no-file-log",
        action="store_true",
        help="Не писать лог в файл, только в консоль",
    )
    parser.add_argument(
        "--verbose",
        "-v",
        action="store_true",
        help="Выводить отладочную информацию",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    logger = configure_logger(verbose=args.verbose, log_to_file=not args.no_file_log)

    source_dir = args.source.expanduser()
    db_path = args.db.expanduser()

    if not source_dir.exists():
        logger.error("Каталог с файлами не найден: %s", source_dir)
        return 1

    if not db_path.exists():
        logger.error("База данных не найдена: %s", db_path)
        return 1

    files = discover_excel_files(source_dir, args.pattern, args.recursive)
    if not files:
        logger.warning("В каталоге %s нет файлов по шаблону %s", source_dir, args.pattern)
        return 1

    if not check_required_tables(db_path, logger):
        return 1

    processed = 0
    skipped = 0
    failed: List[Path] = []

    with sqlite3.connect(db_path) as conn:
        files_with_dates = []
        for file_path in files:
            try:
                file_date = extract_report_date(file_path)
            except ValueError as exc:
                logger.error("Пропускаю %s: %s", file_path, exc)
                failed.append(file_path)
                continue
            files_with_dates.append((file_date, file_path))

        files_with_dates.sort(key=lambda item: item[0])

        for file_date, file_path in files_with_dates:
            if args.skip_existing and is_date_already_loaded(conn, file_date):
                logger.info(
                    "Дата %s уже есть в LCR_Combined, пропускаю файл %s",
                    file_date,
                    file_path.name,
                )
                skipped += 1
                continue

            if process_file(file_path, file_date, conn, logger, args.dry_run):
                processed += 1
            else:
                failed.append(file_path)

    logger.info(
        "Готово: обработано=%s, пропущено=%s, с ошибками=%s",
        processed,
        skipped,
        len(failed),
    )

    if failed:
        logger.error(
            "Не удалось загрузить файлы:\n%s",
            "\n".join(f" - {path}" for path in failed),
        )
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
