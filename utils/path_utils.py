# Импортируем класс Path из модуля pathlib для удобной работы с файловыми путями
from pathlib import Path

# Определяем функцию, которая возвращает полный путь к файлу в папке "sql"
def get_sql_path(filename: str) -> Path:
    # Получаем путь к текущему файлу (__file__), переводим его в абсолютный путь (resolve())
    # Затем получаем родительскую папку (parent), ещё раз родительскую папку (parent.parent)
    # После этого добавляем подпапку "sql" и имя файла
    return Path(__file__).resolve().parent.parent / "sql" / filename
