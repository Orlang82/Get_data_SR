# Импортируем библиотеку pandas для работы с таблицами и данными
import pandas as pd
# Импортируем функцию для подключения к базе данных Oracle
from db.connect_db_oracle import get_oracle_connection

# Определяем функцию для выполнения SQL-запроса и получения результатов в виде DataFrame
def query(sql: str, params: dict = None) -> pd.DataFrame:
    # Получаем соединение с базой данных Oracle
    conn = get_oracle_connection()
    # Создаём курсор для выполнения SQL-запросов
    cursor = conn.cursor()
    try:
        # Выполняем SQL-запрос с переданными параметрами (или без параметров, если params = None)
        cursor.execute(sql, params or {})
        # Извлекаем имена столбцов из cursor.description
        columns = [col[0] for col in cursor.description]
        # Получаем все строки результата запроса
        rows = cursor.fetchall()
        # Преобразуем результат в DataFrame pandas, с правильными именами столбцов
        return pd.DataFrame(rows, columns=columns)
    finally:
        # Закрываем курсор в любом случае (даже если произошла ошибка)
        cursor.close()
        # Закрываем соединение с базой данных
        conn.close()