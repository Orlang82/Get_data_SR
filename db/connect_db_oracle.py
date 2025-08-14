# Авторизация в SR_bank

import json
import os
import oracledb

def get_oracle_connection():
    # 1) Формируем полный путь к файлу в .conda для текущего пользователя
    creds_path = os.path.expanduser(r"~\.conda\db_ac.json")

    # 2) Загружаем параметры
    with open(creds_path, "r", encoding="utf-8") as f:
        creds = json.load(f)

    # 3) Подключаемся в Thin Mode
    return oracledb.connect(
        user=creds["user"],
        password=creds["password"],
        dsn=creds["dsn"]
    )

if __name__ == "__main__":
    conn = get_oracle_connection()
    print("✅ Connected using JSON config in .conda")
    conn.close()
