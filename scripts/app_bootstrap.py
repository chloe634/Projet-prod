from sqlalchemy import text
from common.db import engine

def run_file(path: str):
    with open(path, "r", encoding="utf-8") as f:
        sql = f.read()
    with engine().begin() as conn:
        conn.execute(text(sql))
    print("✅ Migration exécutée")

if __name__ == "__main__":
    run_file("db/migrate.sql")
