# common/db.py
import os
from typing import Any, Mapping, Optional
from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine, Result

# --- DSN depuis les variables d'environnement Kinsta ---
# Requis : DB_HOST, DB_PORT, DB_NAME, DB_USER, DB_PASSWORD
# Optionnel : DB_SSLMODE (sur Kinsta internal network: "disable")
def _dsn_from_env() -> str:
    host = os.getenv("DB_HOST", "localhost")
    port = os.getenv("DB_PORT", "5432")
    name = os.getenv("DB_NAME", "")
    user = os.getenv("DB_USER", "")
    pwd  = os.getenv("DB_PASSWORD", "")
    ssl  = os.getenv("DB_SSLMODE", "disable")  # réseau interne Kinsta -> disable
    return f"postgresql+psycopg://{user}:{pwd}@{host}:{port}/{name}?sslmode={ssl}"

_ENGINE: Optional[Engine] = None

def get_engine() -> Engine:
    global _ENGINE
    if _ENGINE is None:
        _ENGINE = create_engine(
            _dsn_from_env(),
            pool_pre_ping=True,
            future=True,
        )
    return _ENGINE

def run_sql(sql: Any, params: Optional[Mapping[str, Any]] = None) -> Result:
    """
    Exécute une requête SQL (str ou sqlalchemy.text) et renvoie le Result.
    Usage :
        rows = run_sql("SELECT 1 AS x")
        for r in rows: ...
    """
    from sqlalchemy import text as _text
    if isinstance(sql, str):
        sql = _text(sql)
    with get_engine().begin() as conn:
        return conn.execute(sql, params or {})

def masked_dsn() -> str:
    """Pratique pour debug : affiche hôte et sslmode sans secrets."""
    host = os.getenv("DB_HOST", "?")
    ssl  = os.getenv("DB_SSLMODE", "disable")
    return f"host={host} | sslmode={ssl}"
