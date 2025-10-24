# db/__init__.py
from .conn import get_engine as engine, run_sql, ping  # 'engine' = alias vers get_engine()

