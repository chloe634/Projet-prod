from __future__ import annotations
import streamlit as st
from sqlalchemy import text
from db.conn import run_sql

st.set_page_config(page_title="DB migrate", page_icon="🛠️", layout="centered")
st.title("🛠️ Migration DB — password_resets")

SQL = """
CREATE EXTENSION IF NOT EXISTS pgcrypto;

CREATE TABLE IF NOT EXISTS password_resets (
  id          BIGSERIAL PRIMARY KEY,
  user_id     UUID NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  token_hash  TEXT NOT NULL,
  expires_at  TIMESTAMPTZ NOT NULL,
  used_at     TIMESTAMPTZ,
  request_ip  TEXT,
  request_ua  TEXT,
  created_at  TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_password_resets_user  ON password_resets(user_id);
CREATE INDEX IF NOT EXISTS idx_password_resets_token ON password_resets(token_hash);
"""

if st.button("Exécuter la migration", type="primary"):
    try:
        for stmt in [s.strip() for s in SQL.split(";") if s.strip()]:
            run_sql(text(stmt))
        st.success("Migration appliquée ✅")
    except Exception as e:
        st.error(f"Erreur migration : {e}")
