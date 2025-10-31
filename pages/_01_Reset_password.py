# pages/_01_Reset_password.py
from __future__ import annotations
import hashlib
import datetime as dt
import streamlit as st
from sqlalchemy import text
from db.conn import run_sql
from common.auth_reset import consume_token_and_set_password  # on réutilise le setter

st.set_page_config(page_title="Réinitialisation du mot de passe", page_icon="🔐", layout="centered")
st.title("Réinitialisation du mot de passe")

# --- Helpers ---
def _hash_token(t: str) -> str:
    return hashlib.sha256(t.encode()).hexdigest()

def _ensure_table():
    sqls = [
        """
        CREATE TABLE IF NOT EXISTS password_resets (
          id BIGSERIAL PRIMARY KEY,
          user_id UUID NOT NULL REFERENCES users(id) ON DELETE CASCADE,
          token_hash TEXT NOT NULL,
          expires_at TIMESTAMPTZ NOT NULL,
          used_at TIMESTAMPTZ,
          request_ip TEXT,
          request_ua TEXT,
          created_at TIMESTAMPTZ NOT NULL DEFAULT now()
        )
        """,
        "CREATE INDEX IF NOT EXISTS idx_password_resets_user  ON password_resets(user_id)",
        "CREATE INDEX IF NOT EXISTS idx_password_resets_token ON password_resets(token_hash)",
    ]
    for s in sqls:
        run_sql(text(s))

_ensure_table()

# --- Récupération robuste du token depuis l'URL ---
qp = st.query_params  # Streamlit retourne un mapping str -> list[str]
raw = None
if "token" in qp:
    val = qp.get("token")
    if isinstance(val, list):
        raw = (val[0] or "").strip()
    else:
        raw = (val or "").strip()

# Champ de secours si l’URL n’a pas le paramètre
if not raw:
    st.info("Le lien envoyé par e-mail contient un code. Si besoin, copiez-collez le code ci-dessous.")
    raw = st.text_input("Code de réinitialisation", type="password", help="Collez la partie après 'token=' du lien reçu.")
    if not raw:
        st.stop()

token = raw

# --- Diagnostic optionnel : active en ajoutant ?debug=1 à l’URL
debug = ("debug" in qp)
if debug:
    st.caption(f"DEBUG • query_params={dict(qp)}")

# --- Vérification en base (on cherche par token_hash) ---
th = _hash_token(token)
rows = run_sql(text("""
    SELECT id AS reset_id, user_id, expires_at, used_at
    FROM password_resets
    WHERE token_hash = :th
    ORDER BY id DESC
    LIMIT 1
"""), {"th": th})

if not rows:
    st.error("Lien invalide. Le code n'existe pas en base. Refaite une demande depuis « Mot de passe oublié ».")
    if debug:
        st.code(f"token={token}\nsha256={th}", language="text")
    st.stop()

row = rows[0]
expired = dt.datetime.now(dt.timezone.utc) >= row["expires_at"]
used = row["used_at"] is not None

if expired or used:
    if expired:
        st.error("Lien expiré. Refaite une demande depuis « Mot de passe oublié ».")
    else:
        st.error("Ce lien a déjà été utilisé. Refaite une demande depuis « Mot de passe oublié ».")
    if debug:
        st.json(row)
    st.stop()

# --- Formulaire de nouveau mot de passe ---
with st.form("reset_form"):
    pwd1 = st.text_input("Nouveau mot de passe", type="password")
    pwd2 = st.text_input("Confirmer le mot de passe", type="password")
    ok = st.form_submit_button("Mettre à jour mon mot de passe", type="primary")

if ok:
    if len(pwd1) < 8:
        st.warning("Le mot de passe doit faire au moins 8 caractères.")
        st.stop()
    if pwd1 != pwd2:
        st.warning("Les deux mots de passe ne correspondent pas.")
        st.stop()
    try:
        consume_token_and_set_password(row["reset_id"], row["user_id"], pwd1)
        st.success("Mot de passe mis à jour ✅")
        st.page_link("pages/_00_Auth.py", label="➡️ Retour à la connexion")
    except Exception as e:
        st.error(f"Erreur inattendue : {e}")
