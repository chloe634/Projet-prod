# pages/06_Reset_password.py
from __future__ import annotations
import hashlib
import datetime as dt
import streamlit as st
from sqlalchemy import text
from db.conn import run_sql
from common.auth_reset import consume_token_and_set_password

st.set_page_config(page_title="Réinitialisation du mot de passe", page_icon="🔐", layout="centered")
st.title("Réinitialisation du mot de passe")

def _sha256(s: str) -> str:
    return hashlib.sha256(s.encode()).hexdigest()

# Récupération robuste du token (depuis l'URL ou champ manuel)
qp = st.query_params
raw = qp.get("token")
if isinstance(raw, list):
    raw = raw[0]
token = (raw or "").strip()

if not token:
    st.info("Le lien reçu par e-mail contient le code automatiquement. Si besoin, collez-le ci-dessous.")
    token = st.text_input("Code de réinitialisation", type="password", placeholder="coller le token ici")
    if not token:
        st.stop()

# Lookup en base
th = _sha256(token)
rows = run_sql(text("""
    SELECT id AS reset_id, user_id, expires_at, used_at
    FROM password_resets
    WHERE token_hash = :th
    ORDER BY id DESC
    LIMIT 1
"""), {"th": th})

if not rows:
    st.error("Lien invalide. Refaite une demande depuis « Mot de passe oublié ».")
    st.stop()

row = rows[0]
now = dt.datetime.now(dt.timezone.utc)
if row["used_at"] is not None:
    st.error("Ce lien a déjà été utilisé. Refaite une demande depuis « Mot de passe oublié ».")
    st.stop()
if now >= row["expires_at"]:
    st.error("Lien expiré. Refaite une demande depuis « Mot de passe oublié ».")
    st.stop()

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
