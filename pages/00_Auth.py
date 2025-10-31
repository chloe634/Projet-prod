from __future__ import annotations
import streamlit as st
import os

# --- Config page + masquage sidebar ---
st.set_page_config(page_title="Authentification", page_icon="🔐", initial_sidebar_state="collapsed")
st.markdown("""
<style>
section[data-testid="stSidebar"] {display:none !important;}
section[data-testid="stSidebarNav"] {display:none !important;}
</style>
""", unsafe_allow_html=True)

# --- Imports app ---
from common.auth import authenticate, create_user, find_user_by_email
from common.session import login_user, current_user

# Reset password (Brevo + token)
from common.auth_reset import create_password_reset
from common.email import send_reset_email


# --- Reset inline quand on reçoit ?reset_token=... ---
import hashlib
import datetime as dt
from sqlalchemy import text
from db.conn import run_sql
from common.auth_reset import consume_token_and_set_password

def _hash_token_for_lookup(t: str) -> str:
    return hashlib.sha256(t.encode()).hexdigest()

# Si l'URL contient reset_token, on affiche directement le formulaire de reset ici
_qp = st.query_params
_reset_token = None
if "reset_token" in _qp:
    val = _qp.get("reset_token")
    _reset_token = (val[0] if isinstance(val, list) else val) or ""
    _reset_token = _reset_token.strip()

if _reset_token:
    st.set_page_config(page_title="Réinitialisation du mot de passe", page_icon="🔐", layout="centered")
    st.title("Réinitialisation du mot de passe")

    th = _hash_token_for_lookup(_reset_token)
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
    if row["used_at"] is not None:
        st.error("Ce lien a déjà été utilisé. Refaite une demande depuis « Mot de passe oublié ».")
        st.stop()
    if dt.datetime.now(dt.timezone.utc) >= row["expires_at"]:
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
    st.stop()


# --- Titre ---
st.title("🔐 Authentification")

# --- Si déjà connecté, on redirige vers l'app ---
u = current_user()
if u:
    st.success(f"Déjà connecté en tant que {u['email']}.")
    st.page_link("pages/01_Accueil.py", label="➡️ Aller à la production")
    st.stop()

# ===============================
# UI: Mot de passe oublié (onglet 3)
# ===============================
def forgot_password_ui():
    st.subheader("Mot de passe oublié")
    email = st.text_input("Votre e-mail", placeholder="prenom.nom@exemple.com", key="forgot_email")
    sent = st.session_state.get("reset_sent", False)

    if sent:
        st.success("Si un compte existe pour cet e-mail, un message a été envoyé avec un lien de réinitialisation.")
        # On propose simplement de revenir aux onglets de connexion
        st.info("Retournez dans l’onglet **Se connecter** pour vous authentifier après le changement de mot de passe.")
        if st.button("Envoyer un autre lien"):
            st.session_state["reset_sent"] = False
            st.rerun()
        return

    if st.button("Envoyer le lien de réinitialisation", type="primary"):
        meta = {"ip": st.session_state.get("client_ip"), "ua": st.session_state.get("client_ua")}
        try:
            reset_url = create_password_reset(email, meta=meta)
            result = send_reset_email(email, reset_url or (os.getenv("BASE_URL", "") + "/_01_Reset_password?token=INVALID"))
            # Optionnel: petit log de succès
            st.toast("Email envoyé ✅")
        except Exception as e:
            st.error(f"Erreur d'envoi e-mail : {e}")
            st.stop()
        st.session_state["reset_sent"] = True
        st.rerun()


# ===============================
# Onglets: Connexion / Inscription / Mot de passe oublié
# ===============================
tab_login, tab_signup, tab_forgot = st.tabs(["Se connecter", "Créer un compte", "Mot de passe oublié ?"])

# --- Onglet 1 : Connexion ---
with tab_login:
    st.subheader("Connexion")
    email = st.text_input("Email", placeholder="prenom.nom@exemple.com", key="login_email")
    password = st.text_input("Mot de passe", type="password", key="login_pwd")
    cols = st.columns([1, 1, 2])
    with cols[0]:
        if st.button("Connexion", type="primary", key="btn_login"):
            if not email or not password:
                st.warning("Renseigne email et mot de passe.")
            else:
                user = authenticate(email, password)
                if not user:
                    st.error("Identifiants invalides.")
                else:
                    login_user(user)
                    st.success("Connecté ✅")
                    st.rerun()
    with cols[1]:
        # Petit rappel visuel pour l’onglet 3
        st.caption("💡 Besoin d’aide ? Allez dans l’onglet **Mot de passe oublié ?**")

# --- Onglet 2 : Création de compte ---
with tab_signup:
    st.subheader("Inscription")
    st.caption("Le premier utilisateur d’un tenant devient **admin** automatiquement.")
    new_email = st.text_input("Email", key="su_email")
    new_pwd   = st.text_input("Mot de passe", type="password", key="su_pwd")
    new_pwd2  = st.text_input("Confirme le mot de passe", type="password", key="su_pwd2")
    tenant_name = st.text_input("Nom d’organisation (tenant)", placeholder="Ferment Station", key="su_tenant")

    if st.button("Créer le compte", type="primary", key="btn_signup"):
        if not (new_email and new_pwd and new_pwd2 and tenant_name):
            st.warning("Tous les champs sont obligatoires.")
        elif new_pwd != new_pwd2:
            st.error("Les mots de passe ne correspondent pas.")
        elif find_user_by_email(new_email):
            st.error("Un compte existe déjà avec cet email.")
        else:
            try:
                u = create_user(new_email, new_pwd, tenant_name)
                # Connexion auto après inscription
                u.pop("password_hash", None)
                login_user(u)
                st.success("Compte créé et connecté ✅")
                st.rerun()
            except Exception as e:
                st.exception(e)

# --- Onglet 3 : Mot de passe oublié ---
with tab_forgot:
    forgot_password_ui()
