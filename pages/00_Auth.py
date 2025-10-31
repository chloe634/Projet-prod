from __future__ import annotations
import os
import streamlit as st

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

# Reset password (création de lien + envoi e-mail)
from common.auth_reset import create_password_reset
from common.email import send_reset_email


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
        st.info("Retournez dans l’onglet **Se connecter** pour vous authentifier après le changement de mot de passe.")
        if st.button("Envoyer un autre lien"):
            st.session_state["reset_sent"] = False
            st.rerun()
        return

    if st.button("Envoyer le lien de réinitialisation", type="primary"):
        meta = {"ip": st.session_state.get("client_ip"), "ua": st.session_state.get("client_ua")}
        try:
            # Doit renvoyer une URL du type: {BASE_URL}/06_Reset_password?token=XXXX
            reset_url = create_password_reset(email, meta=meta)
            if reset_url:  # on n'envoie que si on a un vrai lien
                send_reset_email(email, reset_url)
            # sinon: on fait comme si OK pour ne pas divulguer l’existence du compte
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
