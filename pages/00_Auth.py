# app/00_Auth.py
import streamlit as st
from common.auth import authenticate, create_user, find_user_by_email
from common.session import login_user, current_user

st.set_page_config(page_title="Authentification", page_icon="🔐", initial_sidebar_state="expanded")

st.title("🔐 Authentification")

# Si déjà connecté, on informe et propose des liens
u = current_user()
if u:
    st.success(f"Déjà connecté en tant que {u['email']}.")
    st.page_link("app/pages/01_Production.py", label="➡️ Aller à la production")
    st.stop()

tab_login, tab_signup = st.tabs(["Se connecter", "Créer un compte"])

with tab_login:
    st.subheader("Connexion")
    email = st.text_input("Email", placeholder="prenom.nom@exemple.com", key="login_email")
    password = st.text_input("Mot de passe", type="password", key="login_pwd")
    c1, c2 = st.columns([1,2])
    with c1:
        if st.button("Connexion", type="primary"):
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
    with c2:
        st.caption("Mot de passe oublié ? (à implémenter plus tard)")

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
                # connexion auto après inscription
                u.pop("password_hash", None)
                login_user(u)
                st.success("Compte créé et connecté ✅")
                st.rerun()
            except Exception as e:
                st.exception(e)
