# common/session.py
from __future__ import annotations
from typing import Optional, Dict, Any
import streamlit as st

USER_KEY = "auth_user"

def current_user() -> Optional[Dict[str, Any]]:
    return st.session_state.get(USER_KEY)

def is_authenticated() -> bool:
    return current_user() is not None

def login_user(user_dict: Dict[str, Any]) -> None:
    st.session_state[USER_KEY] = user_dict

def logout_user() -> None:
    if USER_KEY in st.session_state:
        del st.session_state[USER_KEY]

def require_login(redirect_to_auth: bool = True) -> Optional[Dict[str, Any]]:
    u = current_user()
    if u:
        return u
    st.error("Veuillez vous connecter pour accéder à cette page.")
    if redirect_to_auth:
        st.page_link("pages/00_Auth.py", label="Aller à l’authentification", icon="🔐")
        st.stop()
    return None

def require_role(*roles: str) -> Dict[str, Any]:
    u = require_login()
    if u["role"] not in roles:
        st.error("Accès refusé (rôle insuffisant).")
        st.stop()
    return u

def user_menu():
    u = current_user()
    if not u:
        return
    with st.sidebar:
        st.markdown(f"**Connecté :** {u['email']}  \n**Rôle :** `{u['role']}`  \n**Tenant:** `{u['tenant_id']}`")
        if st.button("Se déconnecter", use_container_width=True):
            logout_user()
            st.success("Déconnecté.")
            st.rerun()
