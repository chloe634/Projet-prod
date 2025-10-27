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

def _hide_sidebar_nav():
    # Masque le menu des pages tant qu'on n'est pas connecté
    st.markdown("""
        <style>
        section[data-testid="stSidebarNav"] {display:none !important;}
        </style>
    """, unsafe_allow_html=True)

def require_login(redirect_to_auth: bool = True) -> Optional[Dict[str, Any]]:
    """
    A appeler tout en haut de CHAQUE page privée.
    Si non connecté : masque la sidebar + redirige vers pages/00_Auth.py puis stoppe la page.
    """
    u = current_user()
    if u:
        return u

    _hide_sidebar_nav()
    st.error("Veuillez vous connecter pour accéder à cette page.")

    if redirect_to_auth:
        # Redirige vers la page d'auth (toujours relative à l'entrypoint app.py)
        try:
            st.switch_page("pages/00_Auth.py")
        except Exception:
            st.page_link("pages/00_Auth.py", label="Aller à l’authentification", icon="🔐")
    st.stop()
    return None  # pour l'éditeur

def require_role(*roles: str) -> Dict[str, Any]:
    u = require_login()
    if u["role"] not in roles:
        st.error("Accès refusé (rôle insuffisant).")
        st.stop()
    return u
    
def _hide_auth_and_entrypoint_links_when_logged_in():
    # Cache le lien vers la page d’auth + l’entrée "app" dans la nav
    st.markdown("""
    <style>
    /* cache lien vers 00_Auth.py dans la nav */
    section[data-testid="stSidebar"] a[href*="00_Auth.py"] { display: none !important; }
    /* cache le lien d'entrée (app.py) si Streamlit l'affiche */
    section[data-testid="stSidebar"] a[href$="app.py"],
    section[data-testid="stSidebar"] a[href*="app.py?"] { display: none !important; }
    </style>
    """, unsafe_allow_html=True)
    
def user_menu():
    """Petit encart utilisateur dans la sidebar (à appeler après require_login())."""
    u = current_user()
    if not u:
        return
    with st.sidebar:
        st.markdown(
            f"**Connecté :** {u['email']}  \n"
            f"**Rôle :** `{u['role']}`  \n"
            f"**Tenant :** `{u['tenant_id']}`"
        )
        if st.button("Se déconnecter", use_container_width=True):
            logout_user()
            st.success("Déconnecté.")
            st.rerun()
    _hide_auth_and_entrypoint_links_when_logged_in()
