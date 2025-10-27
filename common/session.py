from typing import Optional, Dict, Any
import streamlit as st

def sidebar_nav_logged_in():
    """
    Remplace la navigation standard une fois connecté :
    - on cache TOUTE la nav multipage de Streamlit
    - on affiche notre menu propre, sans 'app' ni 'Auth'
    """
    st.markdown("""
    <style>
      /* Cache toute la nav multipage Streamlit */
      [data-testid="stSidebarNav"]              { display: none !important; }
      [data-testid="stSidebarNavItems"]         { display: none !important; }
      /* Si certaines versions insèrent le bloc nav autrement */
      section[data-testid="stSidebar"] nav      { display: none !important; }
      /* Cache tout lien résiduel vers app.py ou Auth (filet de sécurité) */
      section[data-testid="stSidebar"] a[href$="app.py"],
      section[data-testid="stSidebar"] a[href*="/app"],
      section[data-testid="stSidebar"] a[href*="00_Auth.py"],
      section[data-testid="stSidebar"] a[href*="_00_Auth.py"] { display: none !important; }
    </style>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.markdown("### Navigation")
        st.page_link("pages/01_Accueil.py",                 label="Accueil",                 icon="🏠")
        st.page_link("pages/02_Production.py",              label="Production",              icon="📦")
        st.page_link("pages/03_Optimisation.py",            label="Optimisation",            icon="🧮")
        st.page_link("pages/04_Fiche_de_ramasse.py",        label="Fiche de ramasse",        icon="🚚")
        st.page_link("pages/05_Achats_conditionnements.py", label="Achats conditionnements", icon="📦")
        st.page_link("pages/99_Debug.py",                   label="Debug",                   icon="🛠️")


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
    sidebar_nav_logged_in()
    u = current_user()
    if not u:
        return
    with st.sidebar:
        st.markdown(
            f"**Connecté :** {u['email']}  \n"
            f"**Rôle :** `{u['role']}`  \n"
            f"**Tenant :** `{u['tenant_id']}`"
        )
    with st.sidebar:
        st.markdown("---")
        if st.button("Se déconnecter", use_container_width=True):
            logout_user()
            st.success("Déconnecté.")
            st.rerun()
    _hide_auth_and_entrypoint_links_when_logged_in()
