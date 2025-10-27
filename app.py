from __future__ import annotations
import streamlit as st
from common.session import is_authenticated

st.set_page_config(page_title="Accueil", page_icon="🏠", initial_sidebar_state="collapsed")

# Si l'utilisateur n'est pas connecté → on l'envoie sur la page d'auth
if not is_authenticated():
    st.switch_page("pages/00_Auth.py")

# Si connecté → on redirige vers la page principale de travail
st.switch_page("pages/01_Accueil.py")
