import streamlit as st
from common.design import apply_theme, section

apply_theme("Fiche de ramasse — Ferment Station", "🚚")
section("Fiche de ramasse", "🚚")

if "df_raw" not in st.session_state:
    st.warning("Aucun fichier chargé. Va dans **Accueil** pour déposer l'Excel, puis reviens.")
    st.stop()

st.caption(f"Fichier courant : **{st.session_state.get('file_name','(sans nom)')}** — Fenêtre (B2) : **{st.session_state.get('window_days','—')} jours**")
st.info("Espace réservé — dis-moi les colonnes/tri (tournée, client, SKU, qté...) et je branche l’export.")
