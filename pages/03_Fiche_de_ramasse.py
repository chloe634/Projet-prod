import streamlit as st
from common.design import apply_theme, section

apply_theme("Fiche de ramasse — Ferment Station", "🚚")
section("Fiche de ramasse", "🚚")
st.info("Espace réservé — indique-moi les colonnes/tri (tournée, client, SKU, qté, conditionnement) et je branche l’export PDF.")

