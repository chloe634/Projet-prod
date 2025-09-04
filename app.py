import streamlit as st
from common.design import apply_theme, section
from common.data import get_paths, read_table

apply_theme("Ferment Station — Accueil", "🥤")
section("Accueil", "🏠")
st.caption("Cette app lit **uniquement** les fichiers du repo (`/data`, `/assets`). Aucune importation locale.")

main_table, flavor_map, images_dir = get_paths()
st.write("**Fichier principal :**", main_table)
st.write("**Flavor map :**", flavor_map)
st.write("**Dossier images :**", images_dir)

df_raw = read_table()
if df_raw is None or df_raw.empty:
    st.error("Aucune donnée trouvée. Ajoute ton Excel dans `data/production.xlsx`.")
else:
    st.success("Données détectées ✅ — utilise le menu Pages (à gauche).")
