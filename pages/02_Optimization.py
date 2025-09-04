import pandas as pd, streamlit as st
from common.design import apply_theme, section, kpi
from common.data import get_paths
from core.optimizer import (
    load_flavor_map_from_path,
    apply_canonical_flavor, compute_losses_table_v48
)

apply_theme("Optimisation & pertes — Ferment Station", "📉")
section("Optimisation & pertes", "📉")

# besoin du fichier en mémoire
if "df_raw" not in st.session_state or "window_days" not in st.session_state:
    st.warning("Aucun fichier chargé. Va dans **Accueil** pour déposer l'Excel, puis reviens.")
    st.stop()

_, flavor_map, _ = get_paths()

df_raw = st.session_state.df_raw
window_days = st.session_state.window_days
st.caption(f"Fichier courant : **{st.session_state.get('file_name','(sans nom)')}** — Fenêtre (B2) : **{window_days} jours**")

fm = load_flavor_map_from_path(flavor_map)
df_in = apply_canonical_flavor(df_raw, fm)

price_hL = 500.0
pertes = compute_losses_table_v48(df_in, window_days, price_hL)

colA, colB = st.columns([2,1])
with colA:
    if pertes is not None and not pertes.empty:
        st.dataframe(pertes, use_container_width=True, hide_index=True)
    else:
        st.info("Aucune perte estimée sur 7 jours (données insuffisantes ou stock suffisant).")
with colB:
    total = float(pertes["Perte (€)"].sum()) if isinstance(pertes, pd.DataFrame) and not pertes.empty else 0.0
    kpi("Perte totale (7 j)", f"€{total:,.0f}")
