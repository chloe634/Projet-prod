import pandas as pd, streamlit as st
from common.design import apply_theme, section, kpi
from common.data import get_paths
from core.optimizer import (
    read_input_excel_and_period_from_path, load_flavor_map_from_path,
    apply_canonical_flavor, compute_losses_table_v48
)

apply_theme("Optimisation & pertes — Ferment Station", "📉")
section("Optimisation & pertes", "📉")

main_table, flavor_map, _ = get_paths()
df_raw, window_days = read_input_excel_and_period_from_path(main_table)
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

