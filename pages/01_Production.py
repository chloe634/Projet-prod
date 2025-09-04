import re, pandas as pd, streamlit as st
from common.design import apply_theme, section, kpi, find_image_path, load_image_bytes
from common.data import get_paths
from core.optimizer import (
    read_input_excel_and_period_from_path,
    read_input_excel_and_period_from_upload,
    load_flavor_map_from_path,
    apply_canonical_flavor, sanitize_gouts,
    compute_plan
)

apply_theme("Production — Ferment Station", "📦")
section("Tableau de production", "📦")

main_table, flavor_map, images_dir = get_paths()

# ---------------- Sidebar: paramètres + source des données ----------------
with st.sidebar:
    st.header("Paramètres")
    volume_cible = st.number_input("Volume cible (hL)", 1.0, 1000.0, 64.0, 1.0)
    nb_gouts = st.selectbox("Nombre de goûts simultanés", [1, 2], index=0)
    repartir_pro_rv = st.checkbox("Répartition au prorata des ventes", value=True)

    st.markdown("---")
    st.subheader("Source des données")
    source = st.radio(
        "Choix",
        ["GitHub (data/production.xlsx)", "Upload manuel"],
        index=0
    )
    uploaded = None
    if source == "Upload manuel":
        uploaded = st.file_uploader("Dépose un Excel (.xlsx / .xls)", type=["xlsx", "xls"])

# ---------------- Lecture des données selon la source ----------------
try:
    if source == "GitHub (data/production.xlsx)":
        df_in_raw, window_days = read_input_excel_and_period_from_path(main_table)
    else:
        if not uploaded:
            st.info("Dépose un fichier pour continuer.")
            st.stop()
        df_in_raw, window_days = read_input_excel_and_period_from_upload(uploaded)
except Exception as e:
    st.error(f"Erreur de lecture des données : {e}")
    st.stop()

st.caption(f"Fenêtre détectée (B2) : **{window_days} jours** (défaut 60 si non détecté).")

# ---------------- Préparation & calcul ----------------
fm = load_flavor_map_from_path(flavor_map)
df_in = apply_canonical_flavor(df_in_raw, fm)
df_in["Produit"] = df_in["Produit"].map(lambda s: s if isinstance(s, str) else str(s))
df_in = sanitize_gouts(df_in)

df_min, cap_resume, gouts_cibles, synth_sel, df_calc, df_all = compute_plan(
    df_in=df_in,
    window_days=window_days,
    volume_cible=volume_cible,
    nb_gouts=nb_gouts,
    repartir_pro_rv=repartir_pro_rv,
    manual_keep=None,
    exclude_list=[]
)

# ---------------- KPIs ----------------
total_btl = int(pd.to_numeric(df_min.get("Bouteilles à produire (arrondi)"), errors="coerce").fillna(0).sum()) if "Bouteilles à produire (arrondi)" in df_min.columns else 0
total_vol = float(pd.to_numeric(df_min.get("Volume produit arrondi (hL)"), errors="coerce").fillna(0).sum()) if "Volume produit arrondi (hL)" in df_min.columns else 0.0
c1, c2, c3 = st.columns(3)
with c1: kpi("Total bouteilles à produire", f"{total_btl:,}".replace(",", " "))
with c2: kpi("Volume total (hL)", f"{total_vol:.2f}")
with c3: kpi("Goûts sélectionnés", f"{len(gouts_cibles)}")

# ---------------- Images ----------------
def sku_guess(name: str):
    m = re.search(r"\b([A-Z]{3,6}-\d{2,3})\b", str(name));  return m.group(1) if m else None

df_view = df_min.copy()
df_view["SKU?"] = df_view["Produit"].apply(sku_guess)
df_view["__img_path"] = [
    find_image_path(images_dir, sku=sku_guess(p), flavor=g)
    for p, g in zip(df_view["Produit"], df_view["GoutCanon"])
]
df_view["Image"] = df_view["__img_path"].apply(load_image_bytes)

# ---------------- Tableau ----------------
st.dataframe(
    df_view[["Image","GoutCanon","Produit","Stock","Cartons à produire (arrondi)","Bouteilles à produire (arrondi)","Volume produit arrondi (hL)"]],
    use_container_width=True,
    hide_index=True,
    column_config={
        "Image": st.column_config.ImageColumn("Image", width="small"),
        "GoutCanon": "Goût",
        "Volume produit arrondi (hL)": st.column_config.NumberColumn(format="%.2f"),
    }
)

with st.expander("Pourquoi ces goûts ?"):
    st.dataframe(synth_sel, use_container_width=True)
