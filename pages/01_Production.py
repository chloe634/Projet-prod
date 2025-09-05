import re, pandas as pd, streamlit as st
from common.design import apply_theme, section, kpi, find_image_path, load_image_bytes
from common.data import get_paths
from core.optimizer import (
    load_flavor_map_from_path,
    apply_canonical_flavor, sanitize_gouts,
    compute_plan
)

apply_theme("Production — Ferment Station", "📦")
section("Tableau de production", "📦")

# ex : on garde flavor_map et images_dir du repo
_, flavor_map, images_dir = get_paths()

# --- vérifie que le fichier a été déposé sur Accueil ---
if "df_raw" not in st.session_state or "window_days" not in st.session_state:
    st.warning("Aucun fichier chargé. Va dans **Accueil** pour déposer l'Excel, puis reviens.")
    st.stop()

df_in_raw = st.session_state.df_raw
window_days = st.session_state.window_days

# mapping + nettoyage
fm = load_flavor_map_from_path(flavor_map)
df_in = apply_canonical_flavor(df_in_raw, fm)
df_in["Produit"] = df_in["Produit"].astype(str)
df_in = sanitize_gouts(df_in)

# ---------- SIDEBAR : paramètres + exclusions ----------
with st.sidebar:
    st.header("Paramètres")
    volume_cible = st.number_input("Volume cible (hL)", 1.0, 1000.0, 64.0, 1.0)
    nb_gouts = st.selectbox("Nombre de goûts simultanés", [1, 2], index=0)
    repartir_pro_rv = st.checkbox("Répartition au prorata des ventes", value=True)

    st.markdown("---")
    st.subheader("Filtres")
    all_gouts = sorted(pd.Series(df_in.get("GoutCanon", pd.Series(dtype=str))).dropna().astype(str).str.strip().unique())
    excluded_gouts = st.multiselect(
        "🚫 Exclure certains goûts",
        options=all_gouts,
        default=[]
    )

st.caption(f"Fichier courant : **{st.session_state.get('file_name','(sans nom)')}** — Fenêtre (B2) : **{window_days} jours**")

# ---------- CALCULS ----------
df_min, cap_resume, gouts_cibles, synth_sel, df_calc, df_all = compute_plan(
    df_in=df_in,
    window_days=window_days,
    volume_cible=volume_cible,
    nb_gouts=nb_gouts,
    repartir_pro_rv=repartir_pro_rv,
    manual_keep=None,
    exclude_list=excluded_gouts,   # 👈 prise en compte des exclusions
)

# ---------- KPIs ----------
total_btl = int(pd.to_numeric(df_min.get("Bouteilles à produire (arrondi)"), errors="coerce").fillna(0).sum()) if "Bouteilles à produire (arrondi)" in df_min.columns else 0
total_vol = float(pd.to_numeric(df_min.get("Volume produit arrondi (hL)"), errors="coerce").fillna(0).sum()) if "Volume produit arrondi (hL)" in df_min.columns else 0.0
c1, c2, c3 = st.columns(3)
with c1: kpi("Total bouteilles à produire", f"{total_btl:,}".replace(",", " "))
with c2: kpi("Volume total (hL)", f"{total_vol:.2f}")
with c3: kpi("Goûts sélectionnés", f"{len(gouts_cibles)}")

# ---------- Images ----------
def sku_guess(name: str):
    m = re.search(r"\b([A-Z]{3,6}-\d{2,3})\b", str(name));  return m.group(1) if m else None

df_view = df_min.copy()
df_view["SKU?"] = df_view["Produit"].apply(sku_guess)
df_view["__img_path"] = [
    find_image_path(images_dir, sku=sku_guess(p), flavor=g)
    for p, g in zip(df_view["Produit"], df_view["GoutCanon"])
]
df_view["Image"] = df_view["__img_path"].apply(load_image_bytes)




# ---------- Tableau ----------
st.data_editor(
    df_view[[
        "Image","GoutCanon","Produit","Stock",
        "Cartons à produire (arrondi)","Bouteilles à produire (arrondi)",
        "Volume produit arrondi (hL)"
    ]],
    use_container_width=True,
    hide_index=True,
    disabled=True,
    column_config={
        "Image": st.column_config.ImageColumn("Image", width="small"),
        "GoutCanon": "Goût",
        "Volume produit arrondi (hL)": st.column_config.NumberColumn(format="%.2f"),
    },
)

# --- En-dessous du tableau : saisies & génération PDF ---
import datetime as _dt
from dateutil.relativedelta import relativedelta
from common.pdf import generate_production_pdf

st.markdown("---")
st.subheader("Fiche de production")

colD1, colD2 = st.columns(2)
with colD1:
    date_semaine = st.date_input("Semaine du", value=_dt.date.today())
with colD2:
    ddm_date = st.date_input("DDM", value=_dt.date.today() + relativedelta(months=6))

# Sauvegarde logique (on mémorise la production affichée)
if st.button("💾 Sauvegarder cette production", use_container_width=True):
    st.session_state["saved_production"] = {
        "timestamp": _dt.datetime.now().isoformat(timespec="seconds"),
        "semaine_du": str(date_semaine),
        "ddm": str(ddm_date),
        "df_calc": df_calc.copy(),   # le détail complet (pour le PDF)
        "df_min": df_min.copy(),     # le tableau utilisateur
    }
    st.success("Production sauvegardée pour génération de la fiche.")

# Si on a une prod sauvegardée, on propose la génération PDF
sp = st.session_state.get("saved_production")
if sp:
    # Déduire Produit 1 & Produit 2 depuis df_min sauvegardé (ordre affiché)
    _df_min = sp["df_min"]
    produits_list = _df_min["Produit"].astype(str).tolist() if "Produit" in _df_min.columns else []
    produit_1 = produits_list[0] if len(produits_list) >= 1 else ""
    produit_2 = produits_list[1] if len(produits_list) >= 2 else None

    # Construit le PDF
    pdf_bytes = generate_production_pdf(
        semaine_du=_dt.datetime.fromisoformat(sp["semaine_du"]).date(),
        ddm=_dt.datetime.fromisoformat(sp["ddm"]).date(),
        produit_1=produit_1,
        produit_2=produit_2,
        df_calc=sp["df_calc"],
        entreprise="Ferment Station",
        titre_modele="Fiche de production 7000L",
    )

    # Nom de fichier — les '/' ne sont pas valides dans un nom de fichier
    semaine_label = _dt.datetime.fromisoformat(sp["semaine_du"]).date().strftime("%d-%m-%Y")
    filename = f"Fiche de production (semaine du {semaine_label}).pdf"

    st.download_button(
        "📄 Télécharger la fiche PDF",
        data=pdf_bytes,
        file_name=filename,
        mime="application/pdf",
        use_container_width=True,
    )


