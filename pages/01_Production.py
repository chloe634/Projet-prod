# pages/01_Production.py
import re
import datetime as _dt
import pandas as pd
import streamlit as st
from dateutil.relativedelta import relativedelta

from common.design import apply_theme, section, kpi, find_image_path, load_image_bytes
from common.data import get_paths
from core.optimizer import (
    load_flavor_map_from_path,
    apply_canonical_flavor, sanitize_gouts,
    compute_plan,
)
from common.xlsx_fill import fill_fiche_7000L_xlsx

# ====== Réglages modèle Excel ======
TEMPLATE_PATH = "assets/Fiche de Prod 5K - 250829.xlsx"   # <- mets ici le nom de ton modèle
SHEET_NAME = None   # ou "Fiche de production 7000L" si tu veux forcer un onglet précis

# ---------------- UI header ----------------
apply_theme("Production — Ferment Station", "📦")
section("Tableau de production", "📦")

# ---------------- Pré-requis : fichier chargé sur Accueil ----------------
if "df_raw" not in st.session_state or "window_days" not in st.session_state:
    st.warning("Aucun fichier chargé. Va dans **Accueil** pour déposer l'Excel, puis reviens.")
    st.stop()

# chemins (repo)
_, flavor_map, images_dir = get_paths()

# Données depuis l'accueil
df_in_raw = st.session_state.df_raw
window_days = st.session_state.window_days

# ---------------- Préparation des données ----------------
fm = load_flavor_map_from_path(flavor_map)
df_in = apply_canonical_flavor(df_in_raw, fm)
df_in["Produit"] = df_in["Produit"].astype(str)
df_in = sanitize_gouts(df_in)

# ---------------- Sidebar (paramètres) ----------------
with st.sidebar:
    st.header("Paramètres")
    volume_cible = st.number_input("Volume cible (hL)", 1.0, 1000.0, 64.0, 1.0)
    nb_gouts = st.selectbox("Nombre de goûts simultanés", [1, 2], index=0)
    repartir_pro_rv = st.checkbox("Répartition au prorata des ventes", value=True)

    st.markdown("---")
    st.subheader("Filtres")
    all_gouts = sorted(pd.Series(df_in.get("GoutCanon", pd.Series(dtype=str))).dropna().astype(str).str.strip().unique())
    excluded_gouts = st.multiselect("🚫 Exclure certains goûts", options=all_gouts, default=[])

st.caption(
    f"Fichier courant : **{st.session_state.get('file_name','(sans nom)')}** — Fenêtre (B2) : **{window_days} jours**"
)

# ---------------- Calculs ----------------
df_min, cap_resume, gouts_cibles, synth_sel, df_calc, df_all = compute_plan(
    df_in=df_in,
    window_days=window_days,
    volume_cible=volume_cible,
    nb_gouts=nb_gouts,
    repartir_pro_rv=repartir_pro_rv,
    manual_keep=None,
    exclude_list=excluded_gouts,
)

# ---------------- KPIs ----------------
total_btl = int(pd.to_numeric(df_min.get("Bouteilles à produire (arrondi)"), errors="coerce").fillna(0).sum()) if "Bouteilles à produire (arrondi)" in df_min.columns else 0
total_vol = float(pd.to_numeric(df_min.get("Volume produit arrondi (hL)"), errors="coerce").fillna(0).sum()) if "Volume produit arrondi (hL)" in df_min.columns else 0.0
c1, c2, c3 = st.columns(3)
with c1: kpi("Total bouteilles à produire", f"{total_btl:,}".replace(",", " "))
with c2: kpi("Volume total (hL)", f"{total_vol:.2f}")
with c3: kpi("Goûts sélectionnés", f"{len(gouts_cibles)}")

# ---------------- Images + tableau principal ----------------
def sku_guess(name: str):
    m = re.search(r"\b([A-Z]{3,6}-\d{2,3})\b", str(name))
    return m.group(1) if m else None

df_view = df_min.copy()
df_view["SKU?"] = df_view["Produit"].apply(sku_guess)
df_view["__img_path"] = [
    find_image_path(images_dir, sku=sku_guess(p), flavor=g)
    for p, g in zip(df_view["Produit"], df_view["GoutCanon"])
]
df_view["Image"] = df_view["__img_path"].apply(load_image_bytes)

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

# ======================================================================
# ========== Sauvegarde + génération de la fiche Excel ==================
# ======================================================================
section("Fiche de production (modèle Excel)", "🧾")

# Valeurs par défaut (si déjà sauvegardé)
_sp_prev = st.session_state.get("saved_production")
default_semaine = _dt.date.fromisoformat(_sp_prev["semaine_du"]) if _sp_prev and "semaine_du" in _sp_prev else _dt.date.today()
default_ddm     = _dt.date.fromisoformat(_sp_prev["ddm"])       if _sp_prev and "ddm" in _sp_prev       else _dt.date.today()

colA, colB = st.columns(2)
with colA:
    date_semaine = st.date_input("Semaine du", value=default_semaine)
with colB:
    date_ddm = st.date_input("DDM (date limite)", value=default_ddm)

# Bouton de sauvegarde (fige les données utilisées pour la fiche)
if st.button("💾 Sauvegarder cette production", use_container_width=True):
    # ordre des goûts = ordre d'apparition dans le tableau affiché
    g_order = []
    if isinstance(df_min, pd.DataFrame) and "GoutCanon" in df_min.columns:
        for g in df_min["GoutCanon"].astype(str).tolist():
            if g and g not in g_order:
                g_order.append(g)

    st.session_state.saved_production = {
        "df_min": df_min.copy(),
        "df_calc": df_calc.copy(),
        "gouts": g_order,
        "semaine_du": date_semaine.isoformat(),
        "ddm": date_ddm.isoformat(),
    }
    st.success("Production sauvegardée ✅ — tu peux maintenant générer la fiche.")

# Si on a une sauvegarde, proposer la génération du XLSX
sp = st.session_state.get("saved_production")

def _two_gouts_auto(sp_obj, df_min_cur, gouts_cur):
    """Retourne [g1, g2] (2 goûts max) en suivant l'ordre du tableau sauvegardé."""
    if isinstance(sp_obj, dict):
        g_saved = sp_obj.get("gouts")
        if g_saved:
            uniq = []
            for g in g_saved:
                if g and g not in uniq:
                    uniq.append(g)
            if uniq:
                return (uniq + [None, None])[:2]
    if isinstance(df_min_cur, pd.DataFrame) and "GoutCanon" in df_min_cur.columns:
        seen = []
        for g in df_min_cur["GoutCanon"].astype(str).tolist():
            if g and g not in seen:
                seen.append(g)
        if seen:
            return (seen + [None, None])[:2]
    base = list(gouts_cur) if gouts_cur else []
    return (base + [None, None])[:2]

if sp:
    g1, g2 = _two_gouts_auto(sp, sp.get("df_min", df_min), gouts_cibles)

    if not st.file_uploader:  # rien à faire, juste éviter les warnings mypy
        pass

    if not st.session_state.get("model_path_checked") and not st.session_state.get("model_path_warning"):
        st.session_state.model_path_checked = True

    # Vérifier la présence du modèle
    if not os.path.exists(TEMPLATE_PATH):
        st.error(f"Modèle introuvable. Place le fichier **{TEMPLATE_PATH}** dans le repo.")
    else:
        try:
            xlsx_bytes = fill_fiche_7000L_xlsx(
                template_path=TEMPLATE_PATH,
                semaine_du=_dt.date.fromisoformat(sp["semaine_du"]),
                ddm=_dt.date.fromisoformat(sp["ddm"]),
                gout1=g1 or "",
                gout2=g2,  # peut être None → la page droite sera remplie à 0
                df_calc=sp.get("df_calc", df_calc),
                sheet_name=SHEET_NAME,
                df_min=sp.get("df_min", df_min),   # <- tableau affiché
            )

            semaine_label = _dt.date.fromisoformat(sp["semaine_du"]).strftime("%d-%m-%Y")
            fname_xlsx = f"Fiche de production (semaine du {semaine_label}).xlsx"

            st.download_button(
                "📄 Télécharger la fiche (XLSX, 2 pages, identique au modèle)",
                data=xlsx_bytes,
                file_name=fname_xlsx,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except FileNotFoundError:
            st.error("Modèle introuvable. Vérifie le chemin du fichier modèle.")
        except Exception as e:
            st.error(f"Erreur lors du remplissage du modèle : {e}")
else:
    st.info("Sauvegarde la production ci-dessus pour activer la génération de la fiche.")
