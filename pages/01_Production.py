import re, pandas as pd, streamlit as st
from common.design import apply_theme, section, kpi, find_image_path, load_image_bytes
from common.data import get_paths
from core.optimizer import (
    load_flavor_map_from_path,
    apply_canonical_flavor, sanitize_gouts,
    compute_plan
)
import datetime as _dt
from dateutil.relativedelta import relativedelta
from common.xlsx_fill import fill_fiche_7000L_xlsx


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

# --- Génération Fiche de production (modèle Excel) ---
st.markdown("---")
st.subheader("Fiche de production (modèle Excel)")

cA, cB = st.columns(2)
with cA:
    date_semaine = st.date_input("Semaine du", value=_dt.date.today())
with cB:
    # DDM par défaut = dans 1 an (tu peux changer)
    ddm_date = st.date_input("DDM", value=_dt.date.today() + relativedelta(years=1))

# 1) Bouton qui fige la proposition courante
if st.button("💾 Sauvegarder cette production", use_container_width=True):
    st.session_state["saved_production"] = {
        "timestamp": _dt.datetime.now().isoformat(timespec="seconds"),
        "semaine_du": date_semaine.isoformat(),
        "ddm": ddm_date.isoformat(),
        "df_calc": df_calc.copy(),     # détail complet pour les formats/volumes
        "df_min": df_min.copy(),       # tableau affiché
        "gouts": list(gouts_cibles),   # ✅ IMPORTANT : évite le KeyError
    }
    st.success("Production sauvegardée.")

# (optionnel) bouton pour purger une ancienne sauvegarde incompatible
col_reset, _ = st.columns([1,3])
with col_reset:
    if st.button("♻️ Réinitialiser la fiche sauvegardée"):
        st.session_state.pop("saved_production", None)
        st.success("Fiche sauvegardée réinitialisée.")

# 2) Téléchargement du modèle Excel rempli
sp = st.session_state.get("saved_production")

def _two_gouts_auto(sp_obj, df_min_cur, gouts_cur):
    """Retourne [g1, g2] (2 goûts max) en suivant l'ordre du tableau sauvegardé."""
    # 1) si la sauvegarde contient l'ordre (clé 'gouts')
    if isinstance(sp_obj, dict):
        g_saved = sp_obj.get("gouts")
        if g_saved:
            uniq = []
            for g in g_saved:
                if g and g not in uniq:
                    uniq.append(g)
            if uniq:
                return (uniq + [None, None])[:2]

    # 2) sinon, ordre d'apparition dans df_min
    if isinstance(df_min_cur, pd.DataFrame) and "GoutCanon" in df_min_cur.columns:
        seen = []
        for g in df_min_cur["GoutCanon"].astype(str).tolist():
            if g and g not in seen:
                seen.append(g)
        if seen:
            return (seen + [None, None])[:2]

    # 3) sinon, retomber sur gouts_cibles du calcul courant
    base = list(gouts_cur) if gouts_cur else []
    return (base + [None, None])[:2]

if sp:
    # Goûts à injecter (pas de sélection manuelle)
    g1, g2 = _two_gouts_auto(sp, sp.get("df_min", df_min), gouts_cibles)

    # Chemin du modèle (ta logique actuelle — adapte si tu es passé par config.yaml)
    TEMPLATE_PATH = TEMPLATE_PATH if 'TEMPLATE_PATH' in locals() else "assets/Fiche de Prod 250620.xlsx"

    try:
xlsx_bytes = fill_fiche_7000L_xlsx(
    template_path=TEMPLATE_PATH,
    semaine_du=_dt.date.fromisoformat(sp["semaine_du"]),
    ddm=_dt.date.fromisoformat(sp["ddm"]),
    gout1=g1 or "",
    gout2=g2,   # peut être None → la page droite sera remplie à 0
    df_calc=sp.get("df_calc", df_calc),
    # sheet_name=SHEET_NAME,  # décommente si tu utilises l’option config
    df_min=sp.get("df_min", df_min),   # 👈 AJOUT IMPORTANT
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
