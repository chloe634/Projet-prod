# pages/01_Production.py
import os
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
from common.storage import (
    list_saved, save_snapshot, load_snapshot, delete_snapshot, MAX_SLOTS
)

# ====== Réglages modèle Excel ======
# Mapping entre le choix UI et le fichier modèle à utiliser
TEMPLATE_MAP = {
    "Cuve de 7000L": "assets/Grande.xlsx",   # anciennement "Fiche de Prod 250620.xlsx"
    "Cuve de 5000L": "assets/Petite.xlsx",
}
SHEET_NAME = None  # laisse None si le modèle a une feuille active par défaut

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
try:
    df_in = apply_canonical_flavor(df_in_raw, fm)
except KeyError as e:
    st.error(f"{e}")
    st.info("Astuce : vérifie la 1ère ligne (en-têtes) de ton Excel et renomme la colonne du nom produit en **'Produit'** ou **'Désignation'**.")
    st.stop()

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

    # 🔥 NOUVEAU : forcer certains goûts
    forced_gouts = st.multiselect(
        "✅ Forcer la production de ces goûts",
        options=[g for g in all_gouts if g not in set(excluded_gouts)],
        help="Les goûts sélectionnés ici seront produits quoi qu’il arrive. "
             "Si tu en choisis plus que le nombre de goûts sélectionnés ci-dessus, "
             "le nombre sera automatiquement augmenté."
    )


st.caption(
    f"Fichier courant : **{st.session_state.get('file_name','(sans nom)')}** — Fenêtre (B2) : **{window_days} jours**"
)

# ---------------- Calculs ----------------
# Nombre de goûts effectif : on garantit que tous les 'forcés' rentrent
effective_nb_gouts = max(nb_gouts, len(forced_gouts)) if forced_gouts else nb_gouts

(
    df_min,
    cap_resume,
    gouts_cibles,
    synth_sel,
    df_calc,
    df_all,
    note_msg,    # 👈 7e valeur renvoyée par compute_plan
) = compute_plan(
    df_in=df_in,
    window_days=window_days,
    volume_cible=volume_cible,
    nb_gouts=effective_nb_gouts,         # 👈 prend en compte les 'forcés'
    repartir_pro_rv=repartir_pro_rv,
    manual_keep=forced_gouts or None,    # 👈 forçage
    exclude_list=excluded_gouts,
)

# Affiche la note d’ajustement si présente (ex: contrainte Infusion/Kéfir)
if isinstance(note_msg, str) and note_msg.strip():
    st.info(note_msg)

# ---------------- Exclusions format (famille + format / goût + format) ----------------
import re

def _parse_family(product: str) -> str:
    # "NIKO - kéfir de fruits Mangue Passion" -> "NIKO"
    if " - " in product:
        return product.split(" - ", 1)[0].strip()
    # sinon, on prend le premier ou deux premiers tokens utiles (ex "Water kefir")
    tokens = str(product).strip().split()
    if len(tokens) >= 2 and " ".join(tokens[:2]).lower() in {"water kefir"}:
        return " ".join(tokens[:2]).title()
    return tokens[0].title() if tokens else ""

def _parse_pack_and_volume(stock: str):
    """
    Exemples:
    - "Carton de 12 Bouteilles - 0.33L"
    - "Carton de 6 Bouteilles 75cl Verralia - 0.75L"
    - "Pack de 4 Bouteilles 75cl SAFT - 0.75L"
    -> retourne (pack_size:int or None, volume_cl:int or None)
    """
    s = str(stock)
    m_pack = re.search(r'\b(\d+)\b', s)
    pack_size = int(m_pack.group(1)) if m_pack else None

    m_vol_l = re.search(r'(\d+(?:[.,]\d+)?)\s*L\b', s, flags=re.IGNORECASE)
    volume_cl = None
    if m_vol_l:
        volume_cl = int(round(float(m_vol_l.group(1).replace(',', '.')) * 100))
    else:
        m_vol_cl = re.search(r'(\d+)\s*cL\b', s, flags=re.IGNORECASE)
        if m_vol_cl:
            volume_cl = int(m_vol_cl.group(1))

    return pack_size, volume_cl

def _add_format_keys(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["family"] = df["Produit"].astype(str).apply(_parse_family)
    parsed = df["Stock"].astype(str).apply(_parse_pack_and_volume)
    df["pack_size"] = parsed.apply(lambda x: x[0])
    df["volume_cl"] = parsed.apply(lambda x: x[1])
    df["family_format_key"] = df.apply(
        lambda r: f'{r["family"]} | {r["pack_size"]}x{r["volume_cl"]}cl', axis=1
    )
    # on utilise GoutCanon (déjà présent dans df_min)
    df["flavor_format_key"] = df.apply(
        lambda r: f'{r["GoutCanon"]} | {r["pack_size"]}x{r["volume_cl"]}cl', axis=1
    )
    return df

# Construire les clés sur le résultat de compute_plan (df_min)
df_min_keys = _add_format_keys(df_min)

st.markdown("### Exclusions supplémentaires (format)")
with st.expander("Exclure des formats de production (sans modifier les calculs)"):
    family_opts = sorted(df_min_keys["family_format_key"].dropna().unique().tolist())
    flavor_opts = sorted(df_min_keys["flavor_format_key"].dropna().unique().tolist())

    st.session_state.setdefault("excl_family_format", [])
    st.session_state.setdefault("excl_flavor_format", [])

    st.session_state["excl_family_format"] = st.multiselect(
        "🚫 Exclure par **famille + format** (ex.: `NIKO | 12x33cl` → exclut tous les goûts NIKO en 12x33cl)",
        options=family_opts,
        default=st.session_state["excl_family_format"],
        key="ms_excl_family_format",
    )

    st.session_state["excl_flavor_format"] = st.multiselect(
        "🚫 Exclure par **goût + format** (ex.: `Mangue Passion | 12x33cl` → exclut seulement ce goût en 12x33cl)",
        options=flavor_opts,
        default=st.session_state["excl_flavor_format"],
        key="ms_excl_flavor_format",
    )

    ccl, ccr = st.columns(2)
    if ccl.button("Vider les exclusions"):
        st.session_state["excl_family_format"] = []
        st.session_state["excl_flavor_format"] = []

# Appliquer le filtre (post-calcul)
_mask_excl = (
    df_min_keys["family_format_key"].isin(st.session_state["excl_family_format"]) |
    df_min_keys["flavor_format_key"].isin(st.session_state["excl_flavor_format"])
)
df_min_final = df_min_keys.loc[~_mask_excl].copy()

# NOTE : à partir d’ici, on utilise df_min_final pour KPIs / affichage / sauvegarde


# ---------------- KPIs ----------------
total_btl = int(pd.to_numeric(df_min_final.get("Bouteilles à produire (arrondi)"), errors="coerce").fillna(0).sum()) if "Bouteilles à produire (arrondi)" in df_min_final.columns else 0
total_vol = float(pd.to_numeric(df_min_final.get("Volume produit arrondi (hL)"), errors="coerce").fillna(0).sum()) if "Volume produit arrondi (hL)" in df_min_final.columns else 0.0
c1, c2, c3 = st.columns(3)
with c1: kpi("Total bouteilles à produire", f"{total_btl:,}".replace(",", " "))
with c2: kpi("Volume total (hL)", f"{total_vol:.2f}")
with c3: kpi("Lignes après exclusions", f"{len(df_min_final)}")


# ---------------- Images + tableau principal ----------------
def sku_guess(name: str):
    m = re.search(r"\b([A-Z]{3,6}-\d{2,3})\b", str(name))
    return m.group(1) if m else None

df_view = df_min_final.copy()
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

_sp_prev = st.session_state.get("saved_production")
default_debut = _dt.date.fromisoformat(_sp_prev["semaine_du"]) if _sp_prev and "semaine_du" in _sp_prev else _dt.date.today()

# Sélecteur de modèle (taille de cuve)
cuve_choice = st.radio(
    "Modèle de fiche",
    options=["Cuve de 7000L", "Cuve de 5000L"],
    horizontal=True,
    help="Choisis le modèle de fiche à générer. Les données (cartons/DDM) viennent de la proposition sauvegardée."
)

# Champ unique : date de début fermentation
date_debut = st.date_input("Date de début de fermentation", value=default_debut)

# DDM = début + 1 an
date_ddm = date_debut + _dt.timedelta(days=365)

if st.button("💾 Sauvegarder cette production", use_container_width=True):
    g_order = []
    if isinstance(df_min_final, pd.DataFrame) and "GoutCanon" in df_min_final.columns:
        for g in df_min_final["GoutCanon"].astype(str).tolist():
            if g and g not in g_order:
                g_order.append(g)

    st.session_state.saved_production = {
        "df_min": df_min_final.copy(),   # <<< ici
        "df_calc": df_calc.copy(),
        "gouts": g_order,
        "semaine_du": date_debut.isoformat(),
        "ddm": date_ddm.isoformat(),
    }
    st.success("Production sauvegardée ✅ — tu peux maintenant générer la fiche.")


sp = st.session_state.get("saved_production")

def _two_gouts_auto(sp_obj, df_min_cur, gouts_cur):
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
    # Déduction auto des 2 premiers goûts (si ta fiche a 2 colonnes de goût)
    g1, g2 = _two_gouts_auto(sp, sp.get("df_min", df_min_final), gouts_cibles)

    template_path = TEMPLATE_MAP.get(cuve_choice)
    if not template_path or not os.path.exists(template_path):
        st.error(
            f"Modèle introuvable pour **{cuve_choice}**. "
            f"Place le fichier **{template_path}** dans le repo."
        )
    else:
        try:
            # 👉 On ré-utilise la même fonction de remplissage : elle accepte un template_path générique
            xlsx_bytes = fill_fiche_7000L_xlsx(
                template_path=template_path,
                semaine_du=_dt.date.fromisoformat(sp["semaine_du"]),
                ddm=_dt.date.fromisoformat(sp["ddm"]),
                gout1=g1 or "",
                gout2=g2,
                df_calc=sp.get("df_calc", df_calc),
                sheet_name=SHEET_NAME,
                df_min=sp.get("df_min", df_min),
            )

            semaine_label = _dt.date.fromisoformat(sp["semaine_du"]).strftime("%d-%m-%Y")
            fname_xlsx = f"Fiche de production — {cuve_choice} — {semaine_label}.xlsx"

            st.download_button(
                "📄 Télécharger la fiche (XLSX)",
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

# ================== Mémoire longue (persistante, 4 entrées max) ==================
st.subheader("Mémoire longue — propositions enregistrées")
st.caption(f"Tu peux garder jusqu’à **{MAX_SLOTS}** propositions nommées, persistantes entre sessions.")

coln1, coln2 = st.columns([2,1])
default_name = ""
if "saved_production" in st.session_state:
    # nom par défaut : semaine du JJ-MM-YYYY + 2 premiers goûts
    _sp = st.session_state["saved_production"]
    try:
        sd = _dt.date.fromisoformat(_sp["semaine_du"]).strftime("%d-%m-%Y")
        g1 = (_sp.get("gouts") or [""])[0] if _sp.get("gouts") else ""
        g2 = (_sp.get("gouts") or ["",""])[1] if _sp.get("gouts") else ""
        default_name = f"{sd} — {g1}{(' + ' + g2) if g2 else ''}"
    except Exception:
        default_name = ""

with coln1:
    name_input = st.text_input("Nom de la proposition", value=default_name, placeholder="ex: 21-10-2025 — Gingembre + Mangue")
with coln2:
    if st.button("📌 Enregistrer dans la mémoire", use_container_width=True):
        sp_cur = st.session_state.get("saved_production")
        if not sp_cur:
            st.error("Sauvegarde d’abord la production ci-dessus (bouton 💾).")
        else:
            ok, msg = save_snapshot(name_input, sp_cur)
            (st.success if ok else st.error)(msg)

saved_list = list_saved()
if saved_list:
    labels = [f"{it['name']} — ({it['semaine_du']})" if it.get("semaine_du") else it["name"] for it in saved_list]
    sel = st.selectbox("Sélectionne une proposition enregistrée", options=labels, index=0)
    idx = labels.index(sel)
    picked = saved_list[idx]["name"]

    # -------- Aperçu de la proposition sélectionnée (df_min sauvegardé) --------
    sp_preview = load_snapshot(picked)
    if sp_preview and isinstance(sp_preview.get("df_min"), pd.DataFrame) and not sp_preview["df_min"].empty:
        with st.expander("👀 Aperçu de la proposition sélectionnée", expanded=False):
            prev_df = sp_preview["df_min"].copy()

            # Petits KPIs (comme pour le tableau courant)
            prev_total_btl = int(pd.to_numeric(prev_df.get("Bouteilles à produire (arrondi)"), errors="coerce").fillna(0).sum()) if "Bouteilles à produire (arrondi)" in prev_df.columns else 0
            prev_total_vol = float(pd.to_numeric(prev_df.get("Volume produit arrondi (hL)"), errors="coerce").fillna(0).sum()) if "Volume produit arrondi (hL)" in prev_df.columns else 0.0
            pk1, pk2, pk3 = st.columns(3)
            with pk1: kpi("Total bouteilles (sauvegardé)", f"{prev_total_btl:,}".replace(",", " "))
            with pk2: kpi("Volume total (hL, sauvegardé)", f"{prev_total_vol:.2f}")
            with pk3: kpi("Lignes", f"{len(prev_df)}")

            # Image facultative comme dans le tableau principal
            prev_df["_SKU?"] = prev_df["Produit"].apply(sku_guess)
            prev_df["__img_path"] = [
                find_image_path(images_dir, sku=sku_guess(p), flavor=g)
                for p, g in zip(prev_df["Produit"], prev_df.get("GoutCanon", pd.Series(dtype=str)))
            ]
            prev_df["Image"] = prev_df["__img_path"].apply(load_image_bytes)

            st.data_editor(
                prev_df[[
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
    else:
        st.info("Aperçu indisponible pour cette proposition (df_min manquant ou vide).")

    # -------- Actions --------
    col_load, col_del, col_count = st.columns(3)
    with col_load:
        if st.button("▶️ Charger", use_container_width=True):
            sp_loaded = load_snapshot(picked)
            if sp_loaded:
                st.session_state["saved_production"] = sp_loaded
                st.success(f"Chargé : {picked}")

    with col_del:
        if st.button("🗑️ Supprimer", use_container_width=True):
            if delete_snapshot(picked):
                st.success("Supprimé.")
            else:
                st.error("Échec suppression.")

    with col_count:
        st.metric("Propositions stockées", f"{len(saved_list)}/{MAX_SLOTS}")
else:
    st.info("Aucune proposition enregistrée pour l’instant.")
