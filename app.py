import io
import re
import numpy as np
import pandas as pd
import streamlit as st

# =========================
# Optimiseur de production (v3.4)
# - 1 goût : 64 hL PAR goût
# - 2 goûts : 64 hL AU TOTAL répartis globalement
# - Formats APPLICÉS EN INTERNE :
#     * 33 cL  → carton de 12
#     * 75 cL  → carton de 6
#     * 75 cL  → carton de 4
# - Arrondi au carton appliqué en interne (nearest / half-up)
# =========================

st.set_page_config(page_title="Optimiseur de production — 64 hL / 1–2 goûts", page_icon="🧪", layout="wide")

# -------- Réglages cachés --------
ALLOWED_FORMATS = {  # (nb_bouteilles_par_carton, volume_bouteille_L)
    (12, 0.33),
    (6, 0.75),
    (4, 0.75),
}
ROUND_TO_CARTON = True          # arrondi half-up des cartons en interne
VOL_TOL = 0.005                 # tolérance sur volume bouteille (L) pour matcher 0.33 / 0.75

# ---------- Sidebar ----------
with st.sidebar:
    st.header("Paramètres")
    volume_cible = st.number_input(
        "Volume cible (hL)",
        min_value=1.0,
        value=64.0,
        step=1.0,
        help="Si 1 goût: volume PAR goût. Si 2 goûts: volume TOTAL partagé."
    )
    nb_gouts = st.selectbox("Nombre de goûts simultanés", options=[1, 2], index=0)
    repartir_pro_rv = st.checkbox(
        "Répartir par formats au prorata des vitesses de vente",
        value=True,
        help="Si décoché: répartition égale entre formats d'un même goût."
    )

    st.markdown("---")
    st.subheader("Contraintes goûts (optionnel)")
    use_manual = st.checkbox("Sélection manuelle des goûts", value=False)
    gouts_exclus = st.text_input("Exclure goûts (séparés par des virgules)", value="")

# ---------- Header ----------
st.title("🧪 Optimiseur de production — 64 hL / 1–2 goûts")
st.caption("Charge un Excel d'autonomie, choisis tes options, et génère un plan propre pour l'atelier.")

# ---------- Upload ----------
uploaded = st.file_uploader("Dépose ton fichier Excel (.xlsx/.xls)", type=["xlsx", "xls"]) 

# ---------- Utils ----------
def detect_header_row(df_raw: pd.DataFrame) -> int:
    must_have = {"Produit", "Stock", "Quantité vendue", "Volume vendu (hl)", "Quantité disponible", "Volume disponible (hl)"}
    for i in range(min(10, len(df_raw))):
        row_vals = set(str(x).strip() for x in df_raw.iloc[i].tolist())
        if must_have.issubset(row_vals):
            return i
    return 0

def read_input_excel(file) -> pd.DataFrame:
    raw = pd.read_excel(file, header=None)
    header_idx = detect_header_row(raw)
    df = pd.read_excel(file, header=header_idx)
    return df

def parse_stock(text: str):
    """Retourne (nb_bouteilles, volume_bouteille_L) extraits du champ 'Stock'."""
    if pd.isna(text):
        return np.nan, np.nan
    s = str(text)

    # nb bouteilles/carton
    m_nb = re.search(r"Carton de\s*(\d+)", s, flags=re.IGNORECASE)
    nb = int(m_nb.group(1)) if m_nb else np.nan

    # volume bouteille (L)
    m_l = re.findall(r"(\d+(?:[.,]\d+)?)\s*[lL]", s)
    if m_l:
        vol_l = float(m_l[-1].replace(',', '.'))
    else:
        m_cl = re.findall(r"(\d+(?:[.,]\d+)?)\s*c[lL]", s)
        vol_l = float(m_cl[-1].replace(',', '.')) / 100.0 if m_cl else np.nan
    return nb, vol_l

def safe_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce")

def is_allowed_format(nb_bottles, vol_l) -> bool:
    """Vérifie (nb, vol) contre les formats autorisés avec tolérance sur le volume."""
    if pd.isna(nb_bottles) or pd.isna(vol_l):
        return False
    for nb_ok, vol_ok in ALLOWED_FORMATS:
        if int(nb_bottles) == nb_ok and abs(float(vol_l) - vol_ok) <= VOL_TOL:
            return True
    return False

# ---------- Core calc ----------
def compute_plan(df_in, volume_cible, nb_gouts, repartir_pro_rv, manual_keep, exclude_list):
    required = ["Produit", "Stock", "Quantité vendue", "Volume vendu (hl)", "Quantité disponible", "Volume disponible (hl)"]
    missing = [c for c in required if c not in df_in.columns]
    if missing:
        raise ValueError(f"Colonnes manquantes: {missing}")

    df = df_in[required].copy()
    for c in ["Quantité vendue", "Volume vendu (hl)", "Quantité disponible", "Volume disponible (hl)"]:
        df[c] = safe_num(df[c])

    # Parsing Stock
    df[["Bouteilles/carton", "Volume bouteille (L)"]] = df["Stock"].apply(lambda s: pd.Series(parse_stock(s)))
    # Filtre formats : 33cLx12, 75cLx6, 75cLx4
    df = df[df.apply(lambda r: is_allowed_format(r["Bouteilles/carton"], r["Volume bouteille (L)"]), axis=1)].reset_index(drop=True)

    # Volume/carton
    df["Volume/carton (hL)"] = (df["Bouteilles/carton"] * df["Volume bouteille (L)"]) / 100.0

    # Lignes valides
    df = df.dropna(subset=["Produit", "Volume/carton (hL)", "Volume vendu (hl)", "Volume disponible (hl)"]).reset_index(drop=True)

    # Exclusions & sélection manuelle
    if exclude_list:
        excl = [g.strip() for g in exclude_list]
        df = df[~df["Produit"].astype(str).str.strip().isin(excl)]
    if manual_keep:
        keep = [g.strip() for g in manual_keep]
        df = df[df["Produit"].astype(str).str.strip().isin(keep)]

    # Choix auto des goûts (si pas manuel) : top N par ventes hL
    ventes_par_gout = df.groupby("Produit")["Volume vendu (hl)"].sum().sort_values(ascending=False)
    if not manual_keep:
        gouts_cibles = ventes_par_gout.index.tolist()[:nb_gouts]
        df = df[df["Produit"].isin(gouts_cibles)]
    else:
        gouts_cibles = sorted(set(df["Produit"]))
    if len(gouts_cibles) == 0:
        raise ValueError("Aucun goût sélectionné.")

    # ----- Deux modes -----
    if nb_gouts == 1:
        # Mode 1 goût : 64 hL PAR goût (logique V1)
        df["Somme ventes (hL) par goût"] = df.groupby("Produit")["Volume vendu (hl)"].transform("sum")
        if repartir_pro_rv:
            df["r_i"] = np.where(df["Somme ventes (hL) par goût"] > 0,
                                 df["Volume vendu (hl)"] / df["Somme ventes (hL) par goût"], 0.0)
        else:
            df["r_i"] = 1.0 / df.groupby("Produit")["Produit"].transform("count")

        df["G_i (hL)"] = df["Volume disponible (hl)"]
        df["G_total (hL) par goût"] = df.groupby("Produit")["G_i (hL)"].transform("sum")
        df["Y_total (hL) par goût"] = df["G_total (hL) par goût"] + float(volume_cible)
        df["X_th (hL)"] = df["r_i"] * df["Y_total (hL) par goût"] - df["G_i (hL)"]

        df["X_adj (hL)"] = 0.0
        for gout, grp in df.groupby("Produit"):
            x = grp["X_th (hL)"].to_numpy(dtype=float)
            r = grp["r_i"].to_numpy(dtype=float)
            x = np.maximum(x, 0.0)
            deficit = float(volume_cible) - x.sum()
            if deficit > 1e-9:
                r = np.where(r > 0, r, 0)
                s = r.sum()
                x = x + (deficit * (r / s) if s > 0 else deficit / len(x))
            x = np.where(x < 1e-9, 0.0, x)
            df.loc[grp.index, "X_adj (hL)"] = x

        cap_resume = f"{volume_cible:.2f} hL par goût"

    else:
        # Mode 2 goûts : 64 hL AU TOTAL (répartition globale pour épuisement simultané)
        somme_ventes = df["Volume vendu (hl)"].sum()
        if repartir_pro_rv and somme_ventes > 0:
            df["r_i_global"] = df["Volume vendu (hl)"] / somme_ventes
        else:
            df["r_i_global"] = 1.0 / len(df)

        df["G_i (hL)"] = df["Volume disponible (hl)"]
        G_total_all = df["G_i (hL)"].sum()
        Y_total_all = G_total_all + float(volume_cible)   # stocks + prod

        df["X_th (hL)"] = df["r_i_global"] * Y_total_all - df["G_i (hL)"]

        # Ajustement GLOBAL pour ΣX = volume_cible et X>=0
        x = np.maximum(df["X_th (hL)"].to_numpy(dtype=float), 0.0)
        deficit = float(volume_cible) - x.sum()
        if deficit > 1e-9:
            w = df["r_i_global"].to_numpy(dtype=float)
            s = w.sum()
            x = x + (deficit * (w / s) if s > 0 else deficit / len(x))
        x = np.where(x < 1e-9, 0.0, x)
        df["X_adj (hL)"] = x

        cap_resume = f"{volume_cible:.2f} hL au total (2 goûts)"

    # Cartons (exact + arrondi interne)
    df["Cartons à produire (exact)"] = df["X_adj (hL)"] / df["Volume/carton (hL)"]
    if ROUND_TO_CARTON:
        df["Cartons à produire (arrondi)"] = np.floor(df["Cartons à produire (exact)"] + 0.5).astype("Int64")
        df["Volume produit arrondi (hL)"] = df["Cartons à produire (arrondi)"] * df["Volume/carton (hL)"]

    # Sortie principale (simplifiée)
    df_min = df[[
        "Produit", "Stock",
        "Cartons à produire (exact)", "Cartons à produire (arrondi)", "Volume produit arrondi (hL)"
    ]].copy()

    return df_min, cap_resume, gouts_cibles

# ---------- Flow ----------
if uploaded is None:
    st.info("💡 Charge un fichier Excel pour commencer.")
    st.stop()

try:
    df_in = read_input_excel(uploaded)
except Exception as e:
    st.error(f"Erreur de lecture : {e}")
    st.stop()

manual_keep = None
if use_manual:
    all_gouts = sorted(pd.Series(df_in.get("Produit", pd.Series(dtype=str))).dropna().astype(str).unique())
    chosen = st.multiselect("Choisis les goûts à produire", options=all_gouts, default=all_gouts[:nb_gouts])
    manual_keep = chosen

exclude_list = [g.strip() for g in gouts_exclus.split(',') if g.strip()] if gouts_exclus else None

try:
    df_min, cap_resume, gouts_cibles = compute_plan(
        df_in, volume_cible, nb_gouts, repartir_pro_rv, manual_keep, exclude_list
    )
except Exception as e:
    st.error(f"Erreur de calcul : {e}")
    st.stop()

# ---------- Display ----------
st.subheader("Résumé")
st.metric("Goûts sélectionnés", len(gouts_cibles))
st.metric("Capacité utilisée", cap_resume)

st.subheader("Production simplifiée")
st.dataframe(df_min.head(80), use_container_width=True)
