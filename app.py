import io
import re
import numpy as np
import pandas as pd
import streamlit as st

# =========================
# Optimiseur de production (v3.3)
# - UI simple (pas d'options visibles pour l'arrondi ni les formats)
# - 1 goût : 64 hL PAR goût (logique V1)
# - 2 goûts : 64 hL AU TOTAL (répartition globale pour épuisement simultané)
# - Filtre formats (33cl/75cl) et arrondi cartons appliqués EN INTERNE
# =========================

st.set_page_config(page_title="Optimiseur de production — 64 hL / 1–2 goûts", page_icon="🧪", layout="wide")

# -------- Réglages cachés --------
ALLOWED_BOTTLE_L = {0.33, 0.75}   # 33cl & 75cl
ROUND_TO_CARTON = True            # arrondi half-up des cartons en interne
TOL = 0.005                       # tolérance sur volume bouteille

# ---------- Sidebar (UI minimale) ----------
with st.sidebar:
    st.header("Paramètres")
    volume_cible = st.number_input("Volume cible (hL)", min_value=1.0, value=64.0, step=1.0,
                                   help="Si 1 goût: volume PAR goût. Si 2 goûts: volume TOTAL partagé.")
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
    if pd.isna(text):
        return np.nan, np.nan
    s = str(text)
    m_nb = re.search(r"Carton de\s*(\d+)", s, flags=re.IGNORECASE)
    nb = int(m_nb.group(1)) if m_nb else np.nan
    m_l = re.findall(r"(\d+(?:[.,]\\d+)?)\\s*[lL]", s)
    if m_l:
        vol_l = float(m_l[-1].replace(',', '.'))
    else:
        m_cl = re.findall(r"(\d+(?:[.,]\\d+)?)\\s*c[lL]", s)
        vol_l = float(m_cl[-1].replace(',', '.')) / 100.0 if m_cl else np.nan
    return nb, vol_l

def safe_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce")

# ---------- Core calc ----------
def compute_plan(
    df_in: pd.DataFrame,
    volume_cible: float,
    nb_gouts: int,
    repartir_pro_rv: bool,
    manual_keep: list | None,
    exclude_list: list | None,
):
    required = ["Produit", "Stock", "Quantité vendue", "Volume vendu (hl)", "Quantité disponible", "Volume disponible (hl)"]
    missing = [c for c in required if c not in df_in.columns]
    if missing:
        raise ValueError(f"Colonnes manquantes: {missing}")

    df = df_in[required].copy()
    for c in ["Quantité vendue", "Volume vendu (hl)", "Quantité disponible", "Volume disponible (hl)"]:
        df[c] = safe_num(df[c])

    # Parse Stock → nb bouteilles, volume bouteille
    df[["Bouteilles/carton", "Volume bouteille (L)"]] = df["Stock"].apply(lambda s: pd.Series(parse_stock(s)))
    df["Volume/carton (hL)"] = (df["Bouteilles/carton"] * df["Volume bouteille (L)"]) / 100.0

    # Lignes valides
    df = df.dropna(subset=["Produit", "Volume/carton (hL)", "Volume vendu (hl)", "Volume disponible (hl)"]).reset_index(drop=True)

    # Filtre formats caché (33cl & 75cl)
    def is_allowed_l(v):
        if pd.isna(v):
            return False
        return any(abs(v - a) <= TOL for a in ALLOWED_BOTTLE_L)
    mask_allowed = df["Volume bouteille (L)"].apply(is_allowed_l)
    if mask_allowed.any():
        df = df[mask_allowed].reset_index(drop=True)

    # Exclusions / manuel
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

    if nb_gouts == 1:
        # ---- Mode 1 goût : logique V1 (64 hL PAR goût) ----
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
                if s > 0:
                    x = x + deficit * (r / s)
                else:
                    x = x + deficit / len(x)
            x = np.where(x < 1e-9, 0.0, x)
            df.loc[grp.index, "X_adj (hL)"] = x

        cap_resume = f"{volume_cible:.2f} hL par goût"

    else:
        # ---- Mode 2 goûts : 64 hL AU TOTAL (répartition GLOBALE) ----
        somme_ventes = df["Volume vendu (hl)"].sum()
        if repartir_pro_rv and somme_ventes > 0:
            df["r_i_global"] = df["Volume vendu (hl)"] / somme_ventes
        else:
            df["r_i_global"] = 1.0 / len(df)

        df["G_i (hL)"] = df["Volume disponible (hl)"]
        G_total_all = df["G_i (hL)"].sum()
        Y_total_all = G_total_all + float(volume_cible)  # stocks + production

        df["X_th (hL)"] = df["r_i_global"] * Y_total_all - df["G_i (hL)"]

        # Ajustement GLOBAL : ΣX = volume_cible, X>=0
        x = np.maximum(df["X_th (hL)"].to_numpy(dtype=float), 0.0)
        deficit = float(volume_cible) - x.sum()
        if deficit > 1e-9:
            w = df["r_i_global"].to_numpy(dtype=float)
            w = np.where(w > 0, w, 0)
            s = w.sum()
            if s > 0:
                x = x + deficit * (w / s)
            else:
                x = x + deficit / len(x)
        x = np.where(x < 1e-9, 0.0, x)
        df["X_adj (hL)"] = x

        cap_resume = f"{volume_cible:.2f} hL au total (2 goûts)"

    # Cartons exact + arrondi (interne)
    df["Cartons à produire (exact)"] = df["X_adj (hL)"] / df["Volume/carton (hL)"]
    if ROUND_TO_CARTON:
        df["Cartons à produire (arrondi)"] = np.floor(df["Cartons à produire (exact)"] + 0.5).astype("Int64")
        df["Volume produit arrondi (hL)"] = df["Cartons à produire (arrondi)"] * df["Volume/carton (hL)"]
    else:
        df["Cartons à produire (arrondi)"] = pd.NA
        df["Volume produit arrondi (hL)"] = pd.NA

    # Sorties
    df_min = df[[
        "Produit", "Stock",
        "Cartons à produire (exact)", "Cartons à produire (arrondi)", "Volume produit arrondi (hL)"
    ]].copy()

    df_detail_cols = [
        "Produit", "Stock", "Quantité vendue", "Volume vendu (hl)", "Quantité disponible", "Volume disponible (hl)",
        "Bouteilles/carton", "Volume bouteille (L)", "Volume/carton (hL)",
        "G_i (hL)", "X_th (hL)", "X_adj (hL)",
        "Cartons à produire (exact)", "Cartons à produire (arrondi)", "Volume produit arrondi (hL)"
    ]
    if nb_gouts == 1:
        df_detail_cols.insert(9, "Somme ventes (hL) par goût")
        df_detail_cols.insert(10, "r_i")
        df_detail_cols.insert(12, "Y_total (hL) par goût")
    else:
        df_detail_cols.insert(9, "r_i_global")

    df_detail = df[df_detail_cols].copy()

    # Synthèse
    if nb_gouts == 1:
        synth = df.groupby("Produit").agg(
            Formats=("Stock", "count"),
            Ventes_totales_hL=("Volume vendu (hl)", "sum"),
            Stock_restants_hL=("Volume disponible (hl)", "sum"),
            Production_ajustee_hL=("X_adj (hL)", "sum"),
        ).reset_index()
        synth["Capacité"] = cap_resume
    else:
        synth = df.groupby("Produit").agg(
            Formats=("Stock", "count"),
            Ventes_totales_hL=("Volume vendu (hl)", "sum"),
            Stock_restants_hL=("Volume disponible (hl)", "sum"),
            Production_ajustee_hL=("X_adj (hL)", "sum"),
        ).reset_index()
        synth.loc[len(synth.index)] = ["TOTAL", df.shape[0], df["Volume vendu (hl)"].sum(), df["Volume disponible (hl)"].sum(), df["X_adj (hL)"].sum()]
        synth["Capacité"] = cap_resume

    return df_min, df_detail, synth, gouts_cibles, cap_resume

# ---------- Flow ----------
if uploaded is None:
    st.info("💡 Charge un fichier
