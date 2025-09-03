import io
import re
import numpy as np
import pandas as pd
import streamlit as st

# ------------------------------------------------------------
# Streamlit App: Plan de production (cartons) par goût + format
# ------------------------------------------------------------

st.set_page_config(page_title="Plan de production en cartons", page_icon="📦", layout="wide")
st.title("📦 Plan de production en cartons — Goûts & Formats")
st.write("Téléversez votre fichier Excel et générez automatiquement les cartons à produire par goût et format, de sorte à écouler les stocks le même jour après cette production.")

# -------------------------
# Utilitaires
# -------------------------
def detect_header_row(df_raw: pd.DataFrame) -> int:
    must_have = {"Produit", "Stock", "Quantité vendue", "Volume vendu (hl)", "Quantité disponible", "Volume disponible (hl)"}
    for i in range(min(10, len(df_raw))):
        row_vals = set(str(x).strip() for x in df_raw.iloc[i].tolist())
        if must_have.issubset(row_vals):
            return i
    return 0

def read_input_excel(uploaded_file: io.BytesIO) -> pd.DataFrame:
    raw = pd.read_excel(uploaded_file, header=None)
    header_idx = detect_header_row(raw)
    df = pd.read_excel(uploaded_file, header=header_idx)
    return df

def parse_stock(text: str):
    if pd.isna(text):
        return np.nan, np.nan
    s = str(text)

    m_nb = re.search(r"Carton de\s*(\d+)", s, flags=re.IGNORECASE)
    nb = int(m_nb.group(1)) if m_nb else np.nan

    vol_l = None
    m_l = re.findall(r"(\d+(?:[.,]\d+)?)\s*[lL]", s)
    if m_l:
        vol_l = float(m_l[-1].replace(',', '.'))
    else:
        m_cl = re.findall(r"(\d+(?:[.,]\d+)?)\s*c[lL]", s)
        if m_cl:
            vol_l = float(m_cl[-1].replace(',', '.')) / 100.0

    if vol_l is None:
        vol_l = np.nan

    return nb, vol_l

def safe_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")

def round_mode(values: pd.Series, mode: str) -> pd.Series:
    v = values.astype(float)
    if mode == "up":
        return np.ceil(v)
    if mode == "down":
        return np.floor(v)
    return np.floor(v + 0.5)

def compute_plan(df_in: pd.DataFrame, volume_cible_par_gout: float = 64.0, rounding: str = "nearest"):
    df = df_in.copy()

    for c in ["Quantité vendue", "Volume vendu (hl)", "Quantité disponible", "Volume disponible (hl)"]:
        df[c] = safe_numeric(df[c])

    df[["Bouteilles/carton", "Volume bouteille (L)"]] = df["Stock"].apply(lambda s: pd.Series(parse_stock(s)))
    df["Volume/carton (hL)"] = (df["Bouteilles/carton"] * df["Volume bouteille (L)"]) / 100.0

    df = df.dropna(subset=["Produit", "Volume/carton (hL)", "Volume vendu (hl)", "Volume disponible (hl)"]).reset_index(drop=True)

    df["Somme ventes (hL) par goût"] = df.groupby("Produit")["Volume vendu (hl)"].transform("sum")
    df["Part ventes (r_i)"] = np.where(df["Somme ventes (hL) par goût"] > 0,
                                       df["Volume vendu (hl)"] / df["Somme ventes (hL) par goût"], 0.0)

    df["Stock restant (G_i, hL)"] = df["Volume disponible (hl)"]
    df["G_total (hL) par goût"] = df.groupby("Produit")["Stock restant (G_i, hL)"].transform("sum")
    df["Y_total (hL) par goût"] = df["G_total (hL) par goût"] + float(volume_cible_par_gout)

    df["X_th (hL)"] = df["Part ventes (r_i)"] * df["Y_total (hL) par goût"] - df["Stock restant (G_i, hL)"]

    df["X_adj (hL)"] = 0.0
    for produit, group in df.groupby("Produit"):
        x_th = group["X_th (hL)"].values.astype(float)
        r = group["Part ventes (r_i)"].values.astype(float)
        x_adj = np.maximum(x_th, 0.0)
        deficit = float(volume_cible_par_gout) - x_adj.sum()
        if deficit > 1e-9:
            mask = r > 0
            if mask.any():
                weights = r.copy()
                weights[~mask] = 0.0
                s = weights.sum()
                if s > 0:
                    x_adj += deficit * (weights / s)
            else:
                x_adj += deficit / len(x_adj)
        x_adj = np.where(x_adj < 1e-9, 0.0, x_adj)
        df.loc[group.index, "X_adj (hL)"] = x_adj

    df["Cartons à produire (exact)"] = df["X_adj (hL)"] / df["Volume/carton (hL)"]
    df["Cartons à produire (arrondi)"] = round_mode(df["Cartons à produire (exact)"], rounding).astype("Int64")
    df["Volume produit arrondi (hL)"] = df["Cartons à produire (arrondi)"] * df["Volume/carton (hL)"]

    df_minimal = df[[
        "Produit", "Stock",
        "Cartons à produire (exact)", "Cartons à produire (arrondi)", "Volume produit arrondi (hL)"
    ]].copy()

    return df_minimal

# -------------------------
# Interface Streamlit
# -------------------------
with st.sidebar:
    st.header("Paramètres")
    volume_cible = st.number_input("Volume cible par goût (hL)", min_value=1.0, max_value=10000.0, value=64.0, step=1.0)
    rounding = st.selectbox("Arrondi des cartons", ["nearest", "up", "down"], index=0)

uploaded = st.file_uploader("Déposez votre fichier Excel", type=["xls", "xlsx"])

if uploaded is None:
    st.info("➡️ Importez un Excel avec les colonnes attendues.")
    st.stop()

try:
    df_in = read_input_excel(uploaded)
    df_minimal = compute_plan(df_in, volume_cible_par_gout=volume_cible, rounding=rounding)
except Exception as e:
    st.error(f"Erreur : {e}")
    st.stop()

st.subheader("Aperçu — Production simplifiée (3 colonnes)")
st.dataframe(df_minimal.head(50))

# Export Excel
output = io.BytesIO()
with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    df_minimal.to_excel(writer, index=False, sheet_name="Production simplifiée")
output.seek(0)

st.download_button(
    label="💾 Télécharger Excel — version simplifiée",
    data=output,
    file_name="plan_production_cartons_minimal.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
