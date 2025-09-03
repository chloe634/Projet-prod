import io
import re
import numpy as np
import pandas as pd
import streamlit as st
from pathlib import Path
import difflib

# ------------------------------------------------------------
# Streamlit App
# ------------------------------------------------------------
st.set_page_config(page_title="Plan de production en cartons", page_icon="📦", layout="wide")
st.title("📦 Plan de production en cartons — Goûts & Formats")


# ------------------------------------------------------------
# Utils pour Excel
# ------------------------------------------------------------
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


# ------------------------------------------------------------
# Flavor mapping
# ------------------------------------------------------------
def load_flavor_map(uploaded_file=None) -> pd.DataFrame:
    cols = ["name", "canonical"]
    if uploaded_file is not None:
        try:
            fm = pd.read_csv(uploaded_file)
            fm = fm[[c for c in cols if c in fm.columns]].dropna()
            fm.columns = ["name", "canonical"]
            return fm
        except Exception:
            pass
    p = Path(__file__).parent / "flavor_map.csv"
    if p.exists():
        fm = pd.read_csv(p)
        fm = fm[[c for c in cols if c in fm.columns]].dropna()
        fm.columns = ["name", "canonical"]
        return fm
    return pd.DataFrame(columns=cols)


def apply_canonical_flavor(df: pd.DataFrame, fm: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["Produit_norm"] = out["Produit"].astype(str).str.strip()
    if len(fm):
        mapping = fm.set_index("name")["canonical"].to_dict()
        out["GoutCanon"] = out["Produit_norm"].map(mapping).fillna(out["Produit_norm"])
    else:
        out["GoutCanon"] = out["Produit_norm"]
    return out


def suggest_missing_mappings(df: pd.DataFrame, fm: pd.DataFrame) -> pd.DataFrame:
    known_names = set(fm["name"]) if len(fm) else set()
    unknown = sorted(set(df["Produit"].astype(str).str.strip()) - known_names)
    choices = sorted(set(fm["canonical"])) if len(fm) else sorted(set(df["Produit"].astype(str).str.strip()))
    suggestions = []
    for u in unknown:
        guess = difflib.get_close_matches(u, choices, n=3, cutoff=0.6)
        suggestions.append({"name": u, "suggestions": ", ".join(guess)})
    return pd.DataFrame(suggestions)


# ------------------------------------------------------------
# Calcul principal
# ------------------------------------------------------------
def compute_plan(df_in: pd.DataFrame, volume_cible_par_gout: float = 64.0, rounding: str = "nearest", exclude_list=None):
    col_map = {
        "Produit": "Produit",
        "Stock": "Stock",
        "Quantité vendue": "Quantité vendue",
        "Volume vendu (hl)": "Volume vendu (hl)",
        "Quantité disponible": "Quantité disponible",
        "Volume disponible (hl)": "Volume disponible (hl)",
    }
    missing = [c for c in col_map if c not in df_in.columns]
    if missing:
        raise ValueError(f"Colonnes manquantes dans le fichier: {missing}")

    df = df_in[list(col_map.keys()) + ["GoutCanon"]].copy()

    for c in ["Quantité vendue", "Volume vendu (hl)", "Quantité disponible", "Volume disponible (hl)"]:
        df[c] = safe_numeric(df[c])

    df[["Bouteilles/carton", "Volume bouteille (L)"]] = df["Stock"].apply(lambda s: pd.Series(parse_stock(s)))
    df["Volume/carton (hL)"] = (df["Bouteilles/carton"] * df["Volume bouteille (L)"]) / 100.0
    df = df.dropna(subset=["Produit", "Volume/carton (hL)", "Volume vendu (hl)", "Volume disponible (hl)"]).reset_index(drop=True)

    # --- exclusion case-insensitive ---
    if exclude_list:
        excl = set(x.strip().lower() for x in exclude_list)
        df = df[~df["GoutCanon"].astype(str).str.strip().str.lower().isin(excl)]

    df["Somme ventes (hL) par gout"] = df.groupby("GoutCanon")["Volume vendu (hl)"].transform("sum")
    df["Part ventes (r_i)"] = np.where(df["Somme ventes (hL) par gout"] > 0,
                                        df["Volume vendu (hl)"] / df["Somme ventes (hL) par gout"], 0.0)
    df["Stock restant (G_i, hL)"] = df["Volume disponible (hl)"]
    df["G_total (hL) par gout"] = df.groupby("GoutCanon")["Stock restant (G_i, hL)"].transform("sum")
    df["Y_total (hL) par gout"] = df["G_total (hL) par gout"] + float(volume_cible_par_gout)

    df["X_th (hL)"] = df["Part ventes (r_i)"] * df["Y_total (hL) par gout"] - df["Stock restant (G_i, hL)"]

    df["X_adj (hL)"] = 0.0
    for gout, group in df.groupby("GoutCanon"):
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
                    k = mask.sum()
                    if k > 0:
                        x_adj[mask] += deficit / float(k)
            else:
                x_adj += deficit / len(x_adj)
        x_adj = np.where(x_adj < 1e-9, 0.0, x_adj)
        df.loc[group.index, "X_adj (hL)"] = x_adj

    df["Cartons à produire (exact)"] = df["X_adj (hL)"] / df["Volume/carton (hL)"]
    df["Cartons à produire (arrondi)"] = round_mode(df["Cartons à produire (exact)"], rounding).astype("Int64")
    df["Volume produit arrondi (hL)"] = df["Cartons à produire (arrondi)"] * df["Volume/carton (hL)"]

    df_minimal = df[[
        "GoutCanon", "Produit", "Stock",
        "Cartons à produire (exact)", "Cartons à produire (arrondi)", "Volume produit arrondi (hL)"
    ]].copy()

    return df, df_minimal


# ------------------------------------------------------------
# UI
# ------------------------------------------------------------
with st.sidebar:
    st.header("Paramètres")
    volume_cible = st.number_input("Volume cible par goût (hL)", min_value=1.0, max_value=10000.0, value=64.0, step=1.0)
    rounding = st.selectbox("Arrondi des cartons", ["nearest", "up", "down"], index=0)
    st.subheader("Mapping des goûts")
    uploaded_map = st.file_uploader("Uploader un flavor_map.csv (optionnel)", type=["csv"], key="map")
    st.subheader("Exclure certains goûts (canoniques)")
    excluded_gouts = st.multiselect("Goûts à exclure", [])

uploaded = st.file_uploader("Déposez votre fichier Excel", type=["xls", "xlsx"]) 

if uploaded is None:
    st.info("💡 Charge un fichier Excel pour commencer.")
    st.stop()

try:
    df_in = read_input_excel(uploaded)
except Exception as e:
    st.error(f"Erreur lecture Excel : {e}")
    st.stop()

# Mapping
flavor_map_df = load_flavor_map(uploaded_map)
df_in = apply_canonical_flavor(df_in, flavor_map_df)

# Suggestions
with st.expander("🔎 Produits non mappés (suggestions)"):
    missing = suggest_missing_mappings(df_in, flavor_map_df)
    if len(missing):
        st.write("Ces libellés ne sont pas encore dans flavor_map.csv :")
        st.dataframe(missing, use_container_width=True)
    else:
        st.success("Tous les libellés sont couverts.")

# Normalisation de la liste d'exclusion
excluded_norm = [s.strip().lower() for s in excluded_gouts]

try:
    df_detail, df_minimal = compute_plan(df_in, volume_cible_par_gout=volume_cible, rounding=rounding, exclude_list=excluded_norm)
except Exception as e:
    st.error(f"Erreur pendant les calculs : {e}")
    st.stop()

st.subheader("Production simplifiée (par formats)")
st.dataframe(df_minimal.head(50))
