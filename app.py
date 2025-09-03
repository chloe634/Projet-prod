import io
import re
import numpy as np
import pandas as pd
import streamlit as st

# ------------------------------------------------------------
# Optimiseur de production (v3)
# - Sidebar "Paramètres"
# - Capacité totale (hL) & Nombre de goûts simultanés
# - Option : répartition par formats au prorata des vitesses de vente
# - PAS d'arrondi au carton, PAS de filtres de formats
# - (Optionnel) Contraintes goûts : sélection manuelle / exclusion
# ------------------------------------------------------------

st.set_page_config(page_title="Optimiseur de production / multi-goûts", page_icon="🧪", layout="wide")

# ======= Sidebar =======
with st.sidebar:
    st.header("Paramètres")
    capacite_totale_hl = st.number_input("Capacité de production (hl)", min_value=1.0, value=64.0, step=1.0)
    nb_gouts = st.selectbox("Nombre de goûts simultanés", options=list(range(1, 11)), index=1)
    repartir_pro_rv = st.checkbox(
        "Répartir par formats au prorata des vitesses de vente",
        value=True,
        help="Si désactivé, la production d'un goût est répartie équitablement entre ses formats."
    )

    st.markdown("---")
    st.subheader("Contraintes goûts")
    use_manual = st.checkbox("Sélection manuelle des goûts", value=False)
    gouts_exclus = st.text_input("Exclure goûts (séparés par des virgules)", value="")

# ======= Header =======
st.title("🧪 Optimiseur de production — 64 hl / 2 goûts (v3)")
st.caption("Charge un Excel d'autonomie, choisis tes options, et génère un plan propre pour l'atelier.")

# ======= Upload =======
uploaded = st.file_uploader("Dépose ton fichier Excel (.xlsx)", type=["xlsx", "xls"]) 

# ------------------------- Utils -------------------------

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
    # volume en L : prend le dernier motif trouvé (ex: 0.75L)
    m_l = re.findall(r"(\d+(?:[.,]\d+)?)\s*[lL]", s)
    if m_l:
        vol_l = float(m_l[-1].replace(',', '.'))
    else:
        m_cl = re.findall(r"(\d+(?:[.,]\d+)?)\s*c[lL]", s)
        vol_l = float(m_cl[-1].replace(',', '.')) / 100.0 if m_cl else np.nan
    return nb, vol_l


def safe_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce")

# ------------------------- Core calc -------------------------

def compute_plan(
    df_in: pd.DataFrame,
    capacite_totale_hl: float,
    nb_gouts: int,
    repartir_pro_rv: bool,
    manual_keep: list | None,
    exclude_list: list | None,
):
    # Nettoyage colonnes
    required = [
        "Produit", "Stock", "Quantité vendue", "Volume vendu (hl)", "Quantité disponible", "Volume disponible (hl)"
    ]
    missing = [c for c in required if c not in df_in.columns]
    if missing:
        raise ValueError(f"Colonnes manquantes: {missing}")

    df = df_in[required].copy()
    for c in ["Quantité vendue", "Volume vendu (hl)", "Quantité disponible", "Volume disponible (hl)"]:
        df[c] = safe_num(df[c])

    # Parser Stock → nb bouteilles et volume bouteille (L)
    df[["Bouteilles/carton", "Volume bouteille (L)"]] = df["Stock"].apply(lambda s: pd.Series(parse_stock(s)))
    df["Volume/carton (hL)"] = (df["Bouteilles/carton"] * df["Volume bouteille (L)"]) / 100.0

    # Filtrer lignes valides
    df = df.dropna(subset=["Produit", "Volume/carton (hL)", "Volume vendu (hl)", "Volume disponible (hl)"]).reset_index(drop=True)

    # Exclure / garder
    if exclude_list:
        excl = [g.strip() for g in exclude_list]
        df = df[~df["Produit"].astype(str).str.strip().isin(excl)]
    if manual_keep:
        keep = [g.strip() for g in manual_keep]
        df = df[df["Produit"].astype(str).str.strip().isin(keep)]

    # Sélection automatique des goûts (si pas manuel) : top N par ventes hL
    ventes_par_gout = df.groupby("Produit")["Volume vendu (hl)"].sum().sort_values(ascending=False)
    if not manual_keep:
        gouts_cibles = ventes_par_gout.index.tolist()[:nb_gouts]
        df = df[df["Produit"].isin(gouts_cibles)]
    else:
        gouts_cibles = sorted(set(df["Produit"]))

    if len(gouts_cibles) == 0:
        raise ValueError("Aucun goût sélectionné.")

    # Capacité par goût = capacité totale / nb goûts
    cap_par_gout = float(capacite_totale_hl) / max(1, nb_gouts)

    # Poids par format au sein d'un goût
    df["Somme ventes (hL) par goût"] = df.groupby("Produit")["Volume vendu (hl)"].transform("sum")
    if repartir_pro_rv:
        df["Poids format"] = np.where(
            df["Somme ventes (hL) par goût"] > 0,
            df["Volume vendu (hl)"] / df["Somme ventes (hL) par goût"],
            1.0,
        )
    else:
        df["Poids format"] = 1.0 / df.groupby("Produit")["Produit"].transform("count")

    # Totaux
    df["Stock restant (G_i, hL)"] = df["Volume disponible (hl)"]
    df["G_total (hL) par goût"] = df.groupby("Produit")["Stock restant (G_i, hL)"].transform("sum")
    df["Y_total (hL) par goût"] = df["G_total (hL) par goût"] + cap_par_gout

    # X théorique
    df["X_th (hL)"] = df["Poids format"] * df["Y_total (hL) par goût"] - df["Stock restant (G_i, hL)"]

    # Ajustement : X>=0 et somme par goût = cap_par_gout
    df["X_adj (hL)"] = 0.0
    for gout, grp in df.groupby("Produit"):
        x = grp["X_th (hL)"].to_numpy(dtype=float)
        w = grp["Poids format"].to_numpy(dtype=float)
        x = np.maximum(x, 0.0)
        deficit = cap_par_gout - x.sum()
        if deficit > 1e-9:
            w = np.where(w > 0, w, 0)
            s = w.sum()
            if s > 0:
                x = x + deficit * (w / s)
            else:
                x = x + deficit / len(x)
        x = np.where(x < 1e-9, 0.0, x)
        df.loc[grp.index, "X_adj (hL)"] = x

    # Sorties
    df_min = df[[
        "Produit", "Stock", "Bouteilles/carton", "Volume bouteille (L)", "Volume/carton (hL)",
        "X_adj (hL)"
    ]].copy()
    df_min.rename(columns={"X_adj (hL)": "Volume à produire (hL)"}, inplace=True)

    df_detail = df[[
        "Produit", "Stock", "Quantité vendue", "Volume vendu (hl)", "Quantité disponible", "Volume disponible (hl)",
        "Bouteilles/carton", "Volume bouteille (L)", "Volume/carton (hL)",
        "Somme ventes (hL) par goût", "Poids format",
        "G_total (hL) par goût", "Y_total (hL) par goût",
        "X_th (hL)", "X_adj (hL)"
    ]].copy()

    synth = df_detail.groupby("Produit").agg(
        Formats=("Stock", "count"),
        Ventes_totales_hL=("Volume vendu (hl)", "sum"),
        Stock_restants_hL=("Volume disponible (hl)", "sum"),
        Production_ajustee_hL=("X_adj (hL)", "sum"),
    ).reset_index()
    synth["Capacité par goût (hL)"] = cap_par_gout
    synth["Delta vs capacité"] = synth["Production_ajustee_hL"] - cap_par_gout

    return df_min, df_detail, synth, gouts_cibles, cap_par_gout

# ======= Main flow =======
if uploaded is None:
    st.info("💡 Charge un fichier Excel pour commencer.")
    st.stop()

try:
    df_in = read_input_excel(uploaded)
except Exception as e:
    st.error(f"Erreur de lecture : {e}")
    st.stop()

# Contraintes goûts
manual_keep = None
if use_manual:
    all_gouts = sorted(pd.Series(df_in.get("Produit", pd.Series(dtype=str))).dropna().astype(str).unique())
    chosen = st.multiselect("Choisis les goûts à produire", options=all_gouts, default=all_gouts[:nb_gouts])
    manual_keep = chosen

exclude_list = [g.strip() for g in gouts_exclus.split(',') if g.strip()] if gouts_exclus else None

try:
    df_min, df_detail, synth, gouts_cibles, cap_par_gout = compute_plan(
        df_in,
        capacite_totale_hl=capacite_totale_hl,
        nb_gouts=nb_gouts,
        repartir_pro_rv=repartir_pro_rv,
        manual_keep=manual_keep,
        exclude_list=exclude_list,
    )
except Exception as e:
    st.error(f"Erreur de calcul : {e}")
    st.stop()

# ======= Layout display =======
left, right = st.columns([1, 2])
with left:
    st.markdown("### Résumé")
    st.metric("Goûts sélectionnés", len(gouts_cibles))
    st.metric("Capacité par goût (hL)", f"{cap_par_gout:.2f}")
with right:
    st.markdown("### Aperçu — Plan simplifié")
    st.dataframe(df_min.head(50), use_container_width=True)

with st.expander("Voir la synthèse par goût"):
    st.dataframe(synth, use_container_width=True)

with st.expander("Voir le détail complet des calculs"):
    st.dataframe(df_detail, use_container_width=True)

# ======= Exports =======
col1, col2 = st.columns(2)
with col1:
    output_min = io.BytesIO()
    with pd.ExcelWriter(output_min, engine="xlsxwriter") as w:
        df_min.to_excel(w, index=False, sheet_name="Plan simplifié")
        synth.to_excel(w, index=False, sheet_name="Synthèse")
    output_min.seek(0)
    st.download_button(
        "💾 Télécharger — Plan simplifié",
        data=output_min,
        file_name="plan_production_simplifie.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
with col2:
    output_full = io.BytesIO()
    with pd.ExcelWriter(output_full, engine="xlsxwriter") as w:
        df_detail.to_excel(w, index=False, sheet_name="Plan détaillé")
        synth.to_excel(w, index=False, sheet_name="Synthèse")
    output_full.seek(0)
    st.download_button(
        "💾 Télécharger — Version complète",
        data=output_full,
        file_name="plan_production_complet.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.caption("© Optimiseur — Water-filling + répartition par format. Sans arrondi carton, sans filtres de format.")
