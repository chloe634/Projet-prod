# pages/04_Achats_conditionnements.py
from __future__ import annotations
import io, re, unicodedata
from typing import Tuple, List, Dict
import numpy as np
import pandas as pd
import streamlit as st

from common.design import apply_theme, section, kpi
from core.optimizer import parse_stock, VOL_TOL  # formats 12x33 / 6x75 / 4x75

# ====================== UI ======================
apply_theme("Achats — Conditionnements", "📦")
section("Prévision d’achats (conditionnements)", "📦")

# Besoin du fichier ventes déjà chargé dans l'accueil
if "df_raw" not in st.session_state or "window_days" not in st.session_state:
    st.warning("Va d’abord dans **Accueil** pour déposer l’Excel des ventes/stock, puis reviens ici.")
    st.stop()

df_raw = st.session_state.df_raw.copy()
window_days = float(st.session_state.window_days)

# ---------------- Sidebar ----------------
with st.sidebar:
    st.header("Période à prévoir")
    horizon_j = st.number_input("Horizon (jours)", min_value=1, max_value=365, value=14, step=1)
    st.caption("Le besoin prévoit une consommation sur cet horizon à partir des ventes moyennes.")

    st.markdown("---")
    st.header("Options étiquettes")
    force_labels = st.checkbox("Étiquettes = 1 par bouteille (forcer si 'étiquette' dans le nom)", value=True)

    st.markdown("---")
    st.header("Fichiers à importer")
    conso_file = st.file_uploader("Consommation des articles (Excel)", type=["xlsx","xls"])
    stock_file = st.file_uploader("Stocks des articles (Excel)", type=["xlsx","xls"])

st.caption(
    f"Excel ventes courant : **{st.session_state.get('file_name','(sans nom)')}** — "
    f"Fenêtre de calcul des vitesses : **{int(window_days)} jours** — "
    f"Horizon prévision : **{int(horizon_j)} jours**"
)

# ====================== Helpers ======================

def _norm_txt(s: str) -> str:
    s = str(s or "").strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s+", " ", s)
    return s

def _find_cell(df_nohdr: pd.DataFrame, pattern: str) -> Tuple[int, int] | Tuple[None, None]:
    pat = _norm_txt(pattern)
    for r in range(df_nohdr.shape[0]):
        row = df_nohdr.iloc[r].astype(str).tolist()
        for c, v in enumerate(row):
            if pat in _norm_txt(v):
                return r, c
    return None, None

@st.cache_data(show_spinner=False)
def read_consumption_xlsx(file) -> pd.DataFrame:
    """
    Extrait la zone:
      - de la cellule à DROITE de 'conditionnement' (même ligne, col+1)
      - jusqu'à la LIGNE avant 'contenants'
    Retourne colonnes: key, article, conso, per_hint ('bottle' par défaut, 'carton' si mot-clé).
    """
    df0 = pd.read_excel(file, header=None, dtype=str)
    r_cond, c_cond = _find_cell(df0, "conditionnement")
    r_stop, _ = _find_cell(df0, "contenants")

    if r_cond is None:
        raise RuntimeError("Mot-clé 'conditionnement' introuvable dans le fichier consommation.")
    if r_stop is None:
        r_stop = df0.shape[0]

    row_start = r_cond
    col_val = c_cond + 1 if (c_cond is not None and c_cond + 1 < df0.shape[1]) else 1
    row_end = max(row_start + 1, r_stop)
    col_article = 0 if 0 < df0.shape[1] else max(0, c_cond - 1)

    block = df0.iloc[row_start+1:row_end, [col_article, col_val]].copy()
    block.columns = ["article", "conso_raw"]
    block["article"] = block["article"].astype(str).str.strip()
    block["conso"] = pd.to_numeric(block["conso_raw"].astype(str).str.replace(",", ".", regex=False), errors="coerce")
    block = block.dropna(subset=["article"])
    block = block[block["article"].str.len() > 0]
    block["conso"] = block["conso"].fillna(0.0)

    def _per_hint(a: str) -> str:
        a0 = _norm_txt(a)
        if any(w in a0 for w in ["carton", "caisse", "colis", "etui", "étui"]):
            return "carton"
        return "bottle"

    block["per_hint"] = block["article"].map(_per_hint)
    block["key"] = block["article"].map(_norm_txt)
    block = block.groupby(["key","article","per_hint"], as_index=False)["conso"].sum()
    return block[["key","article","conso","per_hint"]]

@st.cache_data(show_spinner=False)
def read_stock_xlsx(file) -> pd.DataFrame:
    """Repère l'en-tête 'Quantité virtuelle' et lit les stocks."""
    df0 = pd.read_excel(file, header=None, dtype=str)
    r_hdr, c_q = _find_cell(df0, "quantité virtuelle")
    if r_hdr is None:
        raise RuntimeError("En-tête 'Quantité virtuelle' introuvable dans l'Excel de stocks.")

    name_candidates = ["article", "designation", "désignation", "libelle", "libellé"]
    c_name = None
    for cc in range(df0.shape[1]):
        if _norm_txt(str(df0.iloc[r_hdr, cc])) in name_candidates:
            c_name = cc; break
    if c_name is None:
        for cc in range(max(0, c_q-1), -1, -1):
            if str(df0.iloc[r_hdr, cc]).strip():
                c_name = cc; break
    if c_name is None:
        c_name = 0

    body = df0.iloc[r_hdr+1:, [c_name, c_q]].copy()
    body.columns = ["article", "stock_raw"]
    body["article"] = body["article"].astype(str).str.strip()
    body["stock"] = pd.to_numeric(body["stock_raw"].astype(str).str.replace(",", ".", regex=False), errors="coerce").fillna(0.0)
    body = body[body["article"].str.len() > 0]
    body["key"] = body["article"].map(_norm_txt)
    body = body.groupby(["key","article"], as_index=False)["stock"].sum()
    return body[["key","article","stock"]]

def _fmt_from_stock_text(stock_txt: str) -> str | None:
    """Retourne '12x33' / '6x75' / '4x75' depuis la colonne Stock."""
    nb, vol = parse_stock(stock_txt)
    if pd.isna(nb) or pd.isna(vol): return None
    nb = int(nb); vol = float(vol)
    if nb == 12 and abs(vol - 0.33) <= VOL_TOL: return "12x33"
    if nb == 6  and abs(vol - 0.75) <= VOL_TOL: return "6x75"
    if nb == 4  and abs(vol - 0.75) <= VOL_TOL: return "4x75"
    return None

def aggregate_forecast_by_format(df_sales: pd.DataFrame, window_days: float, horizon_j: int) -> Dict[str, Dict[str, float]]:
    """Calcule bouteilles et cartons prévus par format sur l’horizon (à partir des vitesses)."""
    req = ["Stock", "Volume vendu (hl)"]
    if any(c not in df_sales.columns for c in req):
        return {}

    tmp = df_sales.copy()
    tmp["fmt"] = tmp["Stock"].map(_fmt_from_stock_text)
    tmp = tmp.dropna(subset=["fmt"])
    parsed = tmp["Stock"].map(parse_stock)
    tmp[["nb_btl_cart", "vol_L"]] = pd.DataFrame(parsed.tolist(), index=tmp.index)
    tmp["vol_hL_per_btl"] = (tmp["vol_L"].astype(float) / 100.0)
    tmp["nb_btl_cart"] = pd.to_numeric(tmp["nb_btl_cart"], errors="coerce")
    tmp["v_hL_j"] = pd.to_numeric(tmp["Volume vendu (hl)"], errors="coerce") / max(float(window_days), 1.0)
    tmp = tmp.replace([np.inf, -np.inf], np.nan).dropna(subset=["vol_hL_per_btl", "nb_btl_cart", "v_hL_j"])
    tmp["btl_j"] = np.where(tmp["vol_hL_per_btl"] > 0, tmp["v_hL_j"] / tmp["vol_hL_per_btl"], 0.0)
    tmp["carton_j"] = np.where(tmp["nb_btl_cart"] > 0, tmp["btl_j"] / tmp["nb_btl_cart"], 0.0)
    tmp["btl_h"] = horizon_j * tmp["btl_j"]
    tmp["carton_h"] = horizon_j * tmp["carton_j"]

    agg = tmp.groupby("fmt").agg(bottles=("btl_h","sum"), cartons=("carton_h","sum"))
    out = {fmt: {"bottles": float(agg.loc[fmt, "bottles"]), "cartons": float(agg.loc[fmt, "cartons"])} for fmt in agg.index}
    for k in ["12x33", "6x75", "4x75"]:
        out.setdefault(k, {"bottles": 0.0, "cartons": 0.0})
    return out

def _article_applies_formats(article: str) -> Tuple[List[str], str]:
    """
    Formats cibles + unité par défaut.
    - '12x33' → 12x33 ; '6x75' → 6x75 ; '4x75' → 4x75
    - '33' seul → 12x33 ; '75' → 6x75 & 4x75 ; sinon → tous formats
    - 'carton/caisse/colis/étui' → unité 'carton', sinon 'bottle'
    """
    a = _norm_txt(article)
    per = "carton" if any(w in a for w in ["carton","caisse","colis","etui","étui"]) else "bottle"
    if "12x33" in a: fmts = ["12x33"]
    elif "6x75" in a: fmts = ["6x75"]
    elif "4x75" in a: fmts = ["4x75"]
    elif "33" in a and "75" not in a: fmts = ["12x33"]
    elif "75" in a: fmts = ["6x75","4x75"]
    else: fmts = ["12x33","6x75","4x75"]
    return fmts, per

def compute_needs_table(df_conso: pd.DataFrame, df_stock: pd.DataFrame, forecast_fmt: Dict[str, Dict[str, float]], *, force_labels: bool) -> pd.DataFrame:
    """
    Besoin = conso × (bouteilles ou cartons prévus) agrégé par formats applicables.
    Cas particulier ÉTIQUETTES:
      - si 'étiquette' dans le nom ET option cochée → conso = 1 par bouteille (ignore le fichier de conso)
    """
    rows = []
    for _, r in df_conso.iterrows():
        art = r["article"]; k = r["key"]; conso = float(r["conso"])
        a_norm = _norm_txt(art)
        fmts, per = _article_applies_formats(art)

        # Règle spéciale étiquettes
        if force_labels and ("etiquette" in a_norm or "étiquette" in a_norm or "etiquettes" in a_norm or "étiquettes" in a_norm):
            per = "bottle"
            conso = 1.0  # 1 étiquette par bouteille

        # sinon on respecte le per_hint si présent
        if str(r.get("per_hint","")).strip() in ("bottle","carton"):
            per = str(r["per_hint"]).strip()

        qty = 0.0
        for f in fmts:
            if per == "bottle":
                qty += conso * float(forecast_fmt.get(f, {}).get("bottles", 0.0))
            else:
                qty += conso * float(forecast_fmt.get(f, {}).get("cartons", 0.0))
        rows.append({"key": k, "Article": art, "Unité": "par bouteille" if per=="bottle" else "par carton", "Besoin horizon": qty})

    need_df = pd.DataFrame(rows)
    if need_df.empty:
        return pd.DataFrame(columns=["Article","Unité","Besoin horizon","Stock dispo","À acheter"])

    st_df = (df_stock[["key","stock"]].rename(columns={"stock":"Stock dispo"}) if df_stock is not None else pd.DataFrame(columns=["key","Stock dispo"]))
    out = need_df.merge(st_df, on="key", how="left").fillna({"Stock dispo": 0.0})
    out["À acheter"] = np.maximum(out["Besoin horizon"] - out["Stock dispo"], 0.0)

    for c in ["Besoin horizon","Stock dispo","À acheter"]:
        out[c] = out[c].astype(float)
    return out.drop(columns=["key"]).sort_values("À acheter", ascending=False).reset_index(drop=True)

# ====================== Calculs ======================

# Prévision par format depuis les ventes historiques
forecast = aggregate_forecast_by_format(df_raw, window_days=window_days, horizon_j=int(horizon_j))

# KPIs — on affiche des ÉTIQUETTES (≈ nb de bouteilles) plutôt que “bouteilles”
b_33 = forecast.get("12x33",{}).get("bottles", 0.0)
b_75 = forecast.get("6x75",{}).get("bottles", 0.0) + forecast.get("4x75",{}).get("bottles", 0.0)
cartons_total = sum(v.get("cartons",0.0) for v in forecast.values())

colA, colB, colC = st.columns([1.1, 1, 1])
with colA:
    kpi("Étiquettes à prévoir — 12x33", f"{b_33:.0f}")
with colB:
    kpi("Étiquettes à prévoir — 75cl", f"{b_75:.0f}")
with colC:
    kpi("Cartons prévus (tous formats)", f"{cartons_total:.0f}")

# Chargement des 2 fichiers
df_conso = None
df_stockc = None
err_block = False

if conso_file is not None:
    try:
        df_conso = read_consumption_xlsx(conso_file)
        st.success("Consommation: zone détectée et lue ✅")
        st.dataframe(df_conso[["article","conso","per_hint"]], use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(f"Erreur lecture consommation: {e}")
        err_block = True
else:
    st.info("Importer l’Excel **Consommation des articles** (onglet latéral).")

if stock_file is not None:
    try:
        df_stockc = read_stock_xlsx(stock_file)
        st.success("Stocks: colonne 'Quantité virtuelle' détectée ✅")
        st.dataframe(df_stockc[["article","stock"]], use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(f"Erreur lecture stocks: {e}")
        err_block = True
else:
    st.info("Importer l’Excel **Stocks des articles** (onglet latéral).")

st.markdown("---")

# ====================== Résultat ======================
if (df_conso is not None) and (df_stockc is not None) and (not err_block):
    result = compute_needs_table(df_conso, df_stockc, forecast, force_labels=force_labels)

    if result.empty:
        st.info("Aucun besoin calculé (vérifie les fichiers de consommation/stocks et les correspondances d’articles).")
        st.stop()

    total_buy = float(result["À acheter"].sum())
    nb_items  = int((result["À acheter"] > 0).sum())
    c1, c2 = st.columns(2)
    with c1: kpi("Articles à acheter (nb)", f"{nb_items}")
    with c2: kpi("Quantité totale à acheter (unités)", f"{total_buy:,.0f}".replace(",", " "))

    st.subheader("Proposition d’achats (triée par 'À acheter' décroissant)")
    st.dataframe(
        result[["Article","Unité","Besoin horizon","Stock dispo","À acheter"]],
        use_container_width=True, hide_index=True
    )

    csv_bytes = result.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Exporter la proposition (CSV)",
        data=csv_bytes,
        file_name=f"achats_conditionnements_{int(horizon_j)}j.csv",
        mime="text/csv",
        use_container_width=True,
    )
else:
    st.info("Charge les deux fichiers pour obtenir la proposition d’achats.")
