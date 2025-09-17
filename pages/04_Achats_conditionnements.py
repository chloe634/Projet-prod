# pages/04_Achats_conditionnements.py
from __future__ import annotations
import re, unicodedata
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

# ---------------- Sidebar (période + options) ----------------
with st.sidebar:
    st.header("Période à prévoir")
    horizon_j = st.number_input("Horizon (jours)", min_value=1, max_value=365, value=14, step=1)
    st.caption("Le besoin prévoit une consommation sur cet horizon à partir des ventes moyennes.")
    st.markdown("---")
    st.header("Options étiquettes")
    force_labels = st.checkbox("Étiquettes = 1 par bouteille (forcer si 'étiquette' dans le nom)", value=True)

st.caption(
    f"Excel ventes courant : **{st.session_state.get('file_name','(sans nom)')}** — "
    f"Fenêtre de calcul des vitesses : **{int(window_days)} jours** — "
    f"Horizon prévision : **{int(horizon_j)} jours**"
)

# ====================== IMPORTS (dans la page) ======================
section("Importer les fichiers", "📥")
c1, c2 = st.columns(2)
with c1:
    st.subheader("Consommation des articles (Excel)")
    conso_file = st.file_uploader(
        "Déposer le fichier *Consommation* ici",
        type=["xlsx","xls"],
        key="uploader_conso",
        label_visibility="collapsed"
    )
with c2:
    st.subheader("Stocks des articles (Excel)")
    stock_file = st.file_uploader(
        "Déposer le fichier *Stocks* ici",
        type=["xlsx","xls"],
        key="uploader_stock",
        label_visibility="collapsed"
    )

# ====================== Helpers ======================

def _norm_txt(s: str) -> str:
    s = str(s or "").strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s+", " ", s)
    return s

def _canon_txt(s: str) -> str:
    s = str(s or "")
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"[^a-zA-Z0-9]+", " ", s).strip().lower()
    return s

def _is_total_row(s: str) -> bool:
    """True si libellé est une ligne de total (TOTAL, Total général, …)."""
    t = _canon_txt(s)
    if not t:
        return False
    if t.startswith("total"):
        return True
    return t in {
        "total general", "grand total", "totaux", "total stock",
        "total stocks", "total consommation", "total consommations",
        "total achats", "total des achats"
    }

def _find_cell(df_nohdr: pd.DataFrame, pattern: str) -> Tuple[int | None, int | None]:
    pat = _norm_txt(pattern)
    for r in range(df_nohdr.shape[0]):
        row = df_nohdr.iloc[r].astype(str).tolist()
        for c, v in enumerate(row):
            if pat in _norm_txt(v):
                return r, c
    return None, None

def _parse_number(x: str | float | int) -> float:
    """Tolère , décimales et séparateurs d'espace/point pour milliers."""
    if isinstance(x, (int, float)) and not pd.isna(x):
        return float(x)
    s = str(x or "").strip()
    if not s:
        return np.nan
    s = s.replace("\u202f", " ").replace("\xa0", " ")
    # si on a à la fois . et , on suppose , = décimal si dernier séparateur est ,
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    else:
        s = s.replace(" ", "")
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return np.nan

@st.cache_data(show_spinner=False)
def read_consumption_xlsx(file) -> pd.DataFrame:
    """
    Extrait la zone :
      - colonne ARTICLE = la colonne où se trouve le mot 'conditionnement'
      - colonne CONSO   = la colonne immédiatement à droite de 'conditionnement'
    Lignes : à partir de la ligne sous 'conditionnement' et jusqu'à **2 lignes avant**
    la ligne qui contient 'contenants'. Ignore les lignes 'TOTAL'.
    Retourne colonnes: key, article, conso, per_hint.
    """
    df0 = pd.read_excel(file, header=None, dtype=str)

    # -- utilitaires locaux
    def _norm_txt_local(s: str) -> str:
        s = str(s or "").strip().lower()
        s = unicodedata.normalize("NFKD", s)
        s = "".join(ch for ch in s if not unicodedata.combining(ch))
        s = re.sub(r"\s+", " ", s)
        return s

    # Trouver la cellule 'conditionnement' (on prend la meilleure ancre = celle qui a
    # le plus de lignes non vides en dessous dans la même colonne).
    anchors = []
    for r in range(df0.shape[0]):
        for c in range(df0.shape[1]):
            if "conditionnement" in _norm_txt_local(df0.iat[r, c]):
                # compte de lignes non vides sous l'ancre
                k = 0
                rr = r + 1
                while rr < df0.shape[0] and str(df0.iat[rr, c]).strip():
                    k += 1
                    rr += 1
                anchors.append((k, r, c))
    if not anchors:
        raise RuntimeError("Mot-clé 'conditionnement' introuvable dans le fichier consommation.")
    _, r_cond, c_cond = max(anchors)  # meilleure ancre

    # Limite basse : 2 lignes avant la 1re occurrence de 'contenants' située APRÈS l'ancre
    r_stop = None
    for r in range(r_cond + 1, df0.shape[0]):
        row_txt = " ".join(str(x) for x in df0.iloc[r].tolist())
        if "contenants" in _norm_txt_local(row_txt):
            r_stop = r
            break
    if r_stop is None:
        r_stop = df0.shape[0]

    row_start = r_cond + 1
    row_end   = max(row_start, r_stop - 2)  # 2 lignes avant 'contenants'

    # Colonnes fixes par exigence : Article = colonne ancre ; Conso = colonne à droite
    col_article = c_cond
    col_val_prefered = c_cond + 1

    # Si la colonne préférée ne contient aucun numérique, on cherche la 1re colonne
    # à droite contenant des numériques (séparateur , ou . géré).
    def _count_numeric(col_idx: int) -> int:
        vals = df0.iloc[row_start:row_end, col_idx].astype(str)
        vals = vals.str.replace(",", ".", regex=False)
        x = pd.to_numeric(vals, errors="coerce")
        return int(x.notna().sum())

    col_val = col_val_prefered
    if col_val >= df0.shape[1] or _count_numeric(col_val) == 0:
        best = None
        for cc in range(col_val_prefered, df0.shape[1]):
            cnt = _count_numeric(cc)
            if cnt > 0:
                best = (cnt, cc) if best is None or cnt > best[0] else best
        if best is None:
            raise RuntimeError(
                "Impossible de trouver une colonne de **consommation** numérique à droite de 'conditionnement'."
            )
        col_val = best[1]

    # Extraction du bloc
    block = df0.iloc[row_start:row_end, [col_article, col_val]].copy()
    block.columns = ["article", "conso_raw"]
    block["article"] = block["article"].astype(str).str.strip()

    # Nettoie / ignore les totaux
    def _is_total_row_local(s: str) -> bool:
        t = _norm_txt_local(s)
        return bool(t) and (t.startswith("total") or t in {
            "total general", "grand total", "totaux", "total stock",
            "total stocks", "total consommation", "total consommations",
            "total achats", "total des achats"
        })

    block = block[block["article"].map(lambda s: not _is_total_row_local(s))]

    # Numérise la conso (gestion des virgules)
    block["conso"] = pd.to_numeric(
        block["conso_raw"].astype(str).str.replace(",", ".", regex=False),
        errors="coerce"
    ).fillna(0.0)

    # Heuristique unité (par carton si mots-clés ; sinon par bouteille)
    def _per_hint(a: str) -> str:
        a0 = _norm_txt_local(a)
        return "carton" if any(w in a0 for w in ["carton", "caisse", "colis", "etui", "étui"]) else "bottle"

    block["per_hint"] = block["article"].map(_per_hint)
    block["key"] = block["article"].map(_norm_txt_local)

    # Agrégat au cas où des lignes dupliquées existent
    block = block.groupby(["key", "article", "per_hint"], as_index=False)["conso"].sum()

    return block[["key", "article", "conso", "per_hint"]]


@st.cache_data(show_spinner=False)
def read_stock_xlsx(file) -> pd.DataFrame:
    """Repère l'en-tête 'Quantité virtuelle' et lit les stocks (en filtrant les TOTAL)."""
    df0 = pd.read_excel(file, header=None, dtype=str)
    r_hdr, c_q = _find_cell(df0, "quantité virtuelle")
    if r_hdr is None:
        raise RuntimeError("En-tête 'Quantité virtuelle' introuvable dans l'Excel de stocks.")

    # Essaye de trouver le libellé article sur la même ligne que l'en-tête
    name_candidates = {"article", "designation", "désignation", "libelle", "libellé"}
    c_name = None
    for cc in range(df0.shape[1]):
        if _norm_txt(str(df0.iloc[r_hdr, cc])) in name_candidates:
            c_name = cc
            break
    if c_name is None:
        # fallback : colonne à gauche de "Quantité virtuelle" si non vide
        for cc in range(max(0, c_q - 1), -1, -1):
            if str(df0.iloc[r_hdr, cc]).strip():
                c_name = cc
                break
    if c_name is None:
        c_name = 0

    body = df0.iloc[r_hdr + 1 :, [c_name, c_q]].copy()
    body.columns = ["article", "stock_raw"]
    body["article"] = body["article"].astype(str).str.strip()
    body = body[body["article"].str.len() > 0]
    body = body[~body["article"].map(_is_total_row)]

    body["stock"] = pd.to_numeric(body["stock_raw"].map(_parse_number), errors="coerce").fillna(0.0)
    body["key"] = body["article"].map(_norm_txt)
    body = body.groupby(["key", "article"], as_index=False)["stock"].sum()
    return body[["key", "article", "stock"]]

def _fmt_from_stock_text(stock_txt: str) -> str | None:
    """Retourne '12x33' / '6x75' / '4x75' depuis la colonne Stock."""
    nb, vol = parse_stock(stock_txt)
    if pd.isna(nb) or pd.isna(vol):
        return None
    nb = int(nb)
    vol = float(vol)
    if nb == 12 and abs(vol - 0.33) <= VOL_TOL:
        return "12x33"
    if nb == 6 and abs(vol - 0.75) <= VOL_TOL:
        return "6x75"
    if nb == 4 and abs(vol - 0.75) <= VOL_TOL:
        return "4x75"
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

    agg = tmp.groupby("fmt").agg(bottles=("btl_h", "sum"), cartons=("carton_h", "sum"))
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
    per = "carton" if any(w in a for w in ["carton", "caisse", "colis", "etui", "étui"]) else "bottle"
    if "12x33" in a:
        fmts = ["12x33"]
    elif "6x75" in a:
        fmts = ["6x75"]
    elif "4x75" in a:
        fmts = ["4x75"]
    elif "33" in a and "75" not in a:
        fmts = ["12x33"]
    elif "75" in a:
        fmts = ["6x75", "4x75"]
    else:
        fmts = ["12x33", "6x75", "4x75"]
    return fmts, per

def compute_needs_table(
    df_conso: pd.DataFrame,
    df_stock: pd.DataFrame,
    forecast_fmt: Dict[str, Dict[str, float]],
    *,
    force_labels: bool
) -> pd.DataFrame:
    """
    Besoin = coefficient d'article × (bouteilles OU cartons prévus sur l’horizon).
    Normalisations:
      - ÉTIQUETTES : 1 par bouteille (si option cochée)
      - CAPSULES   : 1 par bouteille
      - CARTONS (12x33 / 6x75 / 4x75) : 1 par carton
      - Autres articles : on garde la valeur de la colonne B comme coefficient.
    """
    rows: List[Dict] = []

    if df_conso is None or df_conso.empty:
        return pd.DataFrame(columns=["Article", "Unité", "Besoin horizon", "Stock dispo", "À acheter"])

    for _, r in df_conso.iterrows():
        art = str(r.get("article", "")).strip()
        if not art or _is_total_row(art):
            continue

        k = r.get("key", _norm_txt(art))
        conso_file = pd.to_numeric(r.get("conso", 0), errors="coerce")
        conso_file = float(0 if pd.isna(conso_file) else conso_file)

        a_norm = _norm_txt(art)
        fmts, per = _article_applies_formats(art)

        # --- normalisation des consommables standards ---
        if force_labels and ("etiquette" in a_norm or "étiquette" in a_norm):
            per = "bottle"
            conso = 1.0
        elif "capsule" in a_norm:
            per = "bottle"
            conso = 1.0
        elif ("carton" in a_norm or "caisse" in a_norm or "colis" in a_norm or "etui" in a_norm or "étui" in a_norm) and (
            "12x33" in a_norm or "6x75" in a_norm or "4x75" in a_norm or (("33" in a_norm) and ("75" not in a_norm)) or ("75" in a_norm)
        ):
            per = "carton"
            conso = 1.0
        else:
            conso = conso_file

        hint = str(r.get("per_hint", "")).strip().lower()
        if hint in ("bottle", "carton"):
            per = hint

        qty = 0.0
        for f in fmts:
            if per == "bottle":
                qty += conso * float(forecast_fmt.get(f, {}).get("bottles", 0.0))
            else:
                qty += conso * float(forecast_fmt.get(f, {}).get("cartons", 0.0))

        rows.append({
            "key": k,
            "Article": art,
            "Unité": "par bouteille" if per == "bottle" else "par carton",
            "Besoin horizon": qty
        })

    need_df = pd.DataFrame(rows)
    if need_df.empty:
        return pd.DataFrame(columns=["Article", "Unité", "Besoin horizon", "Stock dispo", "À acheter"])

    st_df = (
        df_stock[["key", "stock"]].rename(columns={"stock": "Stock dispo"})
        if df_stock is not None else
        pd.DataFrame(columns=["key", "Stock dispo"])
    )
    out = need_df.merge(st_df, on="key", how="left")
    out["Stock dispo"] = pd.to_numeric(out["Stock dispo"], errors="coerce").fillna(0.0)

    out["À acheter"] = np.maximum(out["Besoin horizon"] - out["Stock dispo"], 0.0)

    for c in ["Besoin horizon", "Stock dispo", "À acheter"]:
        out[c] = np.round(out[c], 0).astype(int)

    return (
        out.drop(columns=["key"])
           .sort_values("À acheter", ascending=False)
           .reset_index(drop=True)
    )

# ====================== Calculs ======================

# Prévision par format depuis les ventes historiques
forecast = aggregate_forecast_by_format(df_raw, window_days=window_days, horizon_j=int(horizon_j))

# KPIs — on affiche des ÉTIQUETTES (≈ nb de bouteilles) plutôt que “bouteilles”
b_33 = forecast.get("12x33", {}).get("bottles", 0.0)
b_75 = forecast.get("6x75", {}).get("bottles", 0.0) + forecast.get("4x75", {}).get("bottles", 0.0)
cartons_total = sum(v.get("cartons", 0.0) for v in forecast.values())

colA, colB, colC = st.columns([1.1, 1, 1])
with colA:
    kpi("Étiquettes à prévoir — 12x33", f"{b_33:.0f}")
with colB:
    kpi("Étiquettes à prévoir — 75cl", f"{b_75:.0f}")
with colC:
    kpi("Cartons prévus (tous formats)", f"{cartons_total:.0f}")

# ====================== Lecture fichiers + résultat ======================

df_conso = None
df_stockc = None
err_block = False

if conso_file is not None:
    try:
        df_conso = read_consumption_xlsx(conso_file)
        st.success("Consommation: zone détectée ✅")
        with st.expander("Voir l’aperçu du fichier **Consommation**", expanded=False):
            st.dataframe(df_conso[["article", "conso", "per_hint"]], use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(f"Erreur lecture consommation: {e}")
        err_block = True
else:
    st.info("Importer l’Excel **Consommation des articles** (bloc ci-dessus).")

if stock_file is not None:
    try:
        df_stockc = read_stock_xlsx(stock_file)
        st.success("Stocks: colonne 'Quantité virtuelle' détectée ✅")
        with st.expander("Voir l’aperçu du fichier **Stocks**", expanded=False):
            st.dataframe(df_stockc[["article", "stock"]], use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(f"Erreur lecture stocks: {e}")
        err_block = True
else:
    st.info("Importer l’Excel **Stocks des articles** (bloc ci-dessus).")

st.markdown("---")

if (df_conso is not None) and (df_stockc is not None) and (not err_block):
    result = compute_needs_table(df_conso, df_stockc, forecast, force_labels=force_labels)

    if result.empty:
        st.info("Aucun besoin calculé (vérifie les fichiers de consommation/stocks et les correspondances d’articles).")
        st.stop()

    total_buy = int(result["À acheter"].sum())
    nb_items = int((result["À acheter"] > 0).sum())
    c1, c2 = st.columns(2)
    with c1:
        kpi("Articles à acheter (nb)", f"{nb_items}")
    with c2:
        kpi("Quantité totale à acheter (unités)", f"{total_buy:,}".replace(",", " "))

    st.subheader("Proposition d’achats (triée par 'À acheter' décroissant)")
    st.dataframe(
        result[["Article", "Unité", "Besoin horizon", "Stock dispo", "À acheter"]],
        use_container_width=True,
        hide_index=True,
        column_config={
            "Besoin horizon": st.column_config.NumberColumn(format="%d"),
            "Stock dispo": st.column_config.NumberColumn(format="%d"),
            "À acheter": st.column_config.NumberColumn(format="%d"),
        },
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
