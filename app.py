import io
import re
import numpy as np
import pandas as pd
import streamlit as st

# =========================
# Optimiseur de production (v4.4)
# - Sélection intelligente (autonomie + ventes)
# - 1 goût : 64 hL PAR goût ; 2 goûts : 64 hL AU TOTAL (épuisement simultané)
# - Formats internes: 12×0.33 L, 6×0.75 L, 4×0.75 L (parseur robuste)
# - Arrondi au carton (half-up)
# - Lecture Excel: ignore les lignes contenant au moins UNE cellule au fond noir
# - CA & pertes estimées avec PRIX SAISIS MANUELLEMENT (0.33L et 0.75L)
# =========================

st.set_page_config(page_title="Optimiseur de production — 64 hL / 1–2 goûts", page_icon="🧪", layout="wide")

ALLOWED_FORMATS = {(12, 0.33), (6, 0.75), (4, 0.75)}
ROUND_TO_CARTON = True
VOL_TOL = 0.02   # tolérance sur 0.33 / 0.75 (L)
EPS = 1e-9

# ---------- Sidebar (fixe) ----------
with st.sidebar:
    st.header("Paramètres")
    volume_cible = st.number_input(
        "Volume cible (hL)", min_value=1.0, value=64.0, step=1.0,
        help="Si 1 goût: volume PAR goût. Si 2 goûts: volume TOTAL partagé."
    )
    nb_gouts = st.selectbox("Nombre de goûts simultanés", [1, 2], index=0)
    repartir_pro_rv = st.checkbox(
        "Répartir par formats au prorata des vitesses de vente",
        value=True,
        help="Si décoché: répartition égale entre formats d'un même goût."
    )

    with st.expander("Options avancées"):
        window_days = st.number_input("Fenêtre de ventes (jours)", min_value=7, max_value=120, value=60, step=1)

    st.markdown("---")
    st.subheader("Prix par bouteille (€)")
    price_033 = st.number_input("Prix 0,33 L (€ / bouteille)", min_value=0.0, value=1.75, step=0.01, format="%.2f")
    price_075 = st.number_input("Prix 0,75 L (€ / bouteille)", min_value=0.0, value=3.10, step=0.01, format="%.2f")

# ---------- Header ----------
st.title("🧪 Optimiseur de production — 64 hL / 1–2 goûts")
st.caption("Sélection auto (autonomie + ventes), plan par formats pour écoulement simultané, et estimation des pertes de CA (prix éditables).")

# ---------- Upload ----------
uploaded = st.file_uploader("Dépose ton fichier Excel (.xlsx/.xls)", type=["xlsx", "xls"])

# ---------- Utils : détection header ----------
def detect_header_row(df_raw: pd.DataFrame) -> int:
    must = {"Produit", "Stock", "Quantité vendue", "Volume vendu (hl)", "Quantité disponible", "Volume disponible (hl)"}
    for i in range(min(10, len(df_raw))):
        if must.issubset(set(str(x).strip() for x in df_raw.iloc[i].tolist())):
            return i
    return 0

# ---------- Utils : filtrer lignes à fond noir ----------
def rows_to_keep_by_fill(excel_bytes: bytes, header_idx: int) -> list[bool]:
    """
    Renvoie une liste booléenne (True=à garder) pour les lignes de données,
    en excluant toute ligne qui contient AU MOINS une cellule avec un fond noir.
    """
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
    ws = wb[wb.sheetnames[0]]
    start_row = header_idx + 2  # données juste après l'en-tête
    keep = []
    for r in range(start_row, ws.max_row + 1):
        is_black = False
        for cell in ws[r]:
            fill = cell.fill
            if fill and fill.fill_type:
                rgb = getattr(getattr(fill, "fgColor", None), "rgb", None) or getattr(getattr(fill, "start_color", None), "rgb", None)
                if rgb and rgb[-6:].upper() == "000000":
                    is_black = True
                    break
        keep.append(not is_black)
    return keep

# ---------- Lecture Excel (avec filtre fond noir) ----------
def read_input_excel(uploaded_file) -> pd.DataFrame:
    file_bytes = uploaded_file.read()  # on lit une seule fois
    # 1) détecter l'en-tête
    raw = pd.read_excel(io.BytesIO(file_bytes), header=None)
    header_idx = detect_header_row(raw)
    # 2) lire les données avec l'en-tête
    df = pd.read_excel(io.BytesIO(file_bytes), header=header_idx)
    # 3) filtrer les lignes au fond noir
    keep_mask = rows_to_keep_by_fill(file_bytes, header_idx)
    if len(keep_mask) < len(df):
        keep_mask = keep_mask + [True] * (len(df) - len(keep_mask))
    df = df.iloc[[i for i, k in enumerate(keep_mask) if k]].reset_index(drop=True)
    return df

# --------- Parse "Stock" robuste ---------
def parse_stock(text: str):
    if pd.isna(text): return np.nan, np.nan
    s = str(text)

    nb = None
    for pat in [
        r"(?:Carton|Caisse|Colis)\s+de\s*(\d+)",
        r"(\d+)\s*[x×]\s*Bouteilles?",
        r"(\d+)\s*[x×]",
        r"(\d+)\s+Bouteilles?",
    ]:
        m = re.search(pat, s, flags=re.IGNORECASE)
        if m:
            try:
                nb = int(m.group(1))
                break
            except:
                pass

    vol_l = None
    m_l = re.findall(r"(\d+(?:[.,]\d+)?)\s*[lL]", s)
    if m_l:
        vol_l = float(m_l[-1].replace(",", "."))
    else:
        m_cl = re.findall(r"(\d+(?:[.,]\d+)?)\s*c[lL]", s)
        if m_cl:
            vol_l = float(m_cl[-1].replace(",", ".")) / 100.0

    if nb is None or vol_l is None:
        m_combo = re.search(r"(\d+)\s*[x×]\s*(\d+(?:[.,]\d+)?)\s*([lc]l?)", s, flags=re.IGNORECASE)
        if m_combo:
            try:
                nb2 = int(m_combo.group(1))
                val = float(m_combo.group(2).replace(",", "."))
                unit = m_combo.group(3).lower()
                vol2 = val if unit.startswith("l") else val/100.0
                if nb is None: nb = nb2
                if vol_l is None: vol_l = vol2
            except:
                pass

    # Secours pour 4×75 cL
    if (nb is None or np.isnan(nb)) and vol_l is not None and abs(vol_l - 0.75) <= VOL_TOL:
        if re.search(r"(?:\b4\s*[x×]\b|Carton\s+de\s*4\b|4\s+Bouteilles?)", s, flags=re.IGNORECASE):
            nb = 4

    return (float(nb) if nb is not None else np.nan,
            float(vol_l) if vol_l is not None else np.nan)

def safe_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce")

def is_allowed_format(nb_bottles, vol_l, stock_txt: str) -> bool:
    if pd.isna(nb_bottles) or pd.isna(vol_l):
        if re.search(r"(?:\b4\s*[x×]\s*75\s*c?l\b|\b4\s+Bouteilles?\b.*75\s*c?l)", stock_txt, flags=re.IGNORECASE):
            nb_bottles = 4; vol_l = 0.75
        else:
            return False
    nb_bottles = int(nb_bottles); vol_l = float(vol_l)
    for nb_ok, vol_ok in ALLOWED_FORMATS:
        if nb_bottles == nb_ok and abs(vol_l - vol_ok) <= VOL_TOL:
            return True
    return False

# ---------- Coeur de calcul ----------
def compute_plan(df_in, volume_cible, nb_gouts, repartir_pro_rv, manual_keep, exclude_list, window_days):
    required = ["Produit", "Stock", "Quantité vendue", "Volume vendu (hl)", "Quantité disponible", "Volume disponible (hl)"]
    miss = [c for c in required if c not in df_in.columns]
    if miss:
        raise ValueError(f"Colonnes manquantes: {miss}")

    df = df_in[required].copy()
    for c in ["Quantité vendue", "Volume vendu (hl)", "Quantité disponible", "Volume disponible (hl)"]:
        df[c] = safe_num(df[c])

    # Parsing & filtre formats (après exclusions par cases et/ou manuel plus tard)
    parsed = df["Stock"].apply(parse_stock)
    df[["Bouteilles/carton", "Volume bouteille (L)"]] = pd.DataFrame(parsed.tolist(), index=df.index)
    mask_allowed = df.apply(lambda r: is_allowed_format(r["Bouteilles/carton"], r["Volume bouteille (L)"], str(r["Stock"])), axis=1)
    df = df.loc[mask_allowed].reset_index(drop=True)

    # Volumes/carton
    df["Volume/carton (hL)"] = (df["Bouteilles/carton"] * df["Volume bouteille (L)"]) / 100.0
    # Lignes valides
    df = df.dropna(subset=["Produit", "Volume/carton (hL)", "Volume vendu (hl)", "Volume disponible (hl)"]).reset_index(drop=True)

    df_all_formats = df.copy()  # copie avant sélection goûts

    # Exclusions via liste (cases à cocher)
    if exclude_list:
        df = df[~df["Produit"].astype(str).str.strip().isin(exclude_list)]
        df_all_formats = df_all_formats[~df_all_formats["Produit"].astype(str).str.strip().isin(exclude_list)]

    # Sélection manuelle optionnelle
    if manual_keep:
        keep = [g.strip() for g in manual_keep]
        df = df[df["Produit"].astype(str).str.strip().isin(keep)]
        df_all_formats = df_all_formats[df_all_formats["Produit"].astype(str).str.strip().isin(keep)]

    # ---------- Sélection intelligente des goûts ----------
    agg = df.groupby("Produit").agg(
        ventes_hl=("Volume vendu (hl)", "sum"),
        stock_hl=("Volume disponible (hl)", "sum")
    )
    agg["vitesse_j"] = agg["ventes_hl"] / max(float(window_days), 1.0)
    agg["jours_autonomie"] = np.where(agg["vitesse_j"] > 0, agg["stock_hl"] / agg["vitesse_j"], np.inf)
    agg["score_urgence"] = agg["vitesse_j"] / (agg["jours_autonomie"] + EPS)  # ≈ vitesse^2 / stock
    agg = agg.sort_values(by=["score_urgence", "jours_autonomie", "ventes_hl"], ascending=[False, True, False])

    if not manual_keep:
        gouts_cibles = agg.index.tolist()[:nb_gouts]
        df_selected = df[df["Produit"].isin(gouts_cibles)].copy()
    else:
        gouts_cibles = sorted(set(df["Produit"]))
        if len(gouts_cibles) > nb_gouts:
            order = [g for g in agg.index if g in gouts_cibles]
            gouts_cibles = order[:nb_gouts]
        df_selected = df[df["Produit"].isin(gouts_cibles)].copy()

    if len(gouts_cibles) == 0:
        raise ValueError("Aucun goût sélectionné (tout a peut-être été exclu).")

    # ---------- Calculs de production ----------
    df_calc = df_selected.copy()
    if nb_gouts == 1:
        df_calc["Somme ventes (hL) par goût"] = df_calc.groupby("Produit")["Volume vendu (hl)"].transform("sum")
        if repartir_pro_rv:
            df_calc["r_i"] = np.where(df_calc["Somme ventes (hL) par goût"] > 0,
                                      df_calc["Volume vendu (hl)"] / df_calc["Somme ventes (hL) par goût"], 0.0)
        else:
            df_calc["r_i"] = 1.0 / df_calc.groupby("Produit")["Produit"].transform("count")

        df_calc["G_i (hL)"] = df_calc["Volume disponible (hl)"]
        df_calc["G_total (hL) par goût"] = df_calc.groupby("Produit")["G_i (hL)"].transform("sum")
        df_calc["Y_total (hL) par goût"] = df_calc["G_total (hL) par goût"] + float(volume_cible)
        df_calc["X_th (hL)"] = df_calc["r_i"] * df_calc["Y_total (hL) par goût"] - df_calc["G_i (hL)"]

        df_calc["X_adj (hL)"] = 0.0
        for gout, grp in df_calc.groupby("Produit"):
            x = grp["X_th (hL)"].to_numpy(float)
            r = grp["r_i"].to_numpy(float)
            x = np.maximum(x, 0.0)
            deficit = float(volume_cible) - x.sum()
            if deficit > 1e-9:
                r = np.where(r > 0, r, 0); s = r.sum()
                x = x + (deficit * (r / s) if s > 0 else deficit / len(x))
            x = np.where(x < 1e-9, 0.0, x)
            df_calc.loc[grp.index, "X_adj (hL)"] = x

        cap_resume = f"{volume_cible:.2f} hL par goût"

    else:
        somme_ventes = df_calc["Volume vendu (hl)"].sum()
        if repartir_pro_rv and somme_ventes > 0:
            df_calc["r_i_global"] = df_calc["Volume vendu (hl)"] / somme_ventes
        else:
            df_calc["r_i_global"] = 1.0 / len(df_calc)

        df_calc["G_i (hL)"] = df_calc["Volume disponible (hl)"]
        G_total_all = df_calc["G_i (hL)"].sum()
        Y_total_all = G_total_all + float(volume_cible)
        df_calc["X_th (hL)"] = df_calc["r_i_global"] * Y_total_all - df_calc["G_i (hL)"]

        x = np.maximum(df_calc["X_th (hL)"].to_numpy(float), 0.0)
        deficit = float(volume_cible) - x.sum()
        if deficit > 1e-9:
            w = df_calc["r_i_global"].to_numpy(float); s = w.sum()
            x = x + (deficit * (w / s) if s > 0 else deficit / len(x))
        x = np.where(x < 1e-9, 0.0, x)
        df_calc["X_adj (hL)"] = x

        cap_resume = f"{volume_cible:.2f} hL au total (2 goûts)"

    # Cartons (exact + arrondi interne)
    df_calc["Cartons à produire (exact)"] = df_calc["X_adj (hL)"] / df_calc["Volume/carton (hL)"]
    if ROUND_TO_CARTON:
        df_calc["Cartons à produire (arrondi)"] = np.floor(df_calc["Cartons à produire (exact)"] + 0.5).astype("Int64")
        df_calc["Volume produit arrondi (hL)"] = df_calc["Cartons à produire (arrondi)"] * df_calc["Volume/carton (hL)"]

    # Sortie simplifiée
    df_min = df_calc[[
        "Produit", "Stock",
        "Cartons à produire (exact)", "Cartons à produire (arrondi)", "Volume produit arrondi (hL)"
    ]].sort_values(["Produit", "Stock"]).reset_index(drop=True)

    # Transparence sélection
    agg_full = df.groupby("Produit").agg(
        ventes_hl=("Volume vendu (hl)", "sum"),
        stock_hl=("Volume disponible (hl)", "sum")
    )
    agg_full["vitesse_j"] = agg_full["ventes_hl"] / max(float(window_days), 1.0)
    agg_full["jours_autonomie"] = np.where(agg_full["vitesse_j"] > 0, agg_full["stock_hl"] / agg_full["vitesse_j"], np.inf)
    agg_full["score_urgence"] = agg_full["vitesse_j"] / (agg_full["jours_autonomie"] + EPS)
    synth_sel = agg_full.loc[gouts_cibles][["ventes_hl", "stock_hl", "vitesse_j", "jours_autonomie", "score_urgence"]].copy()
    synth_sel = synth_sel.rename(columns={
        "ventes_hl": "Ventes 2 mois (hL)",
        "stock_hl": "Stock (hL)",
        "vitesse_j": "Vitesse (hL/j)",
        "jours_autonomie": "Autonomie (jours)",
        "score_urgence": "Score urgence"
    })

    return df_min, cap_resume, gouts_cibles, synth_sel, df_calc, df_all_formats

# ---------- Lecture + UI dynamique (exclusions/manuel) ----------
if uploaded is None:
    st.info("💡 Charge un fichier Excel pour commencer.")
    st.stop()

try:
    df_in = read_input_excel(uploaded)
except Exception as e:
    st.error(f"Erreur de lecture : {e}")
    st.stop()

# UI dynamique : exclusions par cases + sélection manuelle optionnelle
with st.sidebar:
    all_gouts = sorted(pd.Series(df_in.get("Produit", pd.Series(dtype=str))).dropna().astype(str).unique())
    excluded_gouts = st.multiselect("🚫 Exclure certains goûts", options=all_gouts, default=[])

    use_manual = st.checkbox("Sélection manuelle DES goûts à produire", value=False, help="Sinon : sélection automatique (autonomie + ventes).")
    manual_keep = None
    if use_manual:
        manual_keep = st.multiselect("Choisis les goûts à produire", options=[g for g in all_gouts if g not in excluded_gouts], default=[])

# ---------- Calcul principal ----------
try:
    df_min, cap_resume, gouts_cibles, synth_sel, df_selected_calc, df_all_formats = compute_plan(
        df_in=df_in,
        volume_cible=volume_cible,
        nb_gouts=nb_gouts,
        repartir_pro_rv=repartir_pro_rv,
        manual_keep=manual_keep,
        exclude_list=excluded_gouts,
        window_days=window_days
    )
except Exception as e:
    st.error(f"Erreur de calcul : {e}")
    st.stop()

# ---------- Estimation des pertes de CA (avec PRIX SAISIS) ----------
df_all = df_all_formats.copy()
df_all["vitesse_hL_j"] = df_all["Volume vendu (hl)"] / max(float(window_days), 1.0)

# Prix par hL selon prix saisis
def revenue_per_hL(vol_bottle_L: float) -> float:
    if pd.isna(vol_bottle_L): return 0.0
    if abs(vol_bottle_L - 0.33) <= VOL_TOL:
        price = price_033
        vol_key = 0.33
    elif abs(vol_bottle_L - 0.75) <= VOL_TOL:
        price = price_075
        vol_key = 0.75
    else:
        return 0.0
    bottles_per_hL = 100.0 / vol_key  # nb bouteilles par hL
    return bottles_per_hL * price

df_all["€_par_hL"] = df_all["Volume bouteille (L)"].apply(revenue_per_hL)
df_all["€_par_j"] = df_all["vitesse_hL_j"] * df_all["€_par_hL"]

# Horizon T_end = date d'épuisement commune des goûts sélectionnés
df_sel = df_selected_calc.copy()
df_sel["vitesse_hL_j"] = df_sel["Volume vendu (hl)"] / max(float(window_days), 1.0)
total_stock_plus_prod = (df_sel["Volume disponible (hl)"] + df_sel.get("X_adj (hL)", 0)).sum()
total_speed = df_sel["vitesse_hL_j"].sum()
T_end = np.inf if total_speed <= EPS else total_stock_plus_prod / total_speed

# Pertes sur les goûts NON sélectionnés (jusqu'à T_end)
df_non_sel = df_all[~df_all["Produit"].isin(gouts_cibles)].copy()
if np.isinf(T_end) or T_end <= 0:
    df_non_sel["Perte (€)"] = 0.0
else:
    df_non_sel["t_rup_j"] = np.where(df_non_sel["vitesse_hL_j"] > 0,
                                     df_non_sel["Volume disponible (hl)"] / df_non_sel["vitesse_hL_j"],
                                     np.inf)
    df_non_sel["jours_perdus"] = np.clip(T_end - df_non_sel["t_rup_j"], a_min=0.0, a_max=None)
    df_non_sel["Perte (€)"] = df_non_sel["jours_perdus"] * df_non_sel["€_par_j"]

pertes_par_gout = df_non_sel.groupby("Produit", as_index=False)["Perte (€)"].sum().sort_values("Perte (€)", ascending=False)
perte_totale = float(pertes_par_gout["Perte (€)"].sum()) if len(pertes_par_gout) else 0.0

# ---------- Affichages ----------
st.subheader("Résumé")
st.metric("Goûts sélectionnés", len(gouts_cibles))
st.metric("Capacité utilisée", cap_resume)

st.subheader("Production simplifiée")
st.dataframe(df_min.head(200), use_container_width=True)

with st.expander("Pourquoi ces goûts ? (autonomie & ventes)"):
    st.dataframe(
        synth_sel.style.format({
            "Ventes 2 mois (hL)": "{:.2f}",
            "Stock (hL)": "{:.2f}",
            "Vitesse (hL/j)": "{:.3f}",
            "Autonomie (jours)": lambda v: "∞" if np.isinf(v) else f"{v:.1f}",
            "Score urgence": "{:.6f}",
        }),
        use_container_width=True
    )

with st.expander("💶 Impact CA — pertes estimées sur l’horizon de production"):
    st.write(f"**Horizon d'évaluation (T_end)** ≈ {('∞' if np.isinf(T_end) else f'{T_end:.1f} jours')} (jusqu'à épuisement des goûts sélectionnés).")
    col1, col2 = st.columns([2,1])
    with col1:
        if len(pertes_par_gout):
            st.dataframe(pertes_par_gout.style.format({"Perte (€)": "€{:,.0f}"}), use_container_width=True)
        else:
            st.info("Aucune perte estimée (pas de goût non sélectionné en rupture sur l'horizon).")
    with col2:
        st.metric("Perte totale estimée", f"€{perte_totale:,.0f}")

    st.caption("Méthode : T_end = (stocks + production des goûts sélectionnés) / vitesse de vente des goûts sélectionnés. "
               "Perte d’un goût non sélectionné = max(T_end - temps jusqu'à rupture, 0) × CA/jour. "
               "CA/jour = (Volume vendu/jour) × (bouteilles/hL) × prix bouteille (saisis ci-dessus).")
