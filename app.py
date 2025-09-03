import io
import re
import numpy as np
import pandas as pd
import streamlit as st

# =========================
# Optimiseur de production (v4.1)
# - Sélection intelligente des goûts: score = vitesse_de_vente / (jours_autonomie + eps)
# - 1 goût : 64 hL PAR goût, 2 goûts : 64 hL AU TOTAL (épuisement simultané)
# - Formats internes: 12×0.33 L, 6×0.75 L, 4×0.75 L (parseur robuste)
# - Arrondi au carton (half-up)
# - Lecture Excel: ignore les lignes contenant au moins UNE cellule au fond noir
# =========================

st.set_page_config(page_title="Optimiseur de production — 64 hL / 1–2 goûts", page_icon="🧪", layout="wide")

ALLOWED_FORMATS = {(12, 0.33), (6, 0.75), (4, 0.75)}
ROUND_TO_CARTON = True
VOL_TOL = 0.02   # tolérance sur 0.33 / 0.75 (L)
EPS = 1e-9

# ---------- Sidebar ----------
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
    st.subheader("Contraintes goûts (optionnel)")
    use_manual = st.checkbox("Sélection manuelle des goûts", value=False)
    gouts_exclus = st.text_input("Exclure goûts (séparés par des virgules)", value="")

# ---------- Header ----------
st.title("🧪 Optimiseur de production — 64 hL / 1–2 goûts")
st.caption("Sélection automatique des goûts (autonomie + ventes), calcul par formats pour écoulement simultané des stocks.")

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

    # Secours pour 4×75 cL : si on voit 0.75 L et un motif "4x / ×4 / Carton de 4 / 4 Bouteilles"
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

    # Parsing & filtre formats
    parsed = df["Stock"].apply(parse_stock)
    df[["Bouteilles/carton", "Volume bouteille (L)"]] = pd.DataFrame(parsed.tolist(), index=df.index)
    mask_allowed = df.apply(lambda r: is_allowed_format(r["Bouteilles/carton"], r["Volume bouteille (L)"], str(r["Stock"])), axis=1)
    df = df.loc[mask_allowed].reset_index(drop=True)

    df["Volume/carton (hL)"] = (df["Bouteilles/carton"] * df["Volume bouteille (L)"]) / 100.0
    df = df.dropna(subset=["Produit", "Volume/carton (hL)", "Volume vendu (hl)", "Volume disponible (hl)"]).reset_index(drop=True)

    # Exclusions / manuel
    if exclude_list:
        excl = [g.strip() for g in exclude_list]
        df = df[~df["Produit"].astype(str).str.strip().isin(excl)]
    if manual_keep:
        keep = [g.strip() for g in manual_keep]
        df = df[df["Produit"].astype(str).str.strip().isin(keep)]

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
        df = df[df["Produit"].isin(gouts_cibles)]
    else:
        gouts_cibles = sorted(set(df["Produit"]))
        if len(gouts_cibles) > nb_gouts:
            order = [g for g in agg.index if g in gouts_cibles]
            gouts_cibles = order[:nb_gouts]
            df = df[df["Produit"].isin(gouts_cibles)]

    if len(gouts_cibles) == 0:
        raise ValueError("Aucun goût sélectionné.")

    # ---------- Calculs de production ----------
    if nb_gouts == 1:
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
            x = grp["X_th (hL)"].to_numpy(float)
            r = grp["r_i"].to_numpy(float)
            x = np.maximum(x, 0.0)
            deficit = float(volume_cible) - x.sum()
            if deficit > 1e-9:
                r = np.where(r > 0, r, 0); s = r.sum()
                x = x + (deficit * (r / s) if s > 0 else deficit / len(x))
            x = np.where(x < 1e-9, 0.0, x)
            df.loc[grp.index, "X_adj (hL)"] = x

        cap_resume = f"{volume_cible:.2f} hL par goût"

    else:
        somme_ventes = df["Volume vendu (hl)"].sum()
        if repartir_pro_rv and somme_ventes > 0:
            df["r_i_global"] = df["Volume vendu (hl)"] / somme_ventes
        else:
            df["r_i_global"] = 1.0 / len(df)

        df["G_i (hL)"] = df["Volume disponible (hl)"]
        G_total_all = df["G_i (hL)"].sum()
        Y_total_all = G_total_all + float(volume_cible)
        df["X_th (hL)"] = df["r_i_global"] * Y_total_all - df["G_i (hL)"]

        x = np.maximum(df["X_th (hL)"].to_numpy(float), 0.0)
        deficit = float(volume_cible) - x.sum()
        if deficit > 1e-9:
            w = df["r_i_global"].to_numpy(float); s = w.sum()
            x = x + (deficit * (w / s) if s > 0 else deficit / len(x))
        x = np.where(x < 1e-9, 0.0, x)
        df["X_adj (hL)"] = x

        cap_resume = f"{volume_cible:.2f} hL au total (2 goûts)"

    # Cartons (exact + arrondi interne)
    df["Cartons à produire (exact)"] = df["X_adj (hL)"] / df["Volume/carton (hL)"]
    if ROUND_TO_CARTON:
        df["Cartons à produire (arrondi)"] = np.floor(df["Cartons à produire (exact)"] + 0.5).astype("Int64")
        df["Volume produit arrondi (hL)"] = df["Cartons à produire (arrondi)"] * df["Volume/carton (hL)"]

    # Sortie simplifiée
    df_min = df[[
        "Produit", "Stock",
        "Cartons à produire (exact)", "Cartons à produire (arrondi)", "Volume produit arrondi (hL)"
    ]].sort_values(["Produit", "Stock"]).reset_index(drop=True)

    # Transparence sélection
    synth_sel = agg.loc[gouts_cibles][["ventes_hl", "stock_hl", "vitesse_j", "jours_autonomie", "score_urgence"]].copy()
    synth_sel = synth_sel.rename(columns={
        "ventes_hl": "Ventes 2 mois (hL)",
        "stock_hl": "Stock (hL)",
        "vitesse_j": "Vitesse (hL/j)",
        "jours_autonomie": "Autonomie (jours)",
        "score_urgence": "Score urgence"
    })

    return df_min, cap_resume, gouts_cibles, synth_sel

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
    df_min, cap_resume, gouts_cibles, synth_sel = compute_plan(
        df_in, volume_cible, nb_gouts, repartir_pro_rv, manual_keep, exclude_list, window_days
    )
except Exception as e:
    st.error(f"Erreur de calcul : {e}")
    st.stop()

# ---------- Display ----------
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
