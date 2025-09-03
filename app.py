import io
import re
from pathlib import Path
import difflib
import numpy as np
import pandas as pd
import streamlit as st

# =========================
# Optimiseur de production (v4.8)
# - Fenêtre (jours) lue en B2 du xlsx
# - Mapping produits → goûts canoniques (CSV repo ou upload)
# - Sélection intelligente (autonomie + ventes) par GoutCanon
# - 1 goût : 64 hL PAR goût ; 2 goûts : 64 hL AU TOTAL
# - Formats: 12×0.33 L, 6×0.75 L, 4×0.75 L
# - Ignore les lignes au fond noir
# - Prix bouteille éditables (0.33 L / 0.75 L)
# - Pertes CA (vue principale: aucune production pour tous les goûts)
# =========================

st.set_page_config(page_title="Optimiseur de production — 64 hL / 1–2 goûts", page_icon="🧪", layout="wide")

ALLOWED_FORMATS = {(12, 0.33), (6, 0.75), (4, 0.75)}
ROUND_TO_CARTON = True
VOL_TOL = 0.02
EPS = 1e-9
DEFAULT_WINDOW_DAYS = 60  # fallback si B2 introuvable


# ---------- Sidebar (paramètres de prod + prix + mapping) ----------
with st.sidebar:
    st.header("Paramètres de production")
    volume_cible = st.number_input(
        "Volume cible (hL)",
        min_value=1.0, value=64.0, step=1.0,
        help="Si 1 goût: volume PAR goût. Si 2 goûts: volume TOTAL partagé."
    )
    nb_gouts = st.selectbox("Nombre de goûts simultanés", [1, 2], index=0)
    repartir_pro_rv = st.checkbox(
        "Répartition par formats au prorata des vitesses de vente",
        value=True,
        help="Sinon: répartition égale entre formats d'un même goût."
    )
    st.markdown("---")
    st.subheader("Prix par bouteille (€)")
    price_033 = st.number_input("Prix 0,33 L (€ / bt)", min_value=0.0, value=1.75, step=0.01, format="%.2f")
    price_075 = st.number_input("Prix 0,75 L (€ / bt)", min_value=0.0, value=3.10, step=0.01, format="%.2f")
    st.markdown("---")
    st.subheader("Mapping des goûts")
    uploaded_map = st.file_uploader("Uploader un `flavor_map.csv` (optionnel)", type=["csv"], key="map")

st.title("🧪 Optimiseur de production — 64 hL / 1–2 goûts")
st.caption("Fenêtre (jours) lue automatiquement en **B2**. Les calculs s’effectuent par **goût canonique** (mapping CSV).")

# ---------- Upload Excel ----------
uploaded = st.file_uploader("Dépose ton fichier Excel (.xlsx/.xls)", type=["xlsx", "xls"])
if uploaded is None:
    st.info("💡 Charge un fichier Excel pour commencer.")
    st.stop()


# ---------- Outils Excel / parsing ----------
def detect_header_row(df_raw: pd.DataFrame) -> int:
    must = {"Produit", "Stock", "Quantité vendue", "Volume vendu (hl)", "Quantité disponible", "Volume disponible (hl)"}
    for i in range(min(10, len(df_raw))):
        if must.issubset(set(str(x).strip() for x in df_raw.iloc[i].tolist())):
            return i
    return 0

def rows_to_keep_by_fill(excel_bytes: bytes, header_idx: int) -> list[bool]:
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
    ws = wb[wb.sheetnames[0]]
    start_row = header_idx + 2
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

def parse_days_from_b2(value) -> int | None:
    try:
        if isinstance(value, (int, float)) and not pd.isna(value):
            v = int(round(float(value)))
            return v if v > 0 else None
        if value is None:
            return None
        s = str(value).strip()
        m = re.search(r"(\d+)\s*(?:j|jour|jours)\b", s, flags=re.IGNORECASE)
        if m:
            v = int(m.group(1))
            return v if v > 0 else None
        date_pat = r"(\d{1,2}[/-]\d{1,2}[/-]\d{2,4}).*?(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})"
        m2 = re.search(date_pat, s)
        if m2:
            d1 = pd.to_datetime(m2.group(1), dayfirst=True, errors="coerce")
            d2 = pd.to_datetime(m2.group(2), dayfirst=True, errors="coerce")
            if pd.notna(d1) and pd.notna(d2):
                days = int((d2 - d1).days)
                return days if days > 0 else None
        m3 = re.search(r"\b(\d{1,4})\b", s)
        if m3:
            v = int(m3.group(1))
            return v if v > 0 else None
    except Exception:
        return None
    return None

def read_input_excel_and_period(uploaded_file):
    file_bytes = uploaded_file.read()
    raw = pd.read_excel(io.BytesIO(file_bytes), header=None)
    header_idx = detect_header_row(raw)
    df = pd.read_excel(io.BytesIO(file_bytes), header=header_idx)
    keep_mask = rows_to_keep_by_fill(file_bytes, header_idx)
    if len(keep_mask) < len(df):
        keep_mask = keep_mask + [True] * (len(df) - len(keep_mask))
    df = df.iloc[[i for i, k in enumerate(keep_mask) if k]].reset_index(drop=True)
    # lecture de B2
    try:
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        ws = wb[wb.sheetnames[0]]
        b2_val = ws["B2"].value
        wd = parse_days_from_b2(b2_val)
    except Exception:
        wd = None
    return df, (wd if wd and wd > 0 else DEFAULT_WINDOW_DAYS), file_bytes


# ---------- Mapping produits → goûts canoniques ----------
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
    out["Produit"] = out["Produit"].astype(str)
    out["Produit_norm"] = out["Produit"].str.strip()
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


# ---------- Parsing "Stock" et formats ----------
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
            try: nb = int(m.group(1)); break
            except: pass

    vol_l = None
    m_l = re.findall(r"(\d+(?:[.,]\d+)?)\s*[lL]", s)
    if m_l: vol_l = float(m_l[-1].replace(",", "."))
    else:
        m_cl = re.findall(r"(\d+(?:[.,]\d+)?)\s*c[lL]", s)
        if m_cl: vol_l = float(m_cl[-1].replace(",", ".")) / 100.0

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
            except: pass

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


# ---------- Calcul principal (tout en GoutCanon) ----------
def compute_plan(df_in, window_days, volume_cible, nb_gouts, repartir_pro_rv, manual_keep, exclude_list):
    required = ["Produit", "GoutCanon", "Stock", "Quantité vendue", "Volume vendu (hl)", "Quantité disponible", "Volume disponible (hl)"]
    miss = [c for c in required if c not in df_in.columns]
    if miss: raise ValueError(f"Colonnes manquantes: {miss}")

    df = df_in[required].copy()
    for c in ["Quantité vendue", "Volume vendu (hl)", "Quantité disponible", "Volume disponible (hl)"]:
        df[c] = safe_num(df[c])

    parsed = df["Stock"].apply(parse_stock)
    df[["Bouteilles/carton", "Volume bouteille (L)"]] = pd.DataFrame(parsed.tolist(), index=df.index)
    mask_allowed = df.apply(lambda r: is_allowed_format(r["Bouteilles/carton"], r["Volume bouteille (L)"], str(r["Stock"])), axis=1)
    df = df.loc[mask_allowed].reset_index(drop=True)

    df["Volume/carton (hL)"] = (df["Bouteilles/carton"] * df["Volume bouteille (L)"]) / 100.0
    df = df.dropna(subset=["GoutCanon", "Volume/carton (hL)", "Volume vendu (hl)", "Volume disponible (hl)"]).reset_index(drop=True)

    df_all_formats = df.copy()

    # Exclusions par goût canonique
    if exclude_list:
        df = df[~df["GoutCanon"].astype(str).str.strip().isin(exclude_list)]
        df_all_formats = df_all_formats[~df_all_formats["GoutCanon"].astype(str).str.strip().isin(exclude_list)]

    # Sélection manuelle (par goût canonique)
    if manual_keep:
        keep = [g.strip() for g in manual_keep]
        df = df[df["GoutCanon"].astype(str).str.strip().isin(keep)]
        df_all_formats = df_all_formats[df_all_formats["GoutCanon"].astype(str).str.strip().isin(keep)]

    # Sélection intelligente (autonomie + ventes), par goût canonique
    agg = df.groupby("GoutCanon").agg(
        ventes_hl=("Volume vendu (hl)", "sum"),
        stock_hl=("Volume disponible (hl)", "sum")
    )
    agg["vitesse_j"] = agg["ventes_hl"] / max(float(window_days), 1.0)
    agg["jours_autonomie"] = np.where(agg["vitesse_j"] > 0, agg["stock_hl"] / agg["vitesse_j"], np.inf)
    agg["score_urgence"] = agg["vitesse_j"] / (agg["jours_autonomie"] + EPS)
    agg = agg.sort_values(by=["score_urgence", "jours_autonomie", "ventes_hl"], ascending=[False, True, False])

    if not manual_keep:
        gouts_cibles = agg.index.tolist()[:nb_gouts]
        df_selected = df[df["GoutCanon"].isin(gouts_cibles)].copy()
    else:
        gouts_cibles = sorted(set(df["GoutCanon"]))
        if len(gouts_cibles) > nb_gouts:
            order = [g for g in agg.index if g in gouts_cibles]
            gouts_cibles = order[:nb_gouts]
        df_selected = df[df["GoutCanon"].isin(gouts_cibles)].copy()

    if len(gouts_cibles) == 0:
        raise ValueError("Aucun goût sélectionné (tout a peut-être été exclu).")

    # Calculs de production
    df_calc = df_selected.copy()
    if nb_gouts == 1:
        df_calc["Somme ventes (hL) par goût"] = df_calc.groupby("GoutCanon")["Volume vendu (hl)"].transform("sum")
        if repartir_pro_rv:
            df_calc["r_i"] = np.where(df_calc["Somme ventes (hL) par goût"] > 0,
                                      df_calc["Volume vendu (hl)"] / df_calc["Somme ventes (hL) par goût"], 0.0)
        else:
            df_calc["r_i"] = 1.0 / df_calc.groupby("GoutCanon")["GoutCanon"].transform("count")

        df_calc["G_i (hL)"] = df_calc["Volume disponible (hl)"]
        df_calc["G_total (hL) par goût"] = df_calc.groupby("GoutCanon")["G_i (hL)"].transform("sum")
        df_calc["Y_total (hL) par goût"] = df_calc["G_total (hL) par goût"] + float(volume_cible)
        df_calc["X_th (hL)"] = df_calc["r_i"] * df_calc["Y_total (hL) par goût"] - df_calc["G_i (hL)"]

        df_calc["X_adj (hL)"] = 0.0
        for gout, grp in df_calc.groupby("GoutCanon"):
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

    # Cartons
    df_calc["Cartons à produire (exact)"] = df_calc["X_adj (hL)"] / df_calc["Volume/carton (hL)"]
    if ROUND_TO_CARTON:
        df_calc["Cartons à produire (arrondi)"] = np.floor(df_calc["Cartons à produire (exact)"] + 0.5).astype("Int64")
        df_calc["Volume produit arrondi (hL)"] = df_calc["Cartons à produire (arrondi)"] * df_calc["Volume/carton (hL)"]

    # Sortie simplifiée (on expose Produit ET GoutCanon)
    df_min = df_calc[[
        "GoutCanon", "Produit", "Stock",
        "Cartons à produire (exact)", "Cartons à produire (arrondi)", "Volume produit arrondi (hL)"
    ]].sort_values(["GoutCanon", "Produit", "Stock"]).reset_index(drop=True)

    # Transparence sélection
    agg_full = df.groupby("GoutCanon").agg(
        ventes_hl=("Volume vendu (hl)", "sum"),
        stock_hl=("Volume disponible (hl)", "sum")
    )
    agg_full["vitesse_j"] = agg_full["ventes_hl"] / max(float(window_days), 1.0)
    agg_full["jours_autonomie"] = np.where(agg_full["vitesse_j"] > 0, agg_full["stock_hl"] / agg_full["vitesse_j"], np.inf)
    agg_full["score_urgence"] = agg_full["vitesse_j"] / (agg_full["jours_autonomie"] + EPS)
    sel_gouts = sorted(set(df_calc["GoutCanon"]))
    synth_sel = agg_full.loc[sel_gouts][["ventes_hl", "stock_hl", "vitesse_j", "jours_autonomie", "score_urgence"]].copy()
    synth_sel = synth_sel.rename(columns={
        "ventes_hl": "Ventes 2 mois (hL)",
        "stock_hl": "Stock (hL)",
        "vitesse_j": "Vitesse (hL/j)",
        "jours_autonomie": "Autonomie (jours)",
        "score_urgence": "Score urgence"
    })

    return df_min, cap_resume, sel_gouts, synth_sel, df_calc, df_all_formats, agg_full


# ---------- Lecture + période ----------
try:
    df_in_raw, window_days, file_bytes = read_input_excel_and_period(uploaded)
except Exception as e:
    st.error(f"Erreur de lecture : {e}")
    st.stop()

st.info(f"📅 Fenêtre détectée (B2) : **{window_days} jours** (défaut {DEFAULT_WINDOW_DAYS} si non détecté).")

# Appliquer le mapping produits → goûts
flavor_map_df = load_flavor_map(uploaded_map)
df_in = apply_canonical_flavor(df_in_raw, flavor_map_df)

with st.expander("🔎 Produits non mappés (suggestions)"):
    missing = suggest_missing_mappings(df_in, flavor_map_df)
    if len(missing):
        st.write("Ces libellés ne sont pas encore dans `flavor_map.csv` :")
        st.dataframe(missing, use_container_width=True)
        st.caption("Ajoute les lignes correspondantes dans `flavor_map.csv` (colonne `name` = libellé Excel, `canonical` = goût).")
    else:
        st.success("Tous les libellés sont couverts par le mapping.")

# ---------- UI dynamique : exclusions / manuel (par goût canonique) ----------
with st.sidebar:
    all_gouts = sorted(pd.Series(df_in.get("GoutCanon", pd.Series(dtype=str))).dropna().astype(str).unique())
    excluded_gouts = st.multiselect("🚫 Exclure certains goûts (canoniques)", options=all_gouts, default=[])
    use_manual = st.checkbox("Sélection manuelle DES goûts à produire", value=False)
    manual_keep = None
    if use_manual:
        manual_keep = st.multiselect("Choisis les goûts à produire", options=[g for g in all_gouts if g not in excluded_gouts], default=[])

# ---------- Calcul principal ----------
try:
    df_min, cap_resume, gouts_cibles, synth_sel, df_selected_calc, df_all_formats, agg_full = compute_plan(
        df_in=df_in,
        window_days=window_days,
        volume_cible=volume_cible,
        nb_gouts=nb_gouts,
        repartir_pro_rv=repartir_pro_rv,
        manual_keep=manual_keep,
        exclude_list=excluded_gouts
    )
except Exception as e:
    st.error(f"Erreur de calcul : {e}")
    st.stop()

# ---------- Prix & vitesses (pour pertes) ----------
df_all = df_all_formats.copy()
df_all["vitesse_hL_j"] = df_all["Volume vendu (hl)"] / max(float(window_days), 1.0)

def revenue_per_hL(vol_bottle_L: float) -> float:
    if pd.isna(vol_bottle_L): return 0.0
    if abs(vol_bottle_L - 0.33) <= VOL_TOL:
        price = price_033; vol_key = 0.33
    elif abs(vol_bottle_L - 0.75) <= VOL_TOL:
        price = price_075; vol_key = 0.75
    else:
        return 0.0
    bottles_per_hL = 100.0 / vol_key
    return bottles_per_hL * price

df_all["€_par_hL"] = df_all["Volume bouteille (L)"].apply(revenue_per_hL)
df_all["€_par_j"] = df_all["vitesse_hL_j"] * df_all["€_par_hL"]

# Vue PRINCIPALE : pertes si on NE PRODUIT RIEN (tous les goûts canoniques)
agg_all = df_all.groupby("GoutCanon").agg(
    vitesse_hL_j=("vitesse_hL_j", "sum"),
    stock_hL=("Volume disponible (hl)", "sum"),
    euro_par_j=("€_par_j", "sum"),
).reset_index()
agg_all["autonomie_j"] = np.where(agg_all["vitesse_hL_j"] > 0, agg_all["stock_hL"] / agg_all["vitesse_hL_j"], np.inf)
agg_all["jours_perdus"] = np.clip(float(window_days) - agg_all["autonomie_j"], a_min=0.0, a_max=None)
agg_all["Perte (€)"] = agg_all["jours_perdus"] * agg_all["euro_par_j"]
pertes_tous_aucune_prod = agg_all[["GoutCanon", "autonomie_j", "euro_par_j", "Perte (€)"]].copy()
pertes_tous_aucune_prod = pertes_tous_aucune_prod.sort_values("Perte (€)", ascending=False)
pertes_tous_aucune_prod.rename(columns={
    "GoutCanon": "Goût",
    "autonomie_j": "Autonomie (jours)",
    "euro_par_j": "CA/jour (€)",
}, inplace=True)
perte_totale_aucune = float(pertes_tous_aucune_prod["Perte (€)"].sum()) if len(pertes_tous_aucune_prod) else 0.0

# Optionnel : pertes des non-sélectionnés jusqu’à T_end (par canoniques)
df_sel = df_selected_calc.copy()
df_sel["vitesse_hL_j"] = df_sel["Volume vendu (hl)"] / max(float(window_days), 1.0)
total_stock_plus_prod = (df_sel["Volume disponible (hl)"] + df_sel.get("X_adj (hL)", 0)).sum()
total_speed = df_sel["vitesse_hL_j"].sum()
T_end = np.inf if total_speed <= EPS else total_stock_plus_prod / total_speed

df_non_sel = df_all[~df_all["GoutCanon"].isin(gouts_cibles)].copy()
if np.isinf(T_end) or T_end <= 0:
    df_non_sel["Perte (€)"] = 0.0
else:
    df_non_sel["t_rup_j"] = np.where(df_non_sel["vitesse_hL_j"] > 0,
                                     df_non_sel["Volume disponible (hl)"] / df_non_sel["vitesse_hL_j"],
                                     np.inf)
    df_non_sel["jours_perdus"] = np.clip(T_end - df_non_sel["t_rup_j"], a_min=0.0, a_max=None)
    df_non_sel["Perte (€)"] = df_non_sel["jours_perdus"] * df_non_sel["€_par_j"]
pertes_non_sel_Tend = df_non_sel.groupby("GoutCanon", as_index=False)["Perte (€)"].sum().rename(columns={"GoutCanon": "Goût"})
pertes_non_sel_Tend = pertes_non_sel_Tend.sort_values("Perte (€)", ascending=False)
perte_totale_Tend = float(pertes_non_sel_Tend["Perte (€)"].sum()) if len(pertes_non_sel_Tend) else 0.0

# ---------- Affichages ----------
st.subheader("Résumé")
st.metric("Goûts sélectionnés", len(gouts_cibles))
st.metric("Capacité utilisée", cap_resume)
st.caption(f"Fenêtre utilisée (B2) : **{window_days} jours**.")

st.subheader("Production simplifiée (par formats)")
st.dataframe(df_min.head(300), use_container_width=True)

with st.expander("Pourquoi ces goûts ? (autonomie & ventes — par goût canonique)"):
    st.dataframe(
        synth_sel.style.format({
            "Ventes 2 mois (hL)": "{:.2f}",
            "Stock (hL)": "{:.2f}",
            "Vitesse (hL/j)": "{:.3f}",
            "Autonomie (jours)": lambda v: '∞' if np.isinf(v) else f"{v:.1f}",
            "Score urgence": "{:.6f}",
        }),
        use_container_width=True
    )

st.subheader("💶 Pertes si on NE PRODUIT RIEN (tous les goûts canoniques)")
colA, colB = st.columns([2,1])
with colA:
    st.dataframe(
        pertes_tous_aucune_prod.style.format({
            "Autonomie (jours)": lambda v: '∞' if np.isinf(v) else f"{v:.1f}",
            "CA/jour (€)": "€{:,.0f}",
            "Perte (€)": "€{:,.0f}",
        }),
        use_container_width=True
    )
with colB:
    st.metric("Perte totale estimée (aucune production)", f"€{perte_totale_aucune:,.0f}")

with st.expander("(Optionnel) Pertes des goûts NON sélectionnés jusqu'à T_end"):
    st.write(f"**Horizon T_end** ≈ {('∞' if np.isinf(T_end) else f'{T_end:.1f} jours')} (épuisement des goûts sélectionnés).")
    col1, col2 = st.columns([2,1])
    with col1:
        if len(pertes_non_sel_Tend):
            st.dataframe(pertes_non_sel_Tend.style.format({"Perte (€)": "€{:,.0f}"}), use_container_width=True)
        else:
            st.info("Aucune perte estimée (pas de goût non sélectionné en rupture sur l'horizon).")
    with col2:
        st.metric("Perte totale (jusqu'à T_end)", f"€{perte_totale_Tend:,.0f}")
