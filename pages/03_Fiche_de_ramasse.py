# pages/03_Fiche_de_ramasse.py
from __future__ import annotations
import math, re, os, io, datetime as dt
import pandas as pd
import streamlit as st
from dateutil.tz import gettz

from common.design import apply_theme, section, kpi
# ---- Chemin du PDF de référence (dans le repo) ----
REF_PDF_PATH = "assets/LOG_EN_001_01 BL enlèvements Sofripa-2.pdf"

# ===== Constantes de calcul =====
# Règles palettes (usuel logistique) — utilisées pour "Quantité palettes"
PALLETS_RULES = {
    "12x33": 108,  # cartons/palette
    "6x75":   84,
    "4x75":  100,
}
# Bouteilles par carton (pour calcul interne)
BTL_PER_CARTON = {"12x33": 12, "6x75": 6, "4x75": 4}

def _tz_today_paris() -> dt.date:
    return dt.datetime.now(gettz("Europe/Paris")).date()

def _norm_text(s: str) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip()

def _format_tag(s: str) -> str | None:
    """Détecte '12x33', '6x75', '4x75' quand ils sont écrits explicitement."""
    if not s:
        return None
    t = str(s).replace("×", "x").lower()
    # formes "12x33cl", "12 x 33 cl", etc.
    m = re.search(r"\b(12|6|4)\s*x\s*(33|75)\s*c?l?\b", t)
    if m:
        nb, vol = int(m.group(1)), int(m.group(2))
        if nb == 12 and vol == 33: return "12x33"
        if nb == 6 and vol == 75:  return "6x75"
        if nb == 4 and vol == 75:  return "4x75"
    # variantes très tolérantes (sans 'x' mais collées)
    if re.search(r"\b12\s*[*x]?\s*33\b", t): return "12x33"
    if re.search(r"\b6\s*[*x]?\s*75\b",  t): return "6x75"
    if re.search(r"\b4\s*[*x]?\s*75\b",  t): return "4x75"
    return None

def _format_from_stock(stock_txt: str) -> str | None:
    """
    Détecte 12x33 / 6x75 / 4x75 depuis 'Stock' très libre :
    - 0.33L / 0,33 L / 0.75L / 0,75 L
    - 33cl / 75cl (avec/sans espace)
    - tolère mots/tirets entre nb bouteilles et volume
    """
    if not stock_txt:
        return None

    s = str(stock_txt)
    s_low = s.lower().replace("×", "x")

    # 1) '12x33', '6x75' explicites
    m = re.search(r"\b(12|6|4)\s*x\s*(33|75)\b", s_low)
    if m:
        nb, vol = int(m.group(1)), int(m.group(2))
        if nb == 12 and vol == 33: return "12x33"
        if nb == 6  and vol == 75: return "6x75"
        if nb == 4  and vol == 75: return "4x75"

    # 2) Litres : "... 12 ... 0.33L"
    m_l = re.search(r"\b(\d+)\b.*?\b(0[.,]\d+)\s*l\b", s_low, flags=re.IGNORECASE)
    if m_l:
        try:
            nb = int(m_l.group(1))
            vol_l = float(m_l.group(2).replace(",", "."))
            vol_cl = int(round(vol_l * 100))   # 0.33 -> 33 ; 0.75 -> 75
            if nb == 12 and vol_cl == 33: return "12x33"
            if nb == 6  and vol_cl == 75: return "6x75"
            if nb == 4  and vol_cl == 75: return "4x75"
        except Exception:
            pass

    # 3) Centilitres : "... 12 ... 33cl"
    m_cl = re.search(r"\b(\d+)\b.*?\b(\d+)\s*c?l\b", s_low, flags=re.IGNORECASE)
    if m_cl:
        try:
            nb = int(m_cl.group(1))
            vol_cl = int(m_cl.group(2))
            if nb == 12 and vol_cl == 33: return "12x33"
            if nb == 6  and vol_cl == 75: return "6x75"
            if nb == 4  and vol_cl == 75: return "4x75"
        except Exception:
            pass

    return None


# ====== Extraction PDF → mapping (Produit -> {format, ref, poids_carton}) ======
def _parse_reference_pdf(pdf_path: str) -> pd.DataFrame:
    """
    Lit le PDF (tableau Sofripa) et retourne un DataFrame avec colonnes :
    - gout (normalisé)
    - format ('12x33' / '6x75' / '4x75')
    - reference (numéro entre parenthèses dans la désignation)
    - poids_carton_kg (float)
    """
    rows = []
    if os.path.exists(pdf_path):
        try:
            import pdfplumber
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    txt = page.extract_text() or ""
                    for line in txt.splitlines():
                        # Exemple de ligne (extrait du PDF fourni) :
                        # "Kéfir Pamplemousse 12x33cl KÉFIR PAMPLEMOUSSE - 12X33CL (12) ... 7,56"
                        l = _norm_text(line)
                        fmt = _format_from_stock(stock)
                        if not fmt: 
                            continue
                        # Poids (dernier nombre avec virgule ou point)
                        m_w = re.findall(r"(\d+(?:[.,]\d+)?)\s*$", l)
                        poids = float(m_w[-1].replace(",", ".")) if m_w else None
                        # Référence = nombre entre parenthèses dans la désignation
                        m_ref = re.search(r"\((\d{2,7})\)", l)
                        ref = m_ref.group(1) if m_ref else ""
                        # Goût = début avant le format
                        # On prend les mots avant '12x33'/'6x75'/'4x75'
                        try:
                            head = l.lower()
                            pos = head.find(fmt)
                            gout = _norm_text(l[:pos]).replace("kéfir de fruits", "Kéfir").replace("inter - ", "")
                        except Exception:
                            gout = l
                        rows.append({
                            "gout_raw": gout,
                            "format": fmt,
                            "reference": ref,
                            "poids_carton_kg": poids,
                            "row": l,
                        })
        except Exception:
            pass

    df = pd.DataFrame(rows)
    if df.empty:
        # ---- Fallback codé en dur (issu du PDF "Sofripa-2.pdf") ----
        hard = [
            # gout, format, reference, poids
            ("Kéfir de fruits Original", "6x75",  "3383", 7.23),
            ("Kéfir de fruits Original", "12x33", "12",   7.56),
            ("Kéfir Gingembre",          "4x75",  "3382", 4.68),
            ("Kéfir Gingembre",          "6x75",  "3383", 7.23),
            ("Kéfir Gingembre",          "12x33", "12",   7.56),
            ("Kéfir Mangue Passion",     "4x75",  "3382", 4.68),
            ("Kéfir Mangue Passion",     "6x75",  "3383", 7.23),
            ("Kéfir Mangue Passion",     "12x33", "12",   7.56),
            ("Kéfir menthe citron vert", "4x75",  "3382", 4.68),
            ("Kéfir menthe citron vert", "6x75",  "3383", 7.23),
            ("Kéfir menthe citron vert", "12x33", "12",   7.56),
            ("Infusion probiotique menthe poivrée", "12x33", "12", 7.56),
            ("Kéfir Pamplemousse",       "4x75",  "3382", 4.68),
            ("Kéfir Pamplemousse",       "6x75",  "3383", 7.23),
            ("Kéfir Pamplemousse",       "12x33", "12",   7.56),
            ("Infusion probiotique Anis","12x33", "12",   7.56),
            ("IGEBA Pêche",              "12x33", "12",   7.56),
            ("Infusion probiotique Mélisse","12x33","12", 7.56),
            ("Infusion probiotique Zest d'agrumes","12x33","12", 7.56),
            ("Probiotic water Lemonbalm","12x33","12",    7.56),
            ("Probiotic water Peppermint","12x33","12",   7.56),
            ("Water kefir Mango Passion","12x33","12",    7.56),
            ("Water kefir Mint Lime",    "12x33","12",    7.56),
            ("Water kefir Grapefruit",   "12x33","12",    6.741),
            ("NIKO - Kéfir de fruits Menthe citron vert","12x33","13770014427363",6.741),
            ("NIKO - Kéfir de fruits Mangue Passion","12x33","1377...",6.741),
            ("NIKO - Kéfir de fruits Gingembre","12x33","1377...",6.741),
        ]
        df = pd.DataFrame(hard, columns=["gout_raw","format","reference","poids_carton_kg"])
    # normalisation du "goût" (clé d’agrégation)
    df["gout_key"] = df["gout_raw"].astype(str).str.strip().str.lower()
    return df[["gout_key","format","reference","poids_carton_kg"]].drop_duplicates()

# ====== UI ======
apply_theme("Fiche de ramasse — Ferment Station", "🚚")
section("Fiche de ramasse", "🚚")

# Besoin de la production sauvegardée (depuis l’onglet Production)
if "saved_production" not in st.session_state or "df_min" not in st.session_state["saved_production"]:
    st.warning("Va d’abord dans **Production** et clique **💾 Sauvegarder cette production**.")
    st.stop()

sp = st.session_state["saved_production"]
df_min_saved: pd.DataFrame = sp["df_min"].copy()
ddm_saved = dt.date.fromisoformat(sp["ddm"]) if "ddm" in sp else _tz_today_paris()

# 1) Construit la liste Goût + Format réellement présents dans le tableau de production
def _format_from_stock(stock_txt: str) -> str | None:
    return _format_tag(stock_txt or "")

opts_rows = []
for _, r in df_min_saved.iterrows():
    gout = str(r.get("GoutCanon") or "").strip()
    fmt = _format_from_stock(r.get("Stock"))
    if gout and fmt:
        label = f"{gout} — {fmt}"
        key = (gout.lower(), fmt)
        if key not in {(o["gout_key"], o["format"]) for o in opts_rows}:
            opts_rows.append({"label": label, "gout": gout, "gout_key": gout.lower(), "format": fmt})

# -- Construction robuste de la liste "Goût — Format" depuis df_min_saved --
opts_rows = []
seen = set()

for _, r in df_min_saved.iterrows():
    gout = str(r.get("GoutCanon") or "").strip()
    stock = str(r.get("Stock") or "")
    fmt = _format_from_stock(stock)   # <-- utilise la nouvelle fonction robuste
    if gout and fmt:
        key = (gout.lower(), fmt)
        if key not in seen:
            opts_rows.append({
                "label": f"{gout} — {fmt}",
                "gout": gout,
                "gout_key": gout.lower(),
                "format": fmt,
            })
            seen.add(key)
_dbg = df_min_saved.copy()
_dbg["_fmt_detecte"] = _dbg.get("Stock", pd.Series()).apply(_format_from_stock)
st.caption("Debug format détecté (top 10) :")
st.dataframe(_dbg[["GoutCanon","Stock","_fmt_detecte"]].head(10), use_container_width=True, hide_index=True)


if not opts_rows:
    st.error(
        "Impossible de détecter les **formats** (12x33, 6x75, 4x75) dans le tableau de production sauvegardé.\n\n"
        "Vérifie la colonne **Stock** (ex. *Carton de 12 Bouteilles 33 cL*, *Pack de 6 75 cL*, etc.)."
    )
    st.write("Aperçu des 10 premières valeurs de `Stock` :", df_min_saved.get("Stock", pd.Series()).head(10))
    st.stop()

opts_df = pd.DataFrame(opts_rows).sort_values(by="label").reset_index(drop=True)


# 2) Charge le mapping Référence + Poids/carton depuis le PDF
ref_map = _parse_reference_pdf(REF_PDF_PATH)

# 3) Sidebar : dates
with st.sidebar:
    st.header("Paramètres")
    date_creation = _tz_today_paris()
    date_ramasse = st.date_input("Date de ramasse", value=date_creation)
    st.caption(f"DATE DE CRÉATION : **{date_creation.strftime('%d/%m/%Y')}**")
    st.caption(f"DDM (depuis Production) : **{ddm_saved.strftime('%d/%m/%Y')}**")

# 4) Sélecteur multi-produits (Goût + Format)
st.subheader("Sélection des produits")
selection_labels = st.multiselect(
    "Produits à inclure (Goût — Format)",
    options=opts_df["label"].tolist(),
    default=opts_df["label"].tolist()[:1],
)

if not selection_labels:
    st.info("Sélectionne au moins un produit.")
    st.stop()

# 5) Table éditable avec les colonnes demandées
rows = []
for lab in selection_labels:
    row = opts_df[opts_df["label"] == lab].iloc[0]
    gout_key = row["gout_key"]; fmt = row["format"]
    # joint avec le mapping PDF
    m = ref_map[(ref_map["gout_key"] == gout_key) & (ref_map["format"] == fmt)]
    ref = str(m["reference"].iloc[0]) if not m.empty else ""
    poids_carton = float(m["poids_carton_kg"].iloc[0]) if (not m.empty and pd.notna(m["poids_carton_kg"].iloc[0])) else 0.0
    rows.append({
        "Référence": ref,
        "Produit (goût + format)": lab,
        "DDM": ddm_saved.strftime("%d/%m/%Y"),
        "Quantité cartons": 0,
        "Quantité palettes": 0,          # calculé ensuite
        "Poids palettes (kg)": 0,        # calculé ensuite (total poids des cartons)
        "_format": fmt,
        "_poids_carton": poids_carton,
    })

base_df = pd.DataFrame(rows, columns=[
    "Référence", "Produit (goût + format)", "DDM",
    "Quantité cartons", "Quantité palettes", "Poids palettes (kg)",
    "_format","_poids_carton"
])

st.caption("Renseigne **Quantité cartons**. Les **palettes** et le **poids** seront calculés automatiquement.")
edited = st.data_editor(
    base_df,
    key="ramasse_editor_v2",
    use_container_width=True,
    hide_index=True,
    column_config={
        "Quantité cartons": st.column_config.NumberColumn(min_value=0, step=1),
        "Quantité palettes": st.column_config.NumberColumn(disabled=True),
        "Poids palettes (kg)": st.column_config.NumberColumn(disabled=True, format="%.0f"),
    }
)

# 6) Calculs automatiques
def _apply_calculs(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    # Quantité palettes = ceil(cartons / cartons_par_palette)
    def qp(row):
        fmt = row["_format"]; c = int(pd.to_numeric(row["Quantité cartons"], errors="coerce") or 0)
        cpp = PALLETS_RULES.get(fmt, 0)
        return int(math.ceil(c / cpp)) if cpp else 0
    out["Quantité palettes"] = out.apply(qp, axis=1)
    # Poids total (kg) = cartons × poids/carton (issu du PDF)
    def pw(row):
        c = int(pd.to_numeric(row["Quantité cartons"], errors="coerce") or 0)
        pc = float(pd.to_numeric(row["_poids_carton"], errors="coerce") or 0.0)
        return round(c * pc, 0)
    out["Poids palettes (kg)"] = out.apply(pw, axis=1).astype(int)
    return out

df_calc = _apply_calculs(edited)

# 7) KPIs
tot_cartons = int(pd.to_numeric(df_calc["Quantité cartons"], errors="coerce").fillna(0).sum())
tot_palettes = int(pd.to_numeric(df_calc["Quantité palettes"], errors="coerce").fillna(0).sum())
tot_poids = int(pd.to_numeric(df_calc["Poids palettes (kg)"], errors="coerce").fillna(0).sum())

c1, c2, c3 = st.columns(3)
with c1: kpi("Total cartons", f"{tot_cartons:,}".replace(",", " "))
with c2: kpi("Total palettes", f"{tot_palettes}")
with c3: kpi("Poids total (kg)", f"{tot_poids:,}".replace(",", " "))

st.dataframe(
    df_calc[["Référence","Produit (goût + format)","DDM","Quantité cartons","Quantité palettes","Poids palettes (kg)"]],
    use_container_width=True, hide_index=True
)

# 8) PDF
from fpdf import FPDF
def _pdf_ramasse(date_creation: dt.date, date_ramasse: dt.date, df_lines: pd.DataFrame, totals: dict) -> bytes:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=12)
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 14); pdf.cell(0, 8, "FERMENT STATION", ln=1)
    pdf.set_font("Helvetica", "", 10);  pdf.cell(0, 6, f"DATE DE CREATION : {date_creation.strftime('%d/%m/%Y')}", ln=1)
    pdf.cell(0, 6, f"DATE DE RAMMASSE : {date_ramasse.strftime('%d/%m/%Y')}", ln=1)
    pdf.ln(2)
    headers = ["Référence","Produit (goût + format)","DDM","Quantité cartons","Quantité palettes","Poids palettes (kg)"]
    widths  = [28, 86, 20, 28, 28, 28]
    pdf.set_font("Helvetica", "B", 10)
    for h,w in zip(headers, widths): pdf.cell(w, 8, h, border=1, align="C")
    pdf.ln(8); pdf.set_font("Helvetica", "", 10)
    for _, r in df_lines.iterrows():
        row = [
            str(r["Référence"]),
            str(r["Produit (goût + format)"]),
            str(r["DDM"]),
            str(int(pd.to_numeric(r["Quantité cartons"], errors="coerce") or 0)),
            str(int(pd.to_numeric(r["Quantité palettes"], errors="coerce") or 0)),
            str(int(pd.to_numeric(r["Poids palettes (kg)"], errors="coerce") or 0)),
        ]
        for i,(txt,w) in enumerate(zip(row, widths)):
            pdf.cell(w, 8, txt, border=1, align="C" if i!=1 else "L")
        pdf.ln(8)
    pdf.set_font("Helvetica","B",10)
    pdf.cell(widths[0]+widths[1]+widths[2],8,"",border=1)
    pdf.cell(widths[3],8,str(totals["cartons"]),border=1,align="C")
    pdf.cell(widths[4],8,str(totals["palettes"]),border=1,align="C")
    pdf.cell(widths[5],8,str(totals["poids"]),border=1,align="C")
    return pdf.output(dest="S").encode("latin1")

st.markdown("---")
if st.button("🧾 Générer la fiche de ramasse (PDF)", use_container_width=True, type="primary"):
    if tot_cartons <= 0:
        st.error("Renseigne au moins une **Quantité cartons** > 0.")
    else:
        pdf_bytes = _pdf_ramasse(
            _tz_today_paris(), 
            st.session_state.get("date_ramasse", _tz_today_paris()) if False else date_ramasse,
            df_calc[["Référence","Produit (goût + format)","DDM","Quantité cartons","Quantité palettes","Poids palettes (kg)"]],
            {"cartons": tot_cartons, "palettes": tot_palettes, "poids": tot_poids},
        )
        fname = f"BL_enlevements_{_tz_today_paris().strftime('%Y%m%d')}.pdf"
        st.download_button("📥 Télécharger le PDF", data=pdf_bytes, file_name=fname, mime="application/pdf", use_container_width=True)
