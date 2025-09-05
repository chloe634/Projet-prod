# pages/03_Fiche_de_ramasse.py
from __future__ import annotations
import os, re, math, datetime as dt
import pandas as pd
import streamlit as st
from dateutil.tz import gettz
from fpdf import FPDF

from common.design import apply_theme, section, kpi

# ---- Chemin du PDF de référence mis dans /assets du repo ----
# Renomme-le si besoin et adapte la ligne ci-dessous.
REF_PDF_PATH = "assets/LOG_EN_001_01 BL enlèvements Sofripa-2.pdf"

# ====== Règles & constantes ======
# Cartons par palette (à ajuster si besoin)
PALLETS_RULES = {
    "12x33": 108,  # cartons/palette
    "6x75":   84,
    "4x75":  100,
}
# Bouteilles par carton (calcul interne)
BTL_PER_CARTON = {"12x33": 12, "6x75": 6, "4x75": 4}

def _today_paris() -> dt.date:
    return dt.datetime.now(gettz("Europe/Paris")).date()

def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip()

# ---------- Détection format depuis la colonne "Stock" ----------
def _format_from_stock(stock_txt: str) -> str | None:
    """
    Détecte 12x33 / 6x75 / 4x75 dans des libellés très libres, ex.:
    'Carton de 12 Bouteilles - 0.33L', 'Carton de 6 Bouteilles 75cl Verralia - 0.75L',
    'Pack de 4 Bouteilles 75cl SAFT - 0.75L', etc.
    Stratégie :
      1) repère le volume (33cl ou 0.33L → 33 ; 75cl ou 0.75L → 75)
      2) repère le nb de bouteilles (12, 6 ou 4)
      3) mappe (nb, vol) → format
    """
    if not stock_txt:
        return None
    s = str(stock_txt).lower().replace("×", "x").replace(",", ".").replace("\u00a0", " ")

    # volume
    vol = None
    if "0.33l" in s or "0.33 l" in s or re.search(r"\b33\s*c?l\b", s):
        vol = 33
    elif "0.75l" in s or "0.75 l" in s or re.search(r"\b75\s*c?l\b", s):
        vol = 75

    # nb bouteilles (on cherche après "carton|pack|bouteilles|de" si possible)
    nb = None
    m = re.search(r"(?:carton|pack|bouteilles|de)\D*?(12|6|4)\b", s)
    if m:
        nb = int(m.group(1))
    else:
        m = re.search(r"\b(12|6|4)\b", s)
        if m:
            nb = int(m.group(1))

    if vol == 33 and nb == 12:
        return "12x33"
    if vol == 75 and nb == 6:
        return "6x75"
    if vol == 75 and nb == 4:
        return "4x75"
    return None

# ---------- Lecture du PDF Sofripa → références & poids/carton ----------
def _parse_reference_pdf(pdf_path: str) -> pd.DataFrame:
    """
    Retourne un DataFrame avec colonnes :
    - gout_key  (goût normalisé en minuscule)
    - format    ('12x33' / '6x75' / '4x75')
    - reference (texte)
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
                        l = _norm(line)
                        # repère format
                        fmt = None
                        if re.search(r"\b12\s*x\s*33\b", l, re.I) or re.search(r"\b12\s*x?\s*33\s*c?l\b", l, re.I):
                            fmt = "12x33"
                        elif re.search(r"\b6\s*x\s*75\b", l, re.I) or re.search(r"\b6\s*x?\s*75\s*c?l\b", l, re.I):
                            fmt = "6x75"
                        elif re.search(r"\b4\s*x\s*75\b", l, re.I) or re.search(r"\b4\s*x?\s*75\s*c?l\b", l, re.I):
                            fmt = "4x75"
                        if not fmt:
                            continue
                        # référence = nombre entre parenthèses si présent
                        m_ref = re.search(r"\(([\dA-Za-z\-]+)\)", l)
                        ref = m_ref.group(1) if m_ref else ""
                        # poids : dernier nombre au bout de la ligne
                        m_w = re.findall(r"(\d+(?:[.,]\d+)?)\s*$", l)
                        poids = float(m_w[-1].replace(",", ".")) if m_w else None
                        # goût = début avant le format (approx)
                        pos = l.lower().find(fmt[:2])  # '12' / '6 ' / '4 '
                        gout_raw = _norm(l[:pos]) if pos > 0 else l
                        rows.append({
                            "gout_key": _norm(gout_raw).lower(),
                            "format": fmt,
                            "reference": ref,
                            "poids_carton_kg": poids,
                        })
        except Exception:
            pass

    df = pd.DataFrame(rows)
    if df.empty:
        # Fallback minimal (au cas où le PDF n'est pas lisible) — à enrichir si besoin.
        fallback = [
            ("kéfir gingembre", "12x33", "12",  7.56),
            ("kéfir pamplemousse", "12x33", "12", 7.56),
            ("kéfir mangue passion", "12x33", "12", 7.56),
            ("kéfir menthe citron vert", "12x33", "12", 7.56),
            ("kéfir de fruits original", "12x33", "12", 7.56),
            ("kéfir gingembre", "6x75", "3383", 7.23),
            ("kéfir pamplemousse", "6x75", "3383", 7.23),
            ("kéfir mangue passion", "6x75", "3383", 7.23),
            ("kéfir menthe citron vert", "6x75", "3383", 7.23),
            ("kéfir gingembre", "4x75", "3382", 4.68),
            ("kéfir pamplemousse", "4x75", "3382", 4.68),
            ("kéfir mangue passion", "4x75", "3382", 4.68),
            ("kéfir menthe citron vert", "4x75", "3382", 4.68),
        ]
        df = pd.DataFrame(fallback, columns=["gout_key","format","reference","poids_carton_kg"])
    return df.drop_duplicates()

# =========================  UI  =========================
apply_theme("Fiche de ramasse — Ferment Station", "🚚")
section("Fiche de ramasse", "🚚")

# Besoin de la production sauvegardée
if "saved_production" not in st.session_state or "df_min" not in st.session_state["saved_production"]:
    st.warning("Va d’abord dans **Production** et clique **💾 Sauvegarder cette production**.")
    st.stop()

sp = st.session_state["saved_production"]
df_min_saved: pd.DataFrame = sp["df_min"].copy()
ddm_saved = dt.date.fromisoformat(sp["ddm"]) if "ddm" in sp else _today_paris()

# ---- Construire la liste (Goût — Format) depuis df_min sauvegardé ----
opts_rows, seen = [], set()
for _, r in df_min_saved.iterrows():
    gout = str(r.get("GoutCanon") or "").strip()
    fmt = _format_from_stock(r.get("Stock"))
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

if not opts_rows:
    st.error("Impossible de détecter les **formats** (12x33, 6x75, 4x75) dans le tableau de production sauvegardé.")
    st.stop()

opts_df = pd.DataFrame(opts_rows).sort_values(by="label").reset_index(drop=True)

# ---- Mapping Références & Poids/carton depuis le PDF ----
ref_map = _parse_reference_pdf(REF_PDF_PATH)

# ---- Sidebar : dates ----
with st.sidebar:
    st.header("Paramètres")
    date_creation = _today_paris()
    date_ramasse = st.date_input("Date de ramasse", value=date_creation)
    st.caption(f"DATE DE CRÉATION : **{date_creation.strftime('%d/%m/%Y')}**")
    st.caption(f"DDM (depuis Production) : **{ddm_saved.strftime('%d/%m/%Y')}**")

# ---- Sélecteur multi-produits ----
st.subheader("Sélection des produits")
selection_labels = st.multiselect(
    "Produits à inclure (Goût — Format)",
    options=opts_df["label"].tolist(),
    default=opts_df["label"].tolist()[:1],
)

if not selection_labels:
    st.info("Sélectionne au moins un produit.")
    st.stop()

# ---- Table éditable avec les colonnes demandées ----
rows = []
for lab in selection_labels:
    row = opts_df[opts_df["label"] == lab].iloc[0]
    gout_key = row["gout_key"]; fmt = row["format"]
    # joint avec mapping PDF (par goût normalisé + format)
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

# ---- Calculs automatiques (palettes & poids) ----
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

    # (calcul interne non affiché) — bouteilles = cartons × (12/6/4)
    out["_bouteilles"] = [
        int(pd.to_numeric(c, errors="coerce") or 0) * BTL_PER_CARTON.get(fmt, 0)
        for c, fmt in zip(out["Quantité cartons"], out["_format"])
    ]
    return out

df_calc = _apply_calculs(edited)

# ---- KPIs ----
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

# ---- Génération PDF ----
def _pdf_ramasse(date_creation: dt.date, date_ramasse: dt.date,
                 df_lines: pd.DataFrame, totals: dict) -> bytes:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=12)
    pdf.add_page()

    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 8, "FERMENT STATION — FICHE DE RAMASSE", ln=1)
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 6, f"DATE DE CREATION : {date_creation.strftime('%d/%m/%Y')}", ln=1)
    pdf.cell(0, 6, f"DATE DE RAMASSE : {date_ramasse.strftime('%d/%m/%Y')}", ln=1)
    pdf.ln(2)

    headers = ["Référence","Produit (goût + format)","DDM","Quantité cartons","Quantité palettes","Poids palettes (kg)"]
    widths  = [28, 86, 20, 28, 28, 28]

    pdf.set_font("Helvetica", "B", 10)
    for h,w in zip(headers, widths):
        pdf.cell(w, 8, h, border=1, align="C")
    pdf.ln(8)

    pdf.set_font("Helvetica", "", 10)
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
            pdf.cell(w, 8, txt, border=1, align="C" if i != 1 else "L")
        pdf.ln(8)

    pdf.set_font("Helvetica","B",10)
    pdf.cell(widths[0]+widths[1]+widths[2],8,"TOTAL",border=1, align="R")
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
            _today_paris(), date_ramasse,
            df_calc[["Référence","Produit (goût + format)","DDM","Quantité cartons","Quantité palettes","Poids palettes (kg)"]],
            {"cartons": tot_cartons, "palettes": tot_palettes, "poids": tot_poids},
        )
        fname = f"BL_enlevements_{_today_paris().strftime('%Y%m%d')}.pdf"
        st.download_button(
            "📥 Télécharger le PDF",
            data=pdf_bytes,
            file_name=fname,
            mime="application/pdf",
            use_container_width=True,
        )
