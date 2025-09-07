# pages/03_Fiche_de_ramasse.py
from __future__ import annotations
import os, re, math, datetime as dt
import unicodedata
import pandas as pd
import streamlit as st
from dateutil.tz import gettz

from common.design import apply_theme, section, kpi

# ============ Réglages fichiers ============
INFO_CSV_PATH = "info_FDR.csv"          # catalogue produits (Code-barre, Poids, etc.)
TEMPLATE_PDF_PATH = "assets/BL_template.pdf"  # PDF modèle exporté depuis Excel (vierge)

# ============ Réglages de placement (mm) ============ 
# Coin haut-gauche du tableau (mm depuis le bord gauche ; Y depuis le BAS de page)
TAB_X0_MM    = 20.0
TAB_YTOP_MM  = 167.0
ROW_H_MM     = 10.5
# Largeurs colonnes (doivent ≈ 186 mm au total pour marges de 12 mm)
# Réf | Produit | DDM | Qté cartons | Qté palettes | Poids palettes
COL_W_MM = [25, 80, 25, 18, 18, 20]

# Positions des valeurs dans le cartouche (mm depuis le bord gauche / le bas de page)
X_DATE_VAL_MM = 75.0
Y_CREATION_MM = 210.0
Y_RAMMASSE_MM = 198.0

X_DEST_MM     = 129.0
Y_DEST_TIT_MM = 210.0
Y_DEST_L1_MM  = 198.0
Y_DEST_L2_MM  = 192.0
Y_DEST_L3_MM  = 186.0

# ============ Identité / Destinataire ============
DEST_TITLE = "SOFRIPA"
DEST_LINES = [
    "ZAC du Haut de Wissous II,",
    "Rue Hélène Boucher, 91320 Wissous",
]

# =========================================================
#                      Utils génériques
# =========================================================
def _today_paris() -> dt.date:
    return dt.datetime.now(gettz("Europe/Paris")).date()

def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip()

def _strip_accents(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    return "".join(ch for ch in s if not unicodedata.combining(ch))

def _canon(s: str) -> str:
    s = _strip_accents(str(s or "")).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def _format_from_stock(stock_txt: str) -> str | None:
    """Détecte 12x33 / 6x75 / 4x75 depuis la colonne Stock."""
    if not stock_txt:
        return None
    s = str(stock_txt).lower().replace("×","x").replace("\u00a0"," ")
    vol = None
    if "0.33" in s or re.search(r"33\s*c?l", s): vol = 33
    elif "0.75" in s or re.search(r"75\s*c?l", s): vol = 75
    nb = None
    m = re.search(r"(?:carton|pack)\s*de\s*(12|6|4)\b", s)
    if not m: m = re.search(r"\b(12|6|4)\b", s)
    if m: nb = int(m.group(1))
    if vol == 33 and nb == 12: return "12x33"
    if vol == 75 and nb == 6:  return "6x75"
    if vol == 75 and nb == 4:  return "4x75"
    return None

# =========================================================
#        Lecture du catalogue (info_FDR.csv) + lookup
# =========================================================
@st.cache_data(show_spinner=False)
def _load_catalog(path: str) -> pd.DataFrame:
    """Lit info_FDR.csv (séparateur virgule)."""
    if not os.path.exists(path):
        return pd.DataFrame(columns=["Produit","Format","Désignation","Code-barre","Poids"])
    df = pd.read_csv(path, encoding="utf-8")
    # normalisations utiles
    for c in ["Produit","Format","Désignation","Code-barre"]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()
    if "Poids" in df.columns:
        df["Poids"] = pd.to_numeric(df["Poids"], errors="coerce")
    # colonnes auxiliaires
    df["_format_norm"] = df.get("Format","").astype(str).str.lower().str.replace("cl","").str.replace(" ","", regex=False)
    df["_format_norm"] = df["_format_norm"].str.replace("x", "x")  # idempotent
    df["_canon_prod"] = df.get("Produit","").map(_canon)
    df["_canon_des"]  = df.get("Désignation","").map(lambda s: _canon(re.sub(r"\(.*\)","", s)))
    return df

def _csv_lookup(catalog: pd.DataFrame, prod_hint: str, fmt: str) -> tuple[str, float] | None:
    """
    Trouve (référence, poids_carton) dans le CSV à partir d'un hint produit et d'un format 12x33/6x75/4x75.
    - référence = 6 derniers chiffres de 'Code-barre'
    - poids_carton = 'Poids'
    """
    if catalog is None or catalog.empty or not fmt:
        return None
    fmt_norm = fmt.lower().replace("cl","")
    # 1) match via Format + Produit canonisé (le plus robuste)
    c_prod = _canon(prod_hint)
    cand = catalog[(catalog["_format_norm"].str.contains(fmt_norm)) & (catalog["_canon_prod"] == c_prod)]
    if cand.empty:
        # 2) match via Format + Désignation canonisée (fallback)
        cand = catalog[(catalog["_format_norm"].str.contains(fmt_norm)) & (catalog["_canon_des"].str.contains(_canon(prod_hint)))]
    if cand.empty:
        # 3) dernier recours: seulement Format
        cand = catalog[catalog["_format_norm"].str.contains(fmt_norm)]
    if cand.empty:
        return None
    row = cand.iloc[0]
    code = re.sub(r"\D+", "", str(row.get("Code-barre","")))
    ref6 = code[-6:] if len(code) >= 6 else code
    poids = float(row.get("Poids") or 0.0)
    return ref6, poids

# =========================================================
#               Overlay PDF (ReportLab + pypdf)
# =========================================================
from reportlab.pdfgen import canvas as _rl_canvas
from reportlab.lib.pagesizes import A4 as _A4
from reportlab.lib.units import mm as _UNIT_MM
from pypdf import PdfReader as _PdfReader, PdfWriter as _PdfWriter
import io as _io

def _mm(x: float) -> float:
    return float(x) * _UNIT_MM

def _draw_centered(c, x_left_mm, width_mm, y_mm, txt, font="Helvetica", size=10):
    c.setFont(font, size)
    x_left = _mm(x_left_mm);  w = _mm(width_mm);  y = _mm(y_mm)
    txt = str(txt)
    tw = c.stringWidth(txt, font, size)
    c.drawString(x_left + (w - tw) / 2.0, y, txt)

def _draw_right(c, x_right_mm, y_mm, txt, font="Helvetica", size=10):
    c.setFont(font, size)
    txt = str(txt)
    x_right = _mm(x_right_mm); y = _mm(y_mm)
    tw = c.stringWidth(txt, font, size)
    c.drawString(x_right - tw, y, txt)

def _pdf_txt(x) -> str:
    """Rend le texte 'latin-1 safe' (fpdf/rl) : remplace tirets/guillemets typographiques."""
    s = str(x)
    s = (s.replace("—","-").replace("–","-").replace("•","-")
           .replace("’","'").replace("‘","'")
           .replace("“",'"').replace("”",'"'))
    try:
        s.encode("latin-1")
        return s
    except UnicodeEncodeError:
        return s.encode("latin-1","ignore").decode("latin-1")

def _pdf_ramasse_from_template(
    date_creation: dt.date,
    date_ramasse: dt.date,
    df_lines: pd.DataFrame,
    totals: dict,
) -> bytes:
    """
    Superpose un overlay texte (ReportLab) au PDF modèle (pypdf).
    Le modèle contient TOUTE la mise en page : cadres, en-têtes, titres, etc.
    """
    # ---- 1) Overlay en mémoire ----
    buf = _io.BytesIO()
    c = _rl_canvas.Canvas(buf, pagesize=_A4)

    # Cartouche : seulement les valeurs
    c.setFont("Helvetica", 10)
    c.drawString(_mm(X_DATE_VAL_MM), _mm(Y_CREATION_MM), date_creation.strftime("%d/%m/%Y"))
    c.drawString(_mm(X_DATE_VAL_MM), _mm(Y_RAMMASSE_MM), date_ramasse.strftime("%d/%m/%Y"))

    # Destinataire
    c.setFont("Helvetica-Bold", 10)
    c.drawString(_mm(X_DEST_MM), _mm(Y_DEST_TIT_MM), _pdf_txt(DEST_TITLE))
    c.setFont("Helvetica", 10)
    _dest = DEST_LINES + ["", "", ""]
    c.drawString(_mm(X_DEST_MM), _mm(Y_DEST_L1_MM), _pdf_txt(_dest[0]))
    c.drawString(_mm(X_DEST_MM), _mm(Y_DEST_L2_MM), _pdf_txt(_dest[1]))
    c.drawString(_mm(X_DEST_MM), _mm(Y_DEST_L3_MM), _pdf_txt(_dest[2]))

    # Tableau : placements
    col_x_mm = [TAB_X0_MM]
    for w in COL_W_MM[:-1]:
        col_x_mm.append(col_x_mm[-1] + w)
    X_REF, X_PROD, X_DDM, X_QC, X_QP, X_POIDS = col_x_mm
    XR_QC    = X_QC    + COL_W_MM[3] - 1.5
    XR_QP    = X_QP    + COL_W_MM[4] - 1.5
    XR_POIDS = X_POIDS + COL_W_MM[5] - 1.5

    def row_y_mm(i_row: int) -> float:
        # Ligne i (0-based) → position Y bas-aligne texte
        return TAB_YTOP_MM - (i_row * ROW_H_MM) + 2.0

    # Lignes
    c.setFont("Helvetica", 10)
    for i, (_, r) in enumerate(df_lines.iterrows()):
        y_mm = row_y_mm(i)
        ref   = _pdf_txt(r["Référence"])
        prod  = _pdf_txt(str(r["Produit (goût + format)"]).upper())
        ddm   = _pdf_txt(r["DDM"])
        qc    = int(pd.to_numeric(r["Quantité cartons"],  errors="coerce") or 0)
        qp    = int(pd.to_numeric(r["Quantité palettes"], errors="coerce") or 0)
        poids = int(pd.to_numeric(r["Poids palettes (kg)"], errors="coerce") or 0)

        c.drawString(_mm(X_REF  + 2.0), _mm(y_mm), ref)
        c.drawString(_mm(X_PROD + 2.0), _mm(y_mm), prod)
        _draw_centered(c, X_DDM, COL_W_MM[2], y_mm, ddm, size=10)
        _draw_right(c, XR_QC,    y_mm, qc,    size=10)
        _draw_right(c, XR_QP,    y_mm, qp,    size=10)
        _draw_right(c, XR_POIDS, y_mm, poids, size=10)

    # TOTAL (ligne suivante)
    y_tot_mm = row_y_mm(len(df_lines))
    tot_lbl_w = COL_W_MM[0] + COL_W_MM[1] + COL_W_MM[2]
    _draw_right(c, X_DDM + tot_lbl_w - 2.0, y_tot_mm, "TOTAL", font="Helvetica-Bold", size=10)
    _draw_right(c, XR_QC,    y_tot_mm, int(totals.get("cartons", 0)), font="Helvetica-Bold", size=10)
    _draw_right(c, XR_QP,    y_tot_mm, int(totals.get("palettes", 0)), font="Helvetica-Bold", size=10)
    _draw_right(c, XR_POIDS, y_tot_mm, int(totals.get("poids", 0)),    font="Helvetica-Bold", size=10)

    c.showPage()
    c.save()
    overlay_bytes = buf.getvalue()

    # ---- 2) Fusion avec le modèle ----
    if not os.path.exists(TEMPLATE_PDF_PATH):
        return overlay_bytes  # fallback si modèle absent

    bg = _PdfReader(TEMPLATE_PDF_PATH)
    ov = _PdfReader(_io.BytesIO(overlay_bytes))
    page_bg = bg.pages[0]
    page_ov = ov.pages[0]
    page_bg.merge_page(page_ov)

    w = _PdfWriter()
    w.add_page(page_bg)
    out = _io.BytesIO()
    w.write(out)
    return out.getvalue()

# =========================================================
#                           UI
# =========================================================
apply_theme("Fiche de ramasse — Ferment Station", "🚚")
section("Fiche de ramasse", "🚚")

# Besoin de la production sauvegardée
if "saved_production" not in st.session_state or "df_min" not in st.session_state["saved_production"]:
    st.warning("Va d’abord dans **Production** et clique **💾 Sauvegarder cette production**.")
    st.stop()

sp = st.session_state["saved_production"]
df_min_saved: pd.DataFrame = sp["df_min"].copy()
ddm_saved = dt.date.fromisoformat(sp["ddm"]) if "ddm" in sp else _today_paris()

# 1) Options = Goût + Format présents dans df_min sauvegardé
opts_rows, seen = [], set()
for _, r in df_min_saved.iterrows():
    gout = str(r.get("GoutCanon") or "").strip()
    fmt  = _format_from_stock(r.get("Stock"))
    if not (gout and fmt):
        continue
    key = (gout.lower(), fmt)
    if key in seen:
        continue
    seen.add(key)
    opts_rows.append({
        "label": f"{gout} — {fmt}",
        "gout": gout,
        "format": fmt,
        "prod_hint": str(r.get("Produit") or "").strip(),  # on le réutilisera pour matcher le CSV
    })

if not opts_rows:
    st.error("Impossible de détecter les **formats** (12x33, 6x75, 4x75) dans le tableau de production sauvegardé.")
    st.stop()

opts_df = pd.DataFrame(opts_rows).sort_values(by="label").reset_index(drop=True)

# 2) Catalogue (CSV)
catalog = _load_catalog(INFO_CSV_PATH)
if catalog.empty:
    st.warning("⚠️ `info_FDR.csv` introuvable ou vide — les références/poids ne pourront pas être calculés correctement.")

# 3) Sidebar : dates
with st.sidebar:
    st.header("Paramètres")
    date_creation = _today_paris()
    date_ramasse = st.date_input("Date de ramasse", value=date_creation)
    st.caption(f"DATE DE CRÉATION : **{date_creation.strftime('%d/%m/%Y')}**")
    st.caption(f"DDM (depuis Production) : **{ddm_saved.strftime('%d/%m/%Y')}**")

# 4) Sélecteur multi-produits
st.subheader("Sélection des produits")
selection_labels = st.multiselect(
    "Produits à inclure (Goût — Format)",
    options=opts_df["label"].tolist(),
    default=opts_df["label"].tolist(),
)

if not selection_labels:
    st.info("Sélectionne au moins un produit.")
    st.stop()

# 5) Base table (utilise le CSV pour Ref + Poids)
meta_by_label = {}
rows = []
for lab in selection_labels:
    row_opt = opts_df.loc[opts_df["label"] == lab].iloc[0]
    prod_hint = row_opt["prod_hint"]
    fmt = row_opt["format"]

    ref = ""; poids_carton = 0.0
    lk = _csv_lookup(catalog, prod_hint, fmt)
    if lk:
        ref, poids_carton = lk

    meta_by_label[lab] = {"_format": fmt, "_poids_carton": poids_carton, "_reference": ref}

    rows.append({
        "Référence": ref,
        "Produit (goût + format)": lab.replace(" — ", " - "),  # visuel propre
        "DDM": ddm_saved.strftime("%d/%m/%Y"),
        "Quantité cartons": 0,
        "Quantité palettes": 0,
        "Poids palettes (kg)": 0,
    })

display_cols = ["Référence","Produit (goût + format)","DDM","Quantité cartons","Quantité palettes","Poids palettes (kg)"]
base_df = pd.DataFrame(rows, columns=display_cols)

st.caption("Renseigne **Quantité cartons** et, si besoin, **Quantité palettes**. Le **poids** se calcule automatiquement (cartons × poids/carton depuis `info_FDR.csv`).")
edited = st.data_editor(
    base_df,
    key="ramasse_editor_v4",
    use_container_width=True,
    hide_index=True,
    column_config={
        "Quantité cartons": st.column_config.NumberColumn(min_value=0, step=1),
        "Quantité palettes": st.column_config.NumberColumn(min_value=0, step=1),
        "Poids palettes (kg)": st.column_config.NumberColumn(disabled=True, format="%.0f"),
    },
)

# 6) Calculs automatiques (poids)
def _apply_calculs(df_disp: pd.DataFrame) -> pd.DataFrame:
    out = df_disp.copy()
    poids = []
    for _, r in out.iterrows():
        lab = str(r["Produit (goût + format)"]).replace(" - ", " — ")
        meta = meta_by_label.get(lab, meta_by_label.get(str(r["Produit (goût + format)"]), {}))
        pc = float(meta.get("_poids_carton", 0.0))
        cartons = int(pd.to_numeric(r["Quantité cartons"], errors="coerce") or 0)
        poids.append(int(round(cartons * pc, 0)))
    out["Poids palettes (kg)"] = poids
    return out

df_calc = _apply_calculs(edited)

# KPIs
tot_cartons = int(pd.to_numeric(df_calc["Quantité cartons"], errors="coerce").fillna(0).sum())
tot_palettes = int(pd.to_numeric(df_calc["Quantité palettes"], errors="coerce").fillna(0).sum())
tot_poids = int(pd.to_numeric(df_calc["Poids palettes (kg)"], errors="coerce").fillna(0).sum())

c1, c2, c3 = st.columns(3)
with c1: kpi("Total cartons", f"{tot_cartons:,}".replace(",", " "))
with c2: kpi("Total palettes", f"{tot_palettes}")
with c3: kpi("Poids total (kg)", f"{tot_poids:,}".replace(",", " "))

st.dataframe(
    df_calc[display_cols],
    use_container_width=True, hide_index=True
)

# 7) Génération PDF (overlay + modèle)
st.markdown("---")
if st.button("🧾 Générer la fiche de ramasse (PDF)", use_container_width=True, type="primary"):
    if tot_cartons <= 0:
        st.error("Renseigne au moins une **Quantité cartons** > 0.")
    else:
        pdf_bytes = _pdf_ramasse_from_template(
            _today_paris(), date_ramasse,
            df_calc[display_cols],
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
