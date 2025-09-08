# pages/03_Fiche_de_ramasse.py
from __future__ import annotations
import os
import re
import datetime as dt
import pandas as pd
import streamlit as st
from dateutil.tz import gettz
from fpdf import FPDF

from common.design import apply_theme, section, kpi
from common.xlsx_fill import fill_bl_enlevements_xlsx


# ---- Modèle PDF (vierge) exporté depuis Excel ----
TEMPLATE_PDF_PATH = "assets/BL_template.pdf"

# ---- Réglages de placement (mm) ----
# Position du tableau (coin supérieur gauche) sur le modèle
TAB_X0_MM = 20.0
TAB_YTOP_MM = 167.0    # distance depuis bas de page (A4) -> texte première ligne
ROW_H_MM   = 10.5

# Largeurs de colonnes (mm) = 186 mm total avec marges de 12 mm de chaque côté
# Réf | Produit | DDM | Qté cartons | Qté palettes | Poids palettes
COL_W_MM = [25, 80, 25, 18, 18, 20]

# Positions des valeurs dans le cartouche (mm depuis bord gauche / depuis bas de page)
X_DATE_VAL_MM = 75.0
Y_CREATION_MM = 210.0
Y_RAMMASSE_MM = 198.0

X_DEST_MM     = 129.0
Y_DEST_TIT_MM = 210.0
Y_DEST_L1_MM  = 198.0
Y_DEST_L2_MM  = 192.0
Y_DEST_L3_MM  = 186.0

from reportlab.pdfgen import canvas as _rl_canvas
from reportlab.lib.pagesizes import A4 as _A4
from reportlab.lib.units import mm as _mm
from pypdf import PdfReader as _PdfReader, PdfWriter as _PdfWriter
import io as _io

def _mm(x: float) -> float:
    return float(x) * _mm

def _draw_centered(c, x_left_mm, width_mm, y_mm, txt, font="Helvetica", size=10):
    c.setFont(font, size)
    x_left = _mm(x_left_mm);  w = _mm(width_mm);  y = _mm(y_mm)
    tw = c.stringWidth(str(txt), font, size)
    c.drawString(x_left + (w - tw) / 2.0, y, str(txt))

def _draw_right(c, x_right_mm, y_mm, txt, font="Helvetica", size=10):
    c.setFont(font, size)
    x_right = _mm(x_right_mm); y = _mm(y_mm)
    tw = c.stringWidth(str(txt), font, size)
    c.drawString(x_right - tw, y, str(txt))
def _pdf_ramasse_from_template(
    date_creation: dt.date,
    date_ramasse: dt.date,
    df_lines: pd.DataFrame,
    totals: dict,
) -> bytes:
    """
    Génère un PDF en fusionnant un overlay texte avec le modèle 'assets/BL_template.pdf'.
    Le modèle contient déjà toute la mise en page (cadres, titres, en-têtes).
    """
    # 1) Overlay en mémoire (transparent)
    buf = _io.BytesIO()
    c = _rl_canvas.Canvas(buf, pagesize=_A4)
    PAGE_W, PAGE_H = _A4

    # -- Cartouche (on ne dessine que les valeurs) --
    c.setFont("Helvetica", 10)
    c.drawString(_mm(X_DATE_VAL_MM), _mm(Y_CREATION_MM), date_creation.strftime("%d/%m/%Y"))
    c.drawString(_mm(X_DATE_VAL_MM), _mm(Y_RAMMASSE_MM), date_ramasse.strftime("%d/%m/%Y"))

    # Destinataire (titre + 3 lignes d'adresse)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(_mm(X_DEST_MM), _mm(Y_DEST_TIT_MM), _pdf_txt(DEST_TITLE))
    c.setFont("Helvetica", 10)
    _dest = DEST_LINES + ["", "", ""]  # pad
    c.drawString(_mm(X_DEST_MM), _mm(Y_DEST_L1_MM), _pdf_txt(_dest[0]))
    c.drawString(_mm(X_DEST_MM), _mm(Y_DEST_L2_MM), _pdf_txt(_dest[1]))
    c.drawString(_mm(X_DEST_MM), _mm(Y_DEST_L3_MM), _pdf_txt(_dest[2]))

    # -- Tableau : lignes + total --
    # Colonnes en mm -> positions cumulées
    col_x_mm = [TAB_X0_MM]
    for w in COL_W_MM[:-1]:
        col_x_mm.append(col_x_mm[-1] + w)
    # repères utiles
    X_REF, X_PROD, X_DDM, X_QC, X_QP, X_POIDS = col_x_mm
    # droites de fin de cellule (x_right) pour l'alignement à droite des nombres
    XR_QC   = X_QC   + COL_W_MM[3] - 1.5
    XR_QP   = X_QP   + COL_W_MM[4] - 1.5
    XR_POIDS= X_POIDS+ COL_W_MM[5] - 1.5

    # ligne de base du texte dans les cellules (un peu au-dessus du bas de cellule)
    def row_y_mm(i_row: int) -> float:
        return TAB_YTOP_MM - (i_row * ROW_H_MM) + 2.0

    # Corps
    c.setFont("Helvetica", 10)
    for i, (_, r) in enumerate(df_lines.iterrows()):
        y_mm = row_y_mm(i)
        ref   = _pdf_txt(r["Référence"])
        prod  = _pdf_txt(str(r["Produit (goût + format)"]).upper())
        ddm   = _pdf_txt(r["DDM"])
        qc    = int(pd.to_numeric(r["Quantité cartons"],  errors="coerce") or 0)
        qp    = int(pd.to_numeric(r["Quantité palettes"], errors="coerce") or 0)
        poids = int(pd.to_numeric(r["Poids palettes (kg)"], errors="coerce") or 0)

        # Référence (gauche)
        c.drawString(_mm(X_REF + 2.0), _mm(y_mm), ref)
        # Produit (gauche)
        c.drawString(_mm(X_PROD + 2.0), _mm(y_mm), prod)
        # DDM (centrée)
        _draw_centered(c, X_DDM, COL_W_MM[2], y_mm, ddm, size=10)
        # Nombres (droite)
        _draw_right(c, XR_QC,    y_mm, qc,    size=10)
        _draw_right(c, XR_QP,    y_mm, qp,    size=10)
        _draw_right(c, XR_POIDS, y_mm, poids, size=10)

    # Ligne TOTAL (sur la ligne suivant la dernière)
    y_tot_mm = row_y_mm(len(df_lines))
    # libellé TOTAL : fusion des 3 premières colonnes -> on le centre à droite de ce bloc
    tot_label_w = COL_W_MM[0] + COL_W_MM[1] + COL_W_MM[2]
    _draw_right(c, X_DDM + tot_label_w - 2.0, y_tot_mm, "TOTAL", font="Helvetica-Bold", size=10)
    _draw_right(c, XR_QC,    y_tot_mm, int(totals.get("cartons", 0)), font="Helvetica-Bold", size=10)
    _draw_right(c, XR_QP,    y_tot_mm, int(totals.get("palettes", 0)), font="Helvetica-Bold", size=10)
    _draw_right(c, XR_POIDS, y_tot_mm, int(totals.get("poids", 0)),    font="Helvetica-Bold", size=10)

    c.showPage()
    c.save()
    overlay_bytes = buf.getvalue()

    # 2) Fusion overlay + modèle
    if not os.path.exists(TEMPLATE_PDF_PATH):
        # sécurité: si le modèle manque, on renvoie juste l'overlay
        return overlay_bytes

    bg = _PdfReader(TEMPLATE_PDF_PATH)
    ov = _PdfReader(_io.BytesIO(overlay_bytes))
    page_bg = bg.pages[0]
    page_ov = ov.pages[0]
    page_bg.merge_page(page_ov)  # superpose en conservant la mise en page du modèle

    w = _PdfWriter()
    w.add_page(page_bg)
    out = _io.BytesIO()
    w.write(out)
    return out.getvalue()


# ---- Identité / Destinataire (adapter si besoin) ----
LOGO_PATH = "assets/logo_symbiose.png"  # optionnel: si absent on ignore le logo
COMPANY_LINES = [
    "FERMENT STATION",
    "Carré Ivry Bâtiment D2",
    "47 rue Ernest Renan",
    "94200 Ivry-sur-Seine - FRANCE",
    "Tél : 0967504647",
    "Site : https://www.symbiose-kefir.fr",
    "Produits issus de l'Agriculture Biologique certifié par FR-BIO-01",
]

DEST_TITLE = "SOFRIPA"
DEST_LINES = [
    "ZAC du Haut de Wissous II,",
    "Rue Hélène Boucher, 91320 Wissous",
]



# Chemin de la table de référence (remplace l'ancien PDF)
REF_CSV_PATH = "assets/info_FDR.csv"

# ---------------- Constantes ----------------
BTL_PER_CARTON = {"12x33": 12, "6x75": 6, "4x75": 4}  # calcul interne (non affiché)

# Fallbacks si une ligne n’est pas trouvée (utilisés en dernier recours)
FALLBACK_REF = {"12x33": "12", "6x75": "3383", "4x75": "3382"}
FALLBACK_POIDS_CARTON = {"12x33": 7.56, "6x75": 7.23, "4x75": 4.68}

# ---------------- Utils ----------------
def _today_paris() -> dt.date:
    return dt.datetime.now(gettz("Europe/Paris")).date()

def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip()

def _canon_gout(name: str) -> str:
    """
    Canonise un libellé pour matcher entre prod et CSV:
    - enlève préfixes (kéfir, infusion probiotique, water kefir, probiotic water, niko - ...)
    - normalise quelques variantes (citron-vert -> citron vert, peche -> pêche, etc.)
    """
    s = _norm(name).lower()
    s = re.sub(r"niko\s*-\s*", "", s)
    s = re.sub(r"k[ée]fir(\s+de\s+fruits)?\s*", "", s)
    s = re.sub(r"water\s+kefir\s*", "", s)
    s = re.sub(r"infusion\s+probiotique\s*", "", s)
    s = re.sub(r"probiotic\s+water\s*", "", s)
    s = s.replace("citron-vert", "citron vert")
    s = s.replace("zest d", "zeste d")
    s = s.replace("peche", "pêche")

    KEYWORDS = [
        ("mangue passion", ["mangue", "passion"]),
        ("gingembre", ["gingembre"]),
        ("pamplemousse", ["pamplemousse", "grapefruit"]),
        ("menthe citron vert", ["menthe", "citron", "vert", "mint", "lime"]),
        ("original", ["original"]),
        ("mélisse", ["mélisse", "lemonbalm"]),
        ("menthe poivrée", ["menthe", "poivr", "peppermint"]),
        ("zeste d'agrumes", ["zeste", "zest", "agrumes", "citrus"]),
        ("pêche", ["pêche", "peche"]),
    ]
    for canon, tokens in KEYWORDS:
        if all(t in s for t in tokens):
            return canon
    return s

def _format_from_stock(stock_txt: str) -> str | None:
    """
    Détecte 12x33 / 6x75 / 4x75 depuis la colonne Stock du tableau de prod.
    """
    if not stock_txt:
        return None
    s = str(stock_txt).lower().replace("×", "x").replace(",", ".").replace("\u00a0", " ")

    vol = None
    if "0.33l" in s or "0.33 l" in s or re.search(r"33\s*c?l", s):
        vol = 33
    elif "0.75l" in s or "0.75 l" in s or re.search(r"75\s*c?l", s):
        vol = 75

    nb = None
    m = re.search(r"(?:carton|pack)\s*de\s*(12|6|4)\b", s)
    if not m:
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

# --------- Helpers pour lire le CSV ----------
def _fmt_norm(s: str) -> str | None:
    """Normalise '12x33cl', '6x75 cl', 'Pack de 4 x 75cl' -> '12x33' / '6x75' / '4x75'."""
    if not s:
        return None
    t = str(s).lower().replace("×", "x")
    t = re.sub(r"\s+", "", t)
    m = re.match(r"(\d+)\s*x\s*(\d+)\s*c?l", t)
    if m:
        nb = int(m.group(1))
        vol = int(float(m.group(2)))
        if vol == 33 and nb == 12:
            return "12x33"
        if vol == 75 and nb in (6, 4):
            return f"{nb}x75"
    if ("75cl" in t) or ("0.75l" in t):
        if re.search(r"\b4x\b|de4|packde4", t):
            return "4x75"
        if re.search(r"\b6x\b|de6|cartonde6", t):
            return "6x75"
    if ("33cl" in t) or ("0.33l" in t):
        if re.search(r"\b12x\b|de12|cartonde12", t):
            return "12x33"
    return None

def _ref_from_codebarre(code: str) -> str:
    """Retourne les 6 derniers chiffres d'un code-barres, ou '' si impossible."""
    if not code:
        return ""
    digits = re.sub(r"\D", "", str(code))
    return digits[-6:] if len(digits) >= 6 else ""

def _parse_reference_csv(csv_path: str) -> pd.DataFrame:
    """
    Lit assets/info_FDR.csv et renvoie un DataFrame:
    colonnes -> ['canon','format','poids_carton_kg','code_barre']
    - 'canon'   : goût canonisé
    - 'format'  : '12x33' / '6x75' / '4x75'
    - 'poids_carton_kg' : Poids du carton (float)
    - 'code_barre' : EAN (str chiffrée)
    """
    rows = []
    if os.path.exists(csv_path):
        try:
            df = pd.read_csv(csv_path, encoding="utf-8", sep=",")
        except Exception:
            df = pd.read_csv(csv_path, encoding="utf-8", sep=";")

        lower = {c.lower(): c for c in df.columns}
        col_prod  = lower.get("produit")
        col_fmt   = lower.get("format")
        col_des   = lower.get("désignation") or lower.get("designation")
        col_poids = lower.get("poids")
        col_ean   = lower.get("code-barre") or lower.get("codebarre") or lower.get("code barre") or lower.get("ean")

        for _, r in df.iterrows():
            prod = str(r.get(col_prod, "") or "")
            des = str(r.get(col_des, "") or "")
            fmt1 = _fmt_norm(r.get(col_fmt, "")) or _fmt_norm(des) or _fmt_norm(prod)
            if not fmt1:
                continue

            # Poids
            poids = r.get(col_poids, None)
            try:
                if isinstance(poids, str):
                    poids = float(poids.replace(",", "."))
                elif pd.notna(poids):
                    poids = float(poids)
                else:
                    poids = None
            except Exception:
                poids = None

            # Code-barre
            ean_raw = r.get(col_ean, "")
            code_barre = re.sub(r"\D", "", str(ean_raw)) if pd.notna(ean_raw) else ""

            canon = _canon_gout(des or prod)
            rows.append(
                {
                    "canon": canon,
                    "format": fmt1,
                    "poids_carton_kg": poids,
                    "code_barre": code_barre,
                }
            )

    df_out = pd.DataFrame(rows)
    if df_out.empty:
        # Fallback minimal
        fallback = [
            ("mangue passion", "12x33", 7.56, ""),
            ("gingembre", "12x33", 7.56, ""),
            ("pamplemousse", "12x33", 7.56, ""),
            ("menthe citron vert", "12x33", 7.56, ""),
            ("original", "12x33", 7.56, ""),
            ("mangue passion", "6x75", 7.23, ""),
            ("gingembre", "6x75", 7.23, ""),
            ("pamplemousse", "6x75", 7.23, ""),
            ("menthe citron vert", "6x75", 7.23, ""),
            ("mangue passion", "4x75", 4.68, ""),
            ("gingembre", "4x75", 4.68, ""),
            ("pamplemousse", "4x75", 4.68, ""),
            ("menthe citron vert", "4x75", 4.68, ""),
        ]
        df_out = pd.DataFrame(fallback, columns=["canon", "format", "poids_carton_kg", "code_barre"])

    df_out = df_out.dropna(subset=["canon", "format"]).copy()
    df_out["canon"] = df_out["canon"].astype(str).str.strip().str.lower()
    df_out["format"] = df_out["format"].astype(str).str.strip()
    df_out = df_out.drop_duplicates(subset=["canon", "format"], keep="first").reset_index(drop=True)
    return df_out

# =========================================================
#                           UI
# =========================================================
apply_theme("Fiche de ramasse — Ferment Station", "🚚")
section("Fiche de ramasse", "🚚")
st.caption("Références & poids chargés depuis assets/info_FDR.csv")

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
    fmt = _format_from_stock(r.get("Stock"))
    if not (gout and fmt):
        continue
    key = (gout.lower(), fmt)
    if key in seen:
        continue
    seen.add(key)
    opts_rows.append({"label": f"{gout} — {fmt}", "gout": gout, "gout_key": _canon_gout(gout), "format": fmt})

if not opts_rows:
    st.error("Impossible de détecter les **formats** (12x33, 6x75, 4x75) dans le tableau de production sauvegardé.")
    st.stop()

opts_df = pd.DataFrame(opts_rows).sort_values(by="label").reset_index(drop=True)

# 2) Mapping CSV (référence + poids/carton)
ref_map = _parse_reference_csv(REF_CSV_PATH)

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

# 5) Base table (sans colonnes internes visibles)
meta_by_label = {}
rows = []
for lab in selection_labels:
    row = opts_df.loc[opts_df["label"] == lab].iloc[0]
    canon = row["gout_key"]
    fmt = row["format"]

    m = ref_map[(ref_map["canon"] == canon) & (ref_map["format"] == fmt)]

    # 1) Poids carton depuis CSV (sinon fallback)
    if not m.empty and pd.notna(m["poids_carton_kg"].iloc[0]):
        poids_carton = float(m["poids_carton_kg"].iloc[0])
    else:
        poids_carton = float(FALLBACK_POIDS_CARTON.get(fmt, 0.0))

    # 2) Référence = 6 derniers chiffres du Code-barre (sinon fallback)
    codeb = str(m["code_barre"].iloc[0]) if (not m.empty and "code_barre" in m.columns) else ""
    ref6 = _ref_from_codebarre(codeb)
    ref = ref6 if ref6 else FALLBACK_REF.get(fmt, "")

    meta_by_label[lab] = {"_format": fmt, "_poids_carton": poids_carton, "_reference": ref}

    rows.append(
        {
            "Référence": ref,
            "Produit (goût + format)": lab,
            "DDM": ddm_saved.strftime("%d/%m/%Y"),
            "Quantité cartons": 0,
            "Quantité palettes": 0,  # EDITABLE manuellement
            "Poids palettes (kg)": 0,  # calculé
        }
    )

display_cols = [
    "Référence",
    "Produit (goût + format)",
    "DDM",
    "Quantité cartons",
    "Quantité palettes",
    "Poids palettes (kg)",
]
base_df = pd.DataFrame(rows, columns=display_cols)

st.caption("Renseigne **Quantité cartons** et, si besoin, **Quantité palettes**. Le **poids** se calcule automatiquement.")
edited = st.data_editor(
    base_df,
    key="ramasse_editor_v3",
    use_container_width=True,
    hide_index=True,
    column_config={
        "Quantité cartons": st.column_config.NumberColumn(min_value=0, step=1),
        "Quantité palettes": st.column_config.NumberColumn(min_value=0, step=1),
        "Poids palettes (kg)": st.column_config.NumberColumn(disabled=True, format="%.0f"),
    },
)

# 6) Calculs automatiques (poids + bouteilles internes)
def _apply_calculs(df_disp: pd.DataFrame) -> pd.DataFrame:
    out = df_disp.copy()
    poids = []
    bouteilles = []
    for _, r in out.iterrows():
        lab = str(r["Produit (goût + format)"])
        meta = meta_by_label.get(lab, {})
        fmt = meta.get("_format", "")
        pc = float(meta.get("_poids_carton", 0.0))
        cartons = int(pd.to_numeric(r["Quantité cartons"], errors="coerce") or 0)
        # poids total arrondi à l'entier
        poids.append(int(round(cartons * pc, 0)))
        bouteilles.append(cartons * BTL_PER_CARTON.get(fmt, 0))
    out["Poids palettes (kg)"] = poids
    out["_Bouteilles (interne)"] = bouteilles  # non affichée ensuite
    return out

df_calc = _apply_calculs(edited)

# KPIs
tot_cartons = int(pd.to_numeric(df_calc["Quantité cartons"], errors="coerce").fillna(0).sum())
tot_palettes = int(pd.to_numeric(df_calc["Quantité palettes"], errors="coerce").fillna(0).sum())
tot_poids = int(pd.to_numeric(df_calc["Poids palettes (kg)"], errors="coerce").fillna(0).sum())

c1, c2, c3 = st.columns(3)
with c1:
    kpi("Total cartons", f"{tot_cartons:,}".replace(",", " "))
with c2:
    kpi("Total palettes", f"{tot_palettes}")
with c3:
    kpi("Poids total (kg)", f"{tot_poids:,}".replace(",", " "))

st.dataframe(df_calc[display_cols], use_container_width=True, hide_index=True)

def _pdf_txt(x) -> str:
    """
    Rend le texte compatible Latin-1 pour FPDF :
    - remplace les tirets/guillemets typographiques
    - force l'encodage latin-1 (en ignorant les glyphes non supportés)
    """
    s = str(x)
    s = (s
         .replace("—", "-").replace("–", "-").replace("•", "-")
         .replace("’", "'").replace("‘", "'")
         .replace("“", '"').replace("”", '"'))
    try:
        s.encode("latin-1")
        return s
    except UnicodeEncodeError:
        return s.encode("latin-1", "ignore").decode("latin-1")


# 7) Génération PDF
def _pdf_ramasse(date_creation: dt.date, date_ramasse: dt.date,
                 df_lines: pd.DataFrame, totals: dict) -> bytes:
    """
    PDF au format attendu :
    - en-tête (logo + coordonnées)
    - cartouche BON DE LIVRAISON
    - tableau bordé + ligne TOTAL
    """
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=12)
    pdf.add_page()

    # ---------- constantes de mise en page ----------
    MARGIN_L = 12
    PAGE_W = 210
    USABLE_W = PAGE_W - 2 * MARGIN_L  # 186 mm
    LINE_H = 8
    pdf.set_line_width(0.3)
    pdf.set_draw_color(0)

    # ---------- ENTÊTE ----------
    y = 14
    # Logo : hauteur fixe -> évite le chevauchement avec le texte
    logo_h = 16
    if os.path.exists(LOGO_PATH):
        try:
            pdf.image(LOGO_PATH, x=MARGIN_L, y=y, h=logo_h)
        except Exception:
            pass

    # Bloc coordonnées à droite du logo
    x_text = MARGIN_L + 40  # décalage sûr par rapport au logo
    pdf.set_xy(x_text, y)
    pdf.set_font("Helvetica", "", 11)
    for line in COMPANY_LINES[:6]:
        pdf.cell(0, 5, _pdf_txt(line), ln=1)
    pdf.ln(1)
    pdf.set_font("Helvetica", "", 8)
    if len(COMPANY_LINES) > 6:
        pdf.cell(0, 4, _pdf_txt(COMPANY_LINES[6]), ln=1)

    # espace pour respirer avant le cartouche
    pdf.ln(6)

    # ---------- CARTOUCHE 'BON DE LIVRAISON' ----------
    pdf.set_font("Helvetica", "B", 12)
    box_x = MARGIN_L
    box_y = pdf.get_y()
    box_w = USABLE_W
    box_h = 30
    pdf.rect(box_x, box_y, box_w, box_h)
    pdf.set_xy(box_x + 3, box_y + 3)
    pdf.cell(0, 6, _pdf_txt("BON DE LIVRAISON"), ln=1)

    # Colonne gauche
    pdf.set_font("Helvetica", "", 10)
    LBL_W = 46
    pdf.set_xy(box_x + 3, box_y + 12)
    pdf.cell(LBL_W, 6, _pdf_txt("DATE DE CREATION :"), ln=0)
    pdf.cell(35, 6, _pdf_txt(date_creation.strftime("%d/%m/%Y")), ln=1)

    pdf.set_xy(box_x + 3, box_y + 18)
    pdf.cell(LBL_W, 6, _pdf_txt("DATE DE RAMMASSE :"), ln=0)  # conforme au modèle
    pdf.cell(35, 6, _pdf_txt(date_ramasse.strftime("%d/%m/%Y")), ln=1)

    # Colonne droite (destinataire)
    right_x = box_x + 95
    pdf.set_xy(right_x, box_y + 12)
    pdf.cell(32, 6, _pdf_txt("DESTINATAIRE :"), ln=1)
    pdf.set_xy(right_x, box_y + 18)
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 6, _pdf_txt(DEST_TITLE), ln=1)
    pdf.set_font("Helvetica", "", 10)
    for i, l in enumerate(DEST_LINES):
        pdf.set_xy(right_x, box_y + 24 + i * 6)
        pdf.cell(0, 6, _pdf_txt(l), ln=1)

    # Place le curseur SOUS la boîte + marge
    pdf.set_y(box_y + box_h + 8)

    # ---------- TABLEAU ----------
    # Largeurs (somme = 186 mm)
    # Réf / Produit / DDM / Qté cartons / Qté palettes / Poids palettes
    col_w = [25, 80, 25, 18, 18, 20]
    headers = [
        "Référence",
        "Produit",
        "DDM",
        "Quantité cartons",
        "Quantité palettes",
        "Poids palettes (kg)",
    ]
    aligns = ["C", "L", "C", "C", "C", "C"]

    pdf.set_font("Helvetica", "B", 10)
    for h, w, a in zip(headers, col_w, aligns):
        pdf.cell(w, LINE_H, _pdf_txt(h), border=1, align=a)
    pdf.ln(LINE_H)

    pdf.set_font("Helvetica", "", 10)
    for _, r in df_lines.iterrows():
        row = [
            _pdf_txt(r["Référence"]),
            _pdf_txt(str(r["Produit (goût + format)"]).upper()),
            _pdf_txt(r["DDM"]),
            _pdf_txt(int(pd.to_numeric(r["Quantité cartons"], errors="coerce") or 0)),
            _pdf_txt(int(pd.to_numeric(r["Quantité palettes"], errors="coerce") or 0)),
            _pdf_txt(int(pd.to_numeric(r["Poids palettes (kg)"], errors="coerce") or 0)),
        ]
        for txt, w, a in zip(row, col_w, aligns):
            pdf.cell(w, LINE_H, str(txt), border=1, align=a)
        pdf.ln(LINE_H)

    # Ligne TOTAL
    pdf.set_font("Helvetica", "B", 10)
    total_label_w = col_w[0] + col_w[1] + col_w[2]
    pdf.cell(total_label_w, LINE_H, _pdf_txt("TOTAL"), border=1, align="R")
    pdf.cell(col_w[3], LINE_H, _pdf_txt(totals["cartons"]), border=1, align="C")
    pdf.cell(col_w[4], LINE_H, _pdf_txt(totals["palettes"]), border=1, align="C")
    pdf.cell(col_w[5], LINE_H, _pdf_txt(totals["poids"]), border=1, align="C")

    # ---------- Sortie binaire robuste ----------
    buf = pdf.output(dest="S")  # fpdf2 : str / bytes / bytearray selon version
    if isinstance(buf, str):
        data = buf.encode("latin-1", "ignore")
    elif isinstance(buf, bytearray):
        data = bytes(buf)
    else:
        data = buf
    return data


# 7) Génération de la fiche (XLSX ou PDF optionnel)
st.markdown("---")
col_a, col_b = st.columns([1,1])

with col_a:
    if st.button("📄 Télécharger la fiche (XLSX, modèle Sofripa)", use_container_width=True):
        try:
            xlsx_bytes = fill_bl_enlevements_xlsx(
                template_path="assets/LOG_EN_001_01 BL enlèvements Sofripa-2.xlsx",
                date_creation=_today_paris(),
                date_ramasse=date_ramasse,
                destinataire_title=DEST_TITLE,
                destinataire_lines=DEST_LINES,
                df_lines=df_calc[display_cols],
            )
            fname = f"BL_enlevements_{_today_paris().strftime('%Y%m%d')}.xlsx"
            st.download_button(
                "⬇️ Télécharger le XLSX",
                data=xlsx_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"Erreur lors du remplissage du modèle Excel : {e}")

with col_b:
    if st.button("🧾 (Option) Générer un PDF depuis l’app", use_container_width=True):
        # OPTION : conserve ton overlay PDF actuel si tu veux un PDF depuis l’app,
        # mais l’XLSX reste la source fidèle au modèle.
        pdf_bytes = _pdf_ramasse_from_template(
            _today_paris(), date_ramasse,
            df_calc[display_cols],
            {"cartons": tot_cartons, "palettes": tot_palettes, "poids": tot_poids},
        )
        st.download_button(
            "⬇️ Télécharger le PDF (option)",
            data=pdf_bytes,
            file_name=f"BL_enlevements_{_today_paris().strftime('%Y%m%d')}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
