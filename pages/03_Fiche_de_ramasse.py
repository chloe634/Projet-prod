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
    # logo dimensionné en largeur fixe -> évite le chevauchement
    logo_w = 36
    logo_h = 18
    if os.path.exists(LOGO_PATH):
        try:
            pdf.image(LOGO_PATH, x=MARGIN_L, y=y, w=logo_w)  # largeur fixe
        except Exception:
            pass

    # bloc coordonnées à droite du logo (x = marge + logo + espace)
    x_text = MARGIN_L + logo_w + 8
    pdf.set_xy(x_text, y)
    pdf.set_font("Helvetica", "", 11)
    for line in COMPANY_LINES[:6]:
        pdf.cell(0, 5, _pdf_txt(line), ln=1)
    pdf.ln(1)
    pdf.set_font("Helvetica", "", 8)
    if len(COMPANY_LINES) > 6:
        pdf.cell(0, 4, _pdf_txt(COMPANY_LINES[6]), ln=1)

    # ajoute un espace vertical franc pour éviter l'underline parasite
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

    # colonne gauche
    pdf.set_font("Helvetica", "", 10)
    LBL_W = 46
    pdf.set_xy(box_x + 3, box_y + 12)
    pdf.cell(LBL_W, 6, _pdf_txt("DATE DE CREATION :"), ln=0)
    pdf.cell(35, 6, _pdf_txt(date_creation.strftime("%d/%m/%Y")), ln=1)

    pdf.set_xy(box_x + 3, box_y + 18)
    pdf.cell(LBL_W, 6, _pdf_txt("DATE DE RAMMASSE :"), ln=0)  # conforme au modèle
    pdf.cell(35, 6, _pdf_txt(date_ramasse.strftime("%d/%m/%Y")), ln=1)

    # colonne droite (destinataire)
    right_x = box_x + 95  # point de départ visuel proche du milieu
    pdf.set_xy(right_x, box_y + 12)
    pdf.cell(32, 6, _pdf_txt("DESTINATAIRE :"), ln=1)
    pdf.set_xy(right_x, box_y + 18)
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 6, _pdf_txt(DEST_TITLE), ln=1)
    pdf.set_font("Helvetica", "", 10)
    for i, l in enumerate(DEST_LINES):
        pdf.set_xy(right_x, box_y + 24 + i * 6)
        pdf.cell(0, 6, _pdf_txt(l), ln=1)

    # place le curseur SOUS la boîte + marge de respiration
    pdf.set_y(box_y + box_h + 8)

    # ---------- TABLEAU ----------
    # Largeurs calibrées sur 186 mm (≈ rendu de ton modèle)
    # 25 + 85 + 20 + 18 + 18 + 20 = 186
    col_w = [25, 85, 20, 18, 18, 20]
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

    # ligne TOTAL
    pdf.set_font("Helvetica", "B", 10)
    total_label_w = col_w[0] + col_w[1] + col_w[2]
    pdf.cell(total_label_w, LINE_H, _pdf_txt("TOTAL"), border=1, align="R")
    pdf.cell(col_w[3], LINE_H, _pdf_txt(totals["cartons"]), border=1, align="C")
    pdf.cell(col_w[4], LINE_H, _pdf_txt(totals["palettes"]), border=1, align="C")
    pdf.cell(col_w[5], LINE_H, _pdf_txt(totals["poids"]), border=1, align="C")

    # ---------- sortie binaire (toujours bytes) ----------
    buf = pdf.output(dest="S")
    if isinstance(buf, str):
        data = buf.encode("latin-1", "ignore")
    elif isinstance(buf, bytearray):
        data = bytes(buf)
    else:
        data = buf
    return data




st.markdown("---")
if st.button("🧾 Générer la fiche de ramasse (PDF)", use_container_width=True, type="primary"):
    if tot_cartons <= 0:
        st.error("Renseigne au moins une **Quantité cartons** > 0.")
    else:
        pdf_bytes = _pdf_ramasse(
            _today_paris(),
            date_ramasse,
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
