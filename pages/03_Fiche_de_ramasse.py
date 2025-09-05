# pages/03_Fiche_de_ramasse.py
from __future__ import annotations
import os, re, math, io, datetime as dt
import pandas as pd
import streamlit as st
from dateutil.tz import gettz
from fpdf import FPDF

from common.design import apply_theme, section, kpi
from common.data import get_paths

# ===== Réglages =====
# Chemin du fichier Excel de référence (mapping références / conditionnements / palettes / poids)
# -> Mets ici le chemin du fichier que tu as joint dans le repo :
REF_EXCEL_PATH = "assets/LOG_EN_001_01 BL enlèvements Sofripa-2.xlsx"  # adapte si besoin



# Valeurs de secours si le fichier de référence est introuvable
FALLBACK_BY_FORMAT = {
    "12x33": {"cartons_par_palette": 108, "poids_palette_kg": 820},
    "6x75":  {"cartons_par_palette": 84,  "poids_palette_kg": 600},
    "4x75":  {"cartons_par_palette": 100, "poids_palette_kg": 500},
}

COMPANY_HEADER = {
    "name": "FERMENT STATION",
    "addr1": "Carré Ivry Bâtiment D2",
    "addr2": "47 rue Ernest Renan",
    "addr3": "94200 Ivry-sur-Seine - FRANCE",
    "phone": "Tél : 0967504647",
    "site": "Site : https://www.symbiose-kefir.fr",
    "bio": "Produits issus de l'Agriculture Biologique certifié par FR-BIO-01",
    "dest_label": "DESTINATAIRE : SOFRIPA",
    "dest_addr": "ZAC du Haut de Wissous II,\nRue Hélène Boucher, 91320 Wissous",
}

# ===== Helpers =====
def _tz_today_paris() -> dt.date:
    # DATE DE CREATION demandée = date du clic (Europe/Paris)
    now = dt.datetime.now(gettz("Europe/Paris"))
    return now.date()

def _parse_format_from_text(s: str) -> str | None:
    """Détecte '12x33', '6x75' ou '4x75' à partir d'un libellé produit/stock."""
    if not s:
        return None
    t = str(s)
    # tolère x ou × et espaces
    t = t.replace("×", "x").replace("X", "x")
    m = re.search(r"(\d+)\s*x\s*(\d+)\s*(?:cl|cL|CL)", t)
    if m:
        nb = int(m.group(1)); vol = int(m.group(2))
        if nb == 12 and vol in (33, 33_): return "12x33"
        if nb == 6 and vol == 75:         return "6x75"
        if nb == 4 and vol == 75:         return "4x75"
    # quelques variantes fréquentes
    if re.search(r"12\s*x?\s*33", t): return "12x33"
    if re.search(r"6\s*x?\s*75",  t): return "6x75"
    if re.search(r"4\s*x?\s*75",  t): return "4x75"
    return None

def _load_reference_table(path_xlsx: str) -> pd.DataFrame:
    """
    Charge un mapping 'format' → (référence, cartons/palette, poids palette).
    - Si l'Excel n'a pas les colonnes attendues, on tente de les retrouver par heuristique.
    - On garde une ligne par (référence, format). S'il y en a plusieurs, on prend la 1re.
    """
    if not os.path.exists(path_xlsx):
        return pd.DataFrame()

    # concatène toutes les feuilles pour être tolérant
    xls = pd.ExcelFile(path_xlsx)
    frames = []
    for sh in xls.sheet_names:
        try:
            frames.append(pd.read_excel(xls, sheet_name=sh))
        except Exception:
            pass
    if not frames:
        return pd.DataFrame()
    df = pd.concat(frames, ignore_index=True)

    # normalise noms colonnes
    lower = {c: str(c).strip().lower() for c in df.columns}
    df.columns = [lower[c] for c in df.columns]

    # essaie de retrouver les champs clés
    # - 'référence' numérique
    ref_col = next((c for c in df.columns if "réfé" in c or "ref" == c or c == "reference"), None)
    if ref_col is None:
        # fabrique depuis la 1re suite de chiffres 5-7 dans une colonne texte
        text_col = next((c for c in df.columns if "produit" in c or "libell" in c or "désig" in c or "designation" in c), None)
        if text_col and text_col in df.columns:
            df["__ref"] = df[text_col].astype(str).str.extract(r"(\d{5,7})", expand=False)
        else:
            # sinon cherche dans n'importe quelle colonne texte
            any_text = None
            for c in df.columns:
                if df[c].dtype == object:
                    any_text = c; break
            if any_text:
                df["__ref"] = df[any_text].astype(str).str.extract(r"(\d{5,7})", expand=False)
        ref_col = "__ref" if "__ref" in df.columns else None

    # - poids palette
    poids_col = next((c for c in df.columns if "poids" in c and "palette" in c), None)
    # - qté palettes ou cartons/palette
    ctp_col = next((c for c in df.columns if ("carton" in c and "palette" in c) or c in ("ctp","cartons/palette")), None)
    # - libellé produit (pour extraire le format)
    prod_col = next((c for c in df.columns if "produit" in c or "libell" in c or "désig" in c or "designation" in c), None)

    if not ref_col or (not poids_col and not ctp_col) or not prod_col:
        return pd.DataFrame()

    out = df[[ref_col, prod_col] + [c for c in (ctp_col, poids_col) if c and c in df.columns]].copy()
    out.columns = ["reference", "produit"] + [("cartons_par_palette" if c == ctp_col else "poids_palette_kg") for c in (ctp_col, poids_col) if c]
    out["format"] = out["produit"].apply(_parse_format_from_text)
    out = out.dropna(subset=["reference", "format"])
    # nettoie types
    out["reference"] = out["reference"].astype(str).str.strip()
    if "cartons_par_palette" in out.columns:
        out["cartons_par_palette"] = pd.to_numeric(out["cartons_par_palette"], errors="coerce")
    if "poids_palette_kg" in out.columns:
        out["poids_palette_kg"] = pd.to_numeric(out["poids_palette_kg"], errors="coerce")
    # garde 1 ligne / (reference, format)
    out = out.sort_index().drop_duplicates(subset=["reference","format"], keep="first").reset_index(drop=True)
    return out

def _lookup_ref_and_specs(prod_label: str, ref_table: pd.DataFrame) -> dict:
    fmt = _parse_format_from_text(prod_label)
    specs = {"reference": "", "format": fmt, "cartons_par_palette": None, "poids_palette_kg": None}
    if fmt is None:
        return specs

    # 1) essaie Excel
    if isinstance(ref_table, pd.DataFrame) and not ref_table.empty:
        # on ne sait pas la clef exacte → prend n'importe quelle ligne du même format
        cand = ref_table[ref_table["format"] == fmt]
        if not cand.empty:
            row = cand.iloc[0]
            specs["reference"] = str(row.get("reference", "") or "")
            specs["cartons_par_palette"] = row.get("cartons_par_palette")
            specs["poids_palette_kg"] = row.get("poids_palette_kg")

    # 2) fallback par format
    if pd.isna(specs["cartons_par_palette"]) or not specs["cartons_par_palette"]:
        specs["cartons_par_palette"] = FALLBACK_BY_FORMAT.get(fmt, {}).get("cartons_par_palette")
    if pd.isna(specs["poids_palette_kg"]) or not specs["poids_palette_kg"]:
        specs["poids_palette_kg"] = FALLBACK_BY_FORMAT.get(fmt, {}).get("poids_palette_kg")

    return specs

def _ceil_div(a: int, b: int) -> int:
    return int(math.ceil(float(a) / float(b))) if (a and b) else 0

# ======= UI =======
apply_theme("Fiche de ramasse — Ferment Station", "🚚")
section("Fiche de ramasse", "🚚")

# Besoin du fichier + de la sauvegarde de production
if "df_raw" not in st.session_state:
    st.warning("Aucun fichier chargé. Va dans **Accueil** pour déposer l'Excel, puis reviens.")
    st.stop()

sp = st.session_state.get("saved_production")
if not sp or "df_min" not in sp or "ddm" not in sp:
    st.info("Aucune production sauvegardée. Va dans l’onglet **Production** → « 💾 Sauvegarder cette production » d’abord.")
    st.stop()

df_min_saved = sp["df_min"].copy()
ddm_saved = dt.date.fromisoformat(sp["ddm"])  # DDM identique à l’onglet Production

# Options produits (issus du tableau affiché)
produits_options = df_min_saved["Produit"].astype(str).dropna().unique().tolist()
produits_options = sorted(produits_options)

# ---- Sidebar paramètres ----
with st.sidebar:
    st.header("Paramètres fiche de ramasse")
    date_creation = _tz_today_paris()  # automatique
    date_ramasse = st.date_input("Date de ramasse", value=date_creation)

    st.caption(f"DATE DE CRÉATION : **{date_creation.strftime('%d/%m/%Y')}**")
    st.caption(f"DDM (depuis Production) : **{ddm_saved.strftime('%d/%m/%Y')}**")

# ---- Sélection des produits & saisies ----
st.subheader("Sélection des produits")
selection = st.multiselect(
    "Produits à inclure dans la fiche",
    options=produits_options,
    default=produits_options[:1] if produits_options else [],
)

if not selection:
    st.info("Sélectionne au moins un produit.")
    st.stop()

# charge table de référence (si dispo)
ref_table = _load_reference_table(REF_EXCEL_PATH)

# construit table éditable
rows = []
for prod in selection:
    specs = _lookup_ref_and_specs(prod, ref_table)
    rows.append({
        "Produit": prod,
        "Référence": specs["reference"],
        "DDM": ddm_saved.strftime("%d/%m/%Y"),
        "Quantité cartons": 0,
        "Quantité bouteilles": 0,
        "Format": specs["format"] or "",
        "Cartons/palette": int(specs["cartons_par_palette"] or 0),
        "Poids palette (kg)": float(specs["poids_palette_kg"] or 0.0),
    })

df_edit = pd.DataFrame(rows, columns=[
    "Référence","Produit","DDM","Quantité cartons","Quantité bouteilles","Format","Cartons/palette","Poids palette (kg)"
])

st.caption("Renseigne manuellement les **cartons** et **bouteilles**. Les **palettes** et **poids** seront calculés.")
edited = st.data_editor(
    df_edit,
    key="ramasse_editor",
    use_container_width=True,
    hide_index=True,
    column_config={
        "Quantité cartons": st.column_config.NumberColumn(min_value=0, step=1),
        "Quantité bouteilles": st.column_config.NumberColumn(min_value=0, step=1),
        "Cartons/palette": st.column_config.NumberColumn(min_value=0, step=1, disabled=True),
        "Poids palette (kg)": st.column_config.NumberColumn(min_value=0.0, step=1.0, disabled=True),
    }
)

# calc automatiques
def _calc_totaux(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    df = df.copy()
    df["Quantité palettes"] = df.apply(
        lambda r: _ceil_div(int(r.get("Quantité cartons") or 0), int(r.get("Cartons/palette") or 0)),
        axis=1
    )
    df["Poids palettes (kg)"] = (pd.to_numeric(df.get("Quantité palettes"), errors="coerce").fillna(0)
                                 * pd.to_numeric(df.get("Poids palette (kg)"), errors="coerce").fillna(0)).round(0).astype(int)
    tot = {
        "cartons": int(pd.to_numeric(df["Quantité cartons"], errors="coerce").fillna(0).sum()),
        "palettes": int(pd.to_numeric(df["Quantité palettes"], errors="coerce").fillna(0).sum()),
        "poids": int(pd.to_numeric(df["Poids palettes (kg)"], errors="coerce").fillna(0).sum()),
    }
    return df, tot

df_calc, totals = _calc_totaux(edited)

colA, colB, colC = st.columns(3)
with colA: kpi("Total cartons", f"{totals['cartons']:,}".replace(",", " "))
with colB: kpi("Total palettes", f"{totals['palettes']}")
with colC: kpi("Poids total (kg)", f"{totals['poids']:,}".replace(",", " "))

st.dataframe(
    df_calc[["Référence","Produit","DDM","Quantité cartons","Quantité palettes","Poids palettes (kg)"]],
    use_container_width=True, hide_index=True
)

# ===== Génération PDF =====
def _pdf_ramasse(company: dict, date_creation: dt.date, date_ramasse: dt.date, df_lines: pd.DataFrame, totals: dict) -> bytes:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=12)
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 8, company["name"], ln=1)
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 5, company["addr1"], ln=1)
    pdf.cell(0, 5, company["addr2"], ln=1)
    pdf.cell(0, 5, company["addr3"], ln=1)
    pdf.cell(0, 5, company["phone"], ln=1)
    pdf.cell(0, 5, company["site"], ln=1)
    pdf.ln(2)
    pdf.set_font("Helvetica", "I", 9)
    pdf.multi_cell(0, 4.5, company["bio"])
    pdf.ln(4)

    # dates
    pdf.set_font("Helvetica", "", 11)
    pdf.cell(0, 6, f"DATE DE CREATION : {date_creation.strftime('%d/%m/%Y')}", ln=1)
    pdf.cell(0, 6, f"DATE DE RAMMASSE : {date_ramasse.strftime('%d/%m/%Y')}", ln=1)  # orthographe volontaire pour calquer ton modèle
    pdf.ln(2)

    # destinataire
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 6, company["dest_label"], ln=1)
    pdf.set_font("Helvetica", "", 10)
    pdf.multi_cell(0, 5, company["dest_addr"])
    pdf.ln(3)

    # tableau
    headers = ["Référence", "Produit", "DDM", "Quantité cartons", "Quantité palettes", "Poids palettes (kg)"]
    widths  = [25, 75, 28, 28, 24, 28]
    pdf.set_font("Helvetica", "B", 10)
    for h, w in zip(headers, widths):
        pdf.cell(w, 8, h, border=1, align="C")
    pdf.ln(8)

    pdf.set_font("Helvetica", "", 10)
    for _, r in df_lines.iterrows():
        cells = [
            str(r.get("Référence") or ""),
            str(r.get("Produit") or ""),
            str(r.get("DDM") or ""),
            str(int(pd.to_numeric(r.get("Quantité cartons"), errors="coerce") or 0)),
            str(int(pd.to_numeric(r.get("Quantité palettes"), errors="coerce") or 0)),
            str(int(pd.to_numeric(r.get("Poids palettes (kg)"), errors="coerce") or 0)),
        ]
        # lignes
        for i, (txt, w) in enumerate(zip(cells, widths)):
            align = "C" if i != 1 else "L"
            pdf.cell(w, 8, txt, border=1, align=align)
        pdf.ln(8)

    # totals
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(widths[0] + widths[1] + widths[2], 8, "", border=1)           # colonnes vides
    pdf.cell(widths[3], 8, str(totals["cartons"]), border=1, align="C")
    pdf.cell(widths[4], 8, str(totals["palettes"]), border=1, align="C")
    pdf.cell(widths[5], 8, str(totals["poids"]), border=1, align="C")
    pdf.ln(12)

    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 8, "BON DE LIVRAISON", ln=1, align="R")

    return pdf.output(dest="S").encode("latin1")

st.markdown("---")
if st.button("🧾 Générer la fiche de ramasse (PDF)", use_container_width=True, type="primary"):
    # contrôle de base
    if df_calc["Quantité cartons"].sum() <= 0:
        st.error("Renseigne au moins une quantité de cartons > 0.")
    else:
        pdf_bytes = _pdf_ramasse(
            COMPANY_HEADER,
            _tz_today_paris(),
            date_ramasse,
            df_calc[["Référence","Produit","DDM","Quantité cartons","Quantité palettes","Poids palettes (kg)"]],
            totals,
        )
        fname = f"BL_enlevements_{_tz_today_paris().strftime('%Y%m%d')}.pdf"
        st.download_button(
            "📥 Télécharger le PDF",
            data=pdf_bytes,
            file_name=fname,
            mime="application/pdf",
            use_container_width=True,
        )
