# pages/03_Fiche_de_ramasse.py
from __future__ import annotations
import os, re, datetime as dt
import unicodedata
import pandas as pd
import streamlit as st
from dateutil.tz import gettz

from common.design import apply_theme, section, kpi
# au lieu de: from common.xlsx_fill import fill_bl_enlevements_xlsx
import importlib
import common.xlsx_fill as _xlsx_fill
importlib.reload(_xlsx_fill)
from common.xlsx_fill import fill_bl_enlevements_xlsx, build_bl_enlevements_pdf


# ------------------------------------------------------------------
# Réglages
# ------------------------------------------------------------------
INFO_CSV_PATH = "info_FDR.csv"   # ton CSV catalogue (Code-barre, Poids, ...)
TEMPLATE_XLSX_PATH = "assets/BL_enlevements_Sofripa.xlsx"

DEST_TITLE = "SOFRIPA"
DEST_LINES = [
    "ZAC du Haut de Wissous II,",
    "Rue Hélène Boucher, 91320 Wissous",
]

# ------------------------------------------------------------------
# Utils
# ------------------------------------------------------------------
def _today_paris() -> dt.date:
    return dt.datetime.now(gettz("Europe/Paris")).date()

def _strip_accents(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    return "".join(ch for ch in s if not unicodedata.combining(ch))

def _canon(s: str) -> str:
    s = _strip_accents(str(s or "")).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def _format_from_stock(stock_txt: str) -> str | None:
    """
    Détecte 12x33 / 6x75 / 4x75 dans un libellé de Stock.
    """
    if not stock_txt:
        return None
    s = str(stock_txt).lower().replace("×", "x").replace("\u00a0", " ")

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

@st.cache_data(show_spinner=False)
def _load_catalog(path: str) -> pd.DataFrame:
    """
    Lit info_FDR.csv et prépare colonnes auxiliaires pour le matching.
    - normalise Poids (virgule -> point)
    - prépare Format normalisé et formes canonisées de Produit/Désignation
    """
    import pandas as pd, os, re
    if not os.path.exists(path):
        return pd.DataFrame(columns=["Produit","Format","Désignation","Code-barre","Poids"])

    df = pd.read_csv(path, encoding="utf-8")
    for c in ["Produit","Format","Désignation","Code-barre"]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()

    # Poids: "7,23" -> "7.23" puis numeric
    if "Poids" in df.columns:
        df["Poids"] = (
            df["Poids"]
            .astype(str)
            .str.replace(",", ".", regex=False)
        )
        df["Poids"] = pd.to_numeric(df["Poids"], errors="coerce")

    # Format: "12x33cl" -> "12x33", "6x75cl" -> "6x75"
    df["_format_norm"] = df.get("Format","").astype(str).str.lower()
    df["_format_norm"] = (
        df["_format_norm"]
        .str.replace("cl", "", regex=False)
        .str.replace(" ", "", regex=False)
    )

    # Canon pour Produit / Désignation
    df["_canon_prod"] = df.get("Produit","").map(_canon)
    # on retire tout ce qui est entre parenthèses, puis canon
    df["_canon_des"]  = df.get("Désignation","").map(lambda s: _canon(re.sub(r"\(.*?\)", "", s)))

    return df


def _csv_lookup(catalog: pd.DataFrame, gout_canon: str, fmt_label: str) -> tuple[str, float] | None:
    """
    Retourne (référence_6_chiffres, poids_carton) en matchant :
      - format (12x33 / 6x75 / 4x75)
      - + goût canonisé (ex: 'mangue passion') contre Produit/Désignation du CSV
    """
    if catalog is None or catalog.empty or not fmt_label:
        return None

    fmt_norm = fmt_label.lower().replace("cl","").replace(" ", "")
    g_can = _canon(gout_canon)

    # filtre format d'abord
    cand = catalog[catalog["_format_norm"].str.contains(fmt_norm, na=False)]
    if cand.empty:
        return None

    # 1) match strict sur Produit canonisé
    m1 = cand[cand["_canon_prod"] == g_can]
    if m1.empty:
        # 2) sinon, on vérifie que tous les tokens du goût sont dans la désignation canonisée
        toks = [t for t in g_can.split() if t]
        def _contains_all(s):
            s2 = str(s or "")
            return all(t in s2 for t in toks)
        m1 = cand[cand["_canon_des"].map(_contains_all)]

    if m1.empty:
        # en dernier recours, on prend juste le premier du bon format
        m1 = cand

    row = m1.iloc[0]
    code = re.sub(r"\D+", "", str(row.get("Code-barre","")))
    ref6 = code[-6:] if len(code) >= 6 else code
    poids = float(row.get("Poids") or 0.0)
    return (ref6, poids) if ref6 else None


    row = cand.iloc[0]
    code = re.sub(r"\D+", "", str(row.get("Code-barre","")))
    ref6 = code[-6:] if len(code) >= 6 else code
    poids = float(row.get("Poids") or 0.0)
    return ref6, poids

# ------------------------------------------------------------------
# UI
# ------------------------------------------------------------------
apply_theme("Fiche de ramasse — Ferment Station", "🚚")
section("Fiche de ramasse", "🚚")

# Besoin de la production sauvegardée depuis la page "Production"
if "saved_production" not in st.session_state or "df_min" not in st.session_state["saved_production"]:
    st.warning("Va d’abord dans **Production** et clique **💾 Sauvegarder cette production**.")
    st.stop()

sp = st.session_state["saved_production"]
df_min_saved: pd.DataFrame = sp["df_min"].copy()
ddm_saved = dt.date.fromisoformat(sp["ddm"]) if "ddm" in sp else _today_paris()

# 1) Options dérivées de la prod sauvegardée (goût + format)
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
        "prod_hint": str(r.get("Produit") or "").strip(),  # pour matcher le CSV
    })

if not opts_rows:
    st.error("Impossible de détecter les **formats** (12x33, 6x75, 4x75) dans la production sauvegardée.")
    st.stop()

opts_df = pd.DataFrame(opts_rows).sort_values(by="label").reset_index(drop=True)

# 2) Catalogue CSV
catalog = _load_catalog(INFO_CSV_PATH)
if catalog.empty:
    st.warning("⚠️ `info_FDR.csv` introuvable ou vide — références/poids non calculables.")

# 3) Sidebar : dates
with st.sidebar:
    st.header("Paramètres")
    date_creation = _today_paris()
    date_ramasse = st.date_input("Date de ramasse", value=date_creation)
    st.caption(f"DATE DE CRÉATION : **{date_creation.strftime('%d/%m/%Y')}**")
    st.caption(f"DDM (depuis Production) : **{ddm_saved.strftime('%d/%m/%Y')}**")

# 4) Sélection utilisateur
st.subheader("Sélection des produits")
selection_labels = st.multiselect(
    "Produits à inclure (Goût — Format)",
    options=opts_df["label"].tolist(),
    default=opts_df["label"].tolist(),
)

if not selection_labels:
    st.info("Sélectionne au moins un produit.")
    st.stop()

# 5) Prépare la table éditable (Référence + Poids issus du CSV)
meta_by_label = {}
rows = []
for lab in selection_labels:
    row_opt = opts_df.loc[opts_df["label"] == lab].iloc[0]
    gout     = row_opt["gout"]          # <-- on utilise le GOÛT canonisé
    fmt      = row_opt["format"]

    ref = ""; poids_carton = 0.0
    lk = _csv_lookup(catalog, gout, fmt)  # <-- lookup par goût + format
    if lk:
        ref, poids_carton = lk
    meta_by_label[lab] = {"_format": fmt, "_poids_carton": poids_carton, "_reference": ref}

    rows.append({
        "Référence": ref,
        "Produit (goût + format)": lab.replace(" — ", " - "),
        "DDM": ddm_saved.strftime("%d/%m/%Y"),
        "Quantité cartons": 0,
        "Quantité palettes": 0,
        "Poids palettes (kg)": 0,
    })

display_cols = ["Référence","Produit (goût + format)","DDM","Quantité cartons","Quantité palettes","Poids palettes (kg)"]
base_df = pd.DataFrame(rows, columns=display_cols)

st.caption("Renseigne **Quantité cartons** et, si besoin, **Quantité palettes**. Le **poids** se calcule automatiquement (cartons × poids/carton du CSV).")
edited = st.data_editor(
    base_df,
    key="ramasse_editor_xlsx_v1",
    use_container_width=True,
    hide_index=True,
    column_config={
        "Quantité cartons":   st.column_config.NumberColumn(min_value=0, step=1),
        "Quantité palettes":  st.column_config.NumberColumn(min_value=0, step=1),
        "Poids palettes (kg)": st.column_config.NumberColumn(disabled=True, format="%.0f"),
    },
)

# 6) Calcul poids = cartons × poids/carton
def _apply_calculs(df_disp: pd.DataFrame) -> pd.DataFrame:
    out = df_disp.copy()
    poids = []
    for _, r in out.iterrows():
        # On retrouve la clé label côté meta, avec ou sans remplacement du tiret
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

st.dataframe(df_calc[display_cols], use_container_width=True, hide_index=True)

# 7) Téléchargement XLSX (remplissage du modèle)
st.markdown("---")
if st.button("📄 Télécharger la fiche (XLSX, modèle Sofripa)", use_container_width=True, type="primary"):
    if tot_cartons <= 0:
        st.error("Renseigne au moins une **Quantité cartons** > 0.")
    elif not os.path.exists(TEMPLATE_XLSX_PATH):
        st.error(f"Modèle Excel introuvable : `{TEMPLATE_XLSX_PATH}`")
    else:
        try:
            xlsx_bytes = fill_bl_enlevements_xlsx(
                template_path=TEMPLATE_XLSX_PATH,
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

# 7-bis) Téléchargement PDF (rendu propre via fpdf2)
if st.button("🧾 Télécharger la version PDF", use_container_width=True):
    if tot_cartons <= 0:
        st.error("Renseigne au moins une **Quantité cartons** > 0.")
    else:
        try:
            # <<< C’EST ICI que va ton appel >>>
            pdf_bytes = build_bl_enlevements_pdf(
                date_creation=_today_paris(),
                date_ramasse=date_ramasse,
                destinataire_title=DEST_TITLE,
                destinataire_lines=DEST_LINES,
                df_lines=df_calc[display_cols],
                col2_header=DEST_LINES[-1] if DEST_LINES else "Produit",  # <- en-tête 2
            )

            fname_pdf = f"BL_enlevements_{_today_paris().strftime('%Y%m%d')}.pdf"
            st.download_button(
                "⬇️ Télécharger le PDF",
                data=pdf_bytes,
                file_name=fname_pdf,
                mime="application/pdf",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"Erreur lors de la génération du PDF : {e}")
