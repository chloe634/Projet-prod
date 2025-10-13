# pages/03_Fiche_de_ramasse.py
from __future__ import annotations

import os, re, datetime as dt, unicodedata, mimetypes
import pandas as pd
import streamlit as st
from dateutil.tz import gettz

from common.design import apply_theme, section, kpi
import importlib
import common.xlsx_fill as _xlsx_fill
importlib.reload(_xlsx_fill)
from common.xlsx_fill import fill_bl_enlevements_xlsx, build_bl_enlevements_pdf

# === EMAIL ===
import smtplib
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path

from common.storage import list_saved, load_snapshot

# ======================= Helpers email (secrets + fallback) ===================
# tomllib (Py 3.11+) ou tomli (Py 3.8–3.10)
try:
    import tomllib as _toml
except Exception:
    import tomli as _toml  # ➜ ajoute 'tomli' dans requirements.txt si besoin

def _load_email_secrets_fallback() -> dict:
    """
    Priorités:
      1) st.secrets["email"] (Cloud / local)
      2) <racine>/.streamlit/secrets.toml
      3) <racine>/streamlit/secrets.toml (compat ancien dossier)
    """
    if "email" in st.secrets:
        return dict(st.secrets.get("email", {}))

    try:
        proj_root = Path(__file__).resolve().parents[1]
    except Exception:
        proj_root = Path(os.getcwd())

    for p in [proj_root / ".streamlit" / "secrets.toml",
              proj_root / "streamlit" / "secrets.toml"]:
        try:
            if p.exists():
                with open(p, "rb") as f:
                    data = _toml.load(f)
                if isinstance(data, dict) and "email" in data:
                    return dict(data["email"] or {})
        except Exception:
            continue
    return {}

def _get_email_cfg():
    cfg = _load_email_secrets_fallback()
    required = ("host", "port", "user", "password")
    missing = [k for k in required if not str(cfg.get(k, "")).strip()]
    if missing:
        raise RuntimeError(
            "Secrets email manquants: " + ", ".join(missing) +
            " — place le bloc [email] dans Settings → Secrets (Cloud) ou .streamlit/secrets.toml (local)."
        )
    cfg.setdefault("sender", cfg["user"])
    rec = cfg.get("recipients", [])
    if isinstance(rec, str):
        rec = [x.strip() for x in rec.split(",") if x.strip()]
    cfg["recipients"] = rec
    return cfg

# =================== Envoi email (HTML + signature + images) ==================
def send_mail_with_pdf(
    pdf_bytes: bytes,
    filename: str,
    total_palettes: int,
    to_list: list[str],
    date_ramasse: dt.date,
    bcc_me: bool = True
):
    cfg = _get_email_cfg()
    sender = cfg["sender"]                  # = cfg["user"]
    from_value = formataddr(("Ferment Station – Logistique", sender))

    # Corps
    body_txt = f"""Bonjour,

Nous aurions besoin d’une ramasse pour demain.
Pour {total_palettes} palettes.

Merci,
Bon après-midi."""
    body_html = f"""<p>Bonjour,</p>
<p>Nous aurions besoin d’une ramasse pour demain.<br>
Pour <strong>{total_palettes}</strong> palettes.</p>
<p>Merci,<br>Bon après-midi.</p>"""

    # Signature (texte + HTML avec images inline)
    SIG_TXT = """--
Ferment Station
Producteur de boissons fermentées
26 Rue Robert Witchitz – 94200 Ivry-sur-Seine
09 71 22 78 95"""

    SIG_HTML = """
<hr style="border:none;border-top:1px solid #e5e7eb;margin:16px 0">
<div style="font:14px/1.5 -apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#111827">
  <div style="font-size:18px;font-weight:700">Ferment Station</div>
  <div style="font-weight:700;margin-top:2px">Producteur de boissons fermentées</div>
  <div style="margin-top:12px">26 Rue Robert Witchitz – 94200 Ivry-sur-Seine</div>
  <div><a href="tel:+33971227895" style="color:#2563eb;text-decoration:underline">09 71 22 78 95</a></div>
  <div style="margin-top:14px">
    <img src="cid:symbiose" alt="Symbiose" height="36" style="vertical-align:middle;margin-right:14px;border:0">
    <img src="cid:niko"     alt="Niko"     height="36" style="vertical-align:middle;border:0">
  </div>
</div>
"""

    msg = EmailMessage()
    msg["Subject"] = f"Demande de ramasse — {date_ramasse:%d/%m/%Y} — Ferment Station"
    msg["From"] = from_value
    msg["To"] = ", ".join(to_list)
    msg["Reply-To"] = sender
    msg["X-Priority"] = "1"                 # surtout pour Outlook
    msg["X-MSMail-Priority"] = "High"
    msg["Importance"] = "High"
    msg["X-App-Trace"] = "ferment-station/fiche-ramasse"

    # Texte + HTML (+ signature)
    msg.set_content(body_txt + "\n\n" + SIG_TXT)
    msg.add_alternative(body_html + SIG_HTML, subtype="html")

    # Images inline (CID) pour la signature — version minimisée (pas de filename)
    INLINE_IMAGES = {
        "symbiose": "assets/signature/logo_symbiose.png",
        "niko":     "assets/signature/NIKO_Logo.png",
    }
    html_part = msg.get_payload()[-1]  # partie HTML (text/html)
    
    for cid, path in INLINE_IMAGES.items():
        if not os.path.exists(path):
            st.caption(f"⚠️ Signature: fichier introuvable → {path}")
            continue
        try:
            with open(path, "rb") as f:
                data = f.read()
            if not data:
                st.caption(f"⚠️ Signature: fichier vide → {path}")
                continue
    
            related = html_part.add_related(
                data,
                maintype="image",
                subtype="png",             # force PNG
                cid=f"<{cid}>",            # référence via src="cid:cid"
                # ❌ pas de filename pour éviter d’être listé comme PJ
            )
            # disposition explicite en inline
            related.add_header("Content-Disposition", "inline")
            # astuce utilisée par Gmail pour associer CID ↔ image
            related.add_header("X-Attachment-Id", cid)
        except Exception as e:
            st.caption(f"⚠️ Signature: erreur sur {path} → {e}")


    # Pièce jointe PDF
    msg.add_attachment(pdf_bytes, maintype="application", subtype="pdf", filename=filename)

    # BCC vers l’expéditeur (vérif de distribution)
    bcc_list = [sender] if bcc_me else []

    # Envoi (465 SSL ou 587 STARTTLS)
    if int(cfg["port"]) == 465:
        import ssl
        with smtplib.SMTP_SSL(cfg["host"], 465, context=ssl.create_default_context()) as s:
            s.login(cfg["user"], cfg["password"])      # ✅ dict, pas fonction
            refused = s.send_message(msg, from_addr=sender, to_addrs=to_list + bcc_list)
    else:
        with smtplib.SMTP(cfg["host"], int(cfg["port"])) as s:
            s.ehlo(); s.starttls(); s.ehlo()
            s.login(cfg["user"], cfg["password"])
            refused = s.send_message(msg, from_addr=sender, to_addrs=to_list + bcc_list)

    return refused  # {} si tout accepté par le serveur

# ================================ Réglages ====================================
INFO_CSV_PATH = "info_FDR.csv"
TEMPLATE_XLSX_PATH = "assets/BL_enlevements_Sofripa.xlsx"

DEST_TITLE = "SOFRIPA"
DEST_LINES = [
    "ZAC du Haut de Wissous II,",
    "Rue Hélène Boucher, 91320 Wissous",
]

# ================================ Utils =======================================
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
    """
    if not os.path.exists(path):
        return pd.DataFrame(columns=["Produit","Format","Désignation","Code-barre","Poids"])

    df = pd.read_csv(path, encoding="utf-8")
    for c in ["Produit","Format","Désignation","Code-barre"]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()

    if "Poids" in df.columns:
        df["Poids"] = (
            df["Poids"].astype(str).str.replace(",", ".", regex=False)
        )
        df["Poids"] = pd.to_numeric(df["Poids"], errors="coerce")

    df["_format_norm"] = df.get("Format","").astype(str).str.lower()
    df["_format_norm"] = df["_format_norm"].str.replace("cl","", regex=False).str.replace(" ", "", regex=False)

    df["_canon_prod"] = df.get("Produit","").map(_canon)
    df["_canon_des"]  = df.get("Désignation","").map(lambda s: _canon(re.sub(r"\(.*?\)", "", s)))

    return df

def _csv_lookup(catalog: pd.DataFrame, gout_canon: str, fmt_label: str) -> tuple[str, float] | None:
    """
    Retourne (référence_6_chiffres, poids_carton) en matchant :
      - format (12x33 / 6x75 / 4x75)
      - + goût canonisé
    """
    if catalog is None or catalog.empty or not fmt_label:
        return None

    fmt_norm = fmt_label.lower().replace("cl","").replace(" ", "")
    g_can = _canon(gout_canon)

    cand = catalog[catalog["_format_norm"].str.contains(fmt_norm, na=False)]
    if cand.empty:
        return None

    m1 = cand[cand["_canon_prod"] == g_can]
    if m1.empty:
        toks = [t for t in g_can.split() if t]
        def _contains_all(s):
            s2 = str(s or "")
            return all(t in s2 for t in toks)
        m1 = cand[cand["_canon_des"].map(_contains_all)]
    if m1.empty:
        m1 = cand

    row = m1.iloc[0]
    code = re.sub(r"\D+", "", str(row.get("Code-barre","")))
    ref6 = code[-6:] if len(code) >= 6 else code
    poids = float(row.get("Poids") or 0.0)
    return (ref6, poids) if ref6 else None

def _build_opts_from_saved(df_min_saved: pd.DataFrame) -> pd.DataFrame:
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
        opts_rows.append({"label": f"{gout} — {fmt}", "gout": gout, "format": fmt})
    return pd.DataFrame(opts_rows).sort_values(by="label").reset_index(drop=True)

def _build_opts_from_catalog(catalog: pd.DataFrame) -> pd.DataFrame:
    if catalog is None or catalog.empty:
        return pd.DataFrame(columns=["label","gout","format"])
    rows, seen = [], set()
    for _, r in catalog.iterrows():
        gout = str(r.get("Produit","")).strip()     # <- non canonique
        fmt  = str(r.get("Format","")).strip()
        if not (gout and fmt):
            continue
        key = (gout, fmt)
        if key in seen:
            continue
        seen.add(key)
        rows.append({"label": f"{gout} — {fmt}", "gout": gout, "format": fmt})
    return pd.DataFrame(rows).sort_values(by="label").reset_index(drop=True)


# ================================== UI =======================================
apply_theme("Fiche de ramasse — Ferment Station", "🚚")
section("Fiche de ramasse", "🚚")

# NEW — choix de la source des produits
source_mode = st.radio(
    "Source des produits pour la fiche",
    options=["Proposition sauvegardée", "Sélection manuelle"],
    horizontal=True
)

# 2) Catalogue CSV (on le charge tôt car utile en manuel et pour lookup)
catalog = _load_catalog(INFO_CSV_PATH)
if catalog.empty:
    st.warning("⚠️ `info_FDR.csv` introuvable ou vide — références/poids non calculables.")

# === Proposition sauvegardée requise uniquement si on est sur ce mode ===
if source_mode == "Proposition sauvegardée":
    if ("saved_production" not in st.session_state) or ("df_min" not in st.session_state.get("saved_production", {})):
        st.warning(
            "Va d’abord dans **Production** et clique **💾 Sauvegarder cette production** "
            "ou charge une proposition depuis la mémoire longue ci-dessous."
        )

        saved = list_saved()
        if saved:
            labels = [f"{it['name']} — ({it.get('semaine_du','?')})" for it in saved]
            sel = st.selectbox("Charger une proposition enregistrée", options=labels)
            if st.button("▶️ Charger cette proposition", use_container_width=True):
                picked_name = saved[labels.index(sel)]["name"]
                sp_loaded = load_snapshot(picked_name)
                if sp_loaded and sp_loaded.get("df_min") is not None:
                    st.session_state["saved_production"] = sp_loaded
                    st.success(f"Chargé : {picked_name}")
                    st.rerun()
                else:
                    st.error("Proposition invalide (df_min manquant).")
        st.stop()

# === À partir d’ici :
# - en mode Proposition, on a une prod en session
# - en mode Manuel, on n’en a potentiellement pas (et ce n’est pas bloquant)
if source_mode == "Proposition sauvegardée":
    sp = st.session_state["saved_production"]
    df_min_saved: pd.DataFrame = sp["df_min"].copy()
    ddm_saved = dt.date.fromisoformat(sp["ddm"]) if "ddm" in sp else _today_paris()
    opts_df = _build_opts_from_saved(df_min_saved)
else:
    df_min_saved = None
    ddm_saved = _today_paris()  # NEW — valeur par défaut si pas de prod
    opts_df = _build_opts_from_catalog(catalog)

if opts_df.empty:
    if source_mode == "Proposition sauvegardée":
        st.error("Impossible de détecter les **formats** (12x33, 6x75, 4x75) dans la production sauvegardée.")
    else:
        st.error("Aucun produit détecté dans `info_FDR.csv` (colonnes « Produit » et « Format » requises).")
    st.stop()

# === À partir d’ici on a bien une prod en session ===
sp = st.session_state["saved_production"]
df_min_saved: pd.DataFrame = sp["df_min"].copy()
ddm_saved = dt.date.fromisoformat(sp["ddm"]) if "ddm" in sp else _today_paris()

# Choix de la source (place-le tout de suite après section(...))
source_mode = st.radio(
    "Source des produits pour la fiche",
    options=["Proposition sauvegardée", "Sélection manuelle"],
    horizontal=True,
    key="ramasse_source_mode"
)

# Charger le catalogue tôt (utilisé en manuel et pour lookup)
catalog = _load_catalog(INFO_CSV_PATH)
if catalog.empty:
    st.warning("⚠️ `info_FDR.csv` introuvable ou vide — références/poids non calculables.")

# Si 'Proposition sauvegardée' => on exige la prod, sinon on n’arrête pas le flux
if source_mode == "Proposition sauvegardée":
    if ("saved_production" not in st.session_state) or ("df_min" not in st.session_state.get("saved_production", {})):
        st.warning("Va d’abord dans **Production** et clique **💾 Sauvegarder cette production** "
                   "ou charge une proposition depuis la mémoire longue ci-dessous.")
        saved = list_saved()
        if saved:
            labels = [f"{it['name']} — ({it.get('semaine_du','?')})" for it in saved]
            sel = st.selectbox("Charger une proposition enregistrée", options=labels)
            if st.button("▶️ Charger cette proposition", use_container_width=True):
                picked_name = saved[labels.index(sel)]["name"]
                sp_loaded = load_snapshot(picked_name)
                if sp_loaded and sp_loaded.get("df_min") is not None:
                    st.session_state["saved_production"] = sp_loaded
                    st.success(f"Chargé : {picked_name}")
                    st.rerun()
                else:
                    st.error("Proposition invalide (df_min manquant).")
        st.stop()

# Construire opts_df selon le mode
if source_mode == "Proposition sauvegardée":
    sp = st.session_state["saved_production"]
    df_min_saved: pd.DataFrame = sp["df_min"].copy()
    ddm_saved = dt.date.fromisoformat(sp["ddm"]) if "ddm" in sp else _today_paris()
    opts_df = _build_opts_from_saved(df_min_saved)
else:
    df_min_saved = None
    ddm_saved = _today_paris()
    opts_df = _build_opts_from_catalog(catalog)

if opts_df.empty:
    st.error("Aucun produit détecté pour ce mode (vérifie `info_FDR.csv` en manuel).")
    st.stop()

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
    default=opts_df["label"].tolist() if source_mode == "Proposition sauvegardée" else [],
)

# 5) Table éditable
meta_by_label = {}
rows = []
for lab in selection_labels:
    row_opt = opts_df.loc[opts_df["label"] == lab].iloc[0]
    gout = row_opt["gout"]
    fmt  = row_opt["format"]
    ref = ""; poids_carton = 0.0
    lk = _csv_lookup(catalog, gout, fmt)
    if lk: ref, poids_carton = lk
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
    base_df, key="ramasse_editor_xlsx_v1", use_container_width=True, hide_index=True,
    column_config={
        "Quantité cartons":  st.column_config.NumberColumn(min_value=0, step=1),
        "Quantité palettes": st.column_config.NumberColumn(min_value=0, step=1),
        "Poids palettes (kg)": st.column_config.NumberColumn(disabled=True, format="%.0f"),
    },
)

# 6) Calculs
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
tot_cartons  = int(pd.to_numeric(df_calc["Quantité cartons"], errors="coerce").fillna(0).sum())
tot_palettes = int(pd.to_numeric(df_calc["Quantité palettes"], errors="coerce").fillna(0).sum())
tot_poids    = int(pd.to_numeric(df_calc["Poids palettes (kg)"], errors="coerce").fillna(0).sum())

c1, c2, c3 = st.columns(3)
with c1: kpi("Total cartons", f"{tot_cartons:,}".replace(",", " "))
with c2: kpi("Total palettes", f"{tot_palettes}")
with c3: kpi("Poids total (kg)", f"{tot_poids:,}".replace(",", " "))
st.dataframe(df_calc[display_cols], use_container_width=True, hide_index=True)

# 7) Téléchargement XLSX
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

# 7-bis) Téléchargement PDF
if st.button("🧾 Télécharger la version PDF", use_container_width=True):
    if tot_cartons <= 0:
        st.error("Renseigne au moins une **Quantité cartons** > 0.")
    else:
        try:
            pdf_bytes = build_bl_enlevements_pdf(
                date_creation=_today_paris(),
                date_ramasse=date_ramasse,
                destinataire_title=DEST_TITLE,
                destinataire_lines=DEST_LINES,
                df_lines=df_calc[display_cols],
            )
            st.session_state["fiche_ramasse_pdf"] = pdf_bytes
            st.download_button(
                "📄 Télécharger la version PDF",
                data=pdf_bytes,
                file_name=f"Fiche_de_ramasse_{date_ramasse:%Y%m%d}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"Erreur PDF : {e}")

# ======================== ENVOI PAR E-MAIL ====================================
# 1) Total palettes
PALETTE_COL_CANDIDATES = ["Quantité palettes", "N° palettes", "Nb palettes", "Quantite palettes"]
pal_col = next((c for c in PALETTE_COL_CANDIDATES if c in df_calc.columns), None)
if pal_col is None:
    st.error("Colonne des palettes introuvable dans df_calc. Renomme une des colonnes en " + ", ".join(PALETTE_COL_CANDIDATES))
else:
    total_palettes = int(pd.to_numeric(df_calc[pal_col], errors="coerce").fillna(0).sum())

    # 2) Récup PDF
    pdf_bytes = st.session_state.get("fiche_ramasse_pdf")
    if pdf_bytes is None:
        st.info("Génère d’abord la version PDF (bouton de téléchargement) pour pouvoir l’envoyer par e-mail.")

    # 3) UI destinataires (pré-rempli sans masquage ***)
    try:
        _cfg_preview = _get_email_cfg()
        sender_hint = _cfg_preview.get("sender", _cfg_preview.get("user"))
        rec = _cfg_preview.get("recipients", [])
        rec_str = rec if isinstance(rec, str) else ", ".join([x for x in rec if x])
    except RuntimeError as e:
        sender_hint = None
        rec_str = ""
        st.caption(f"ℹ️ {e} — place ton fichier dans **.streamlit/secrets.toml** ou configure les secrets du déploiement.")

    _PREFILL = (rec_str or "") + "\u200b"   # anti-masquage Streamlit
    if "ramasse_email_to" not in st.session_state:
        st.session_state["ramasse_email_to"] = _PREFILL

    to_input = st.text_input(
        "Destinataires (séparés par des virgules)",
        key="ramasse_email_to",
        placeholder="ex: logistique@transporteur.com, expeditions@tonentreprise.fr",
    )

    def _parse_emails(s: str):
        return [e.strip() for e in (s or "").replace("\u200b","").split(",") if e.strip()]

    to_list = _parse_emails(st.session_state.get("ramasse_email_to",""))

    if sender_hint:
        st.caption(f"Expéditeur utilisé : **{sender_hint}**")

    # (Optionnel) Bouton de test auth SMTP
    if st.button("🔌 Tester la connexion SMTP", use_container_width=True):
        try:
            cfg = _get_email_cfg()
            if int(cfg["port"]) == 465:
                import ssl
                with smtplib.SMTP_SSL(cfg["host"], 465, context=ssl.create_default_context()) as s:
                    s.login(cfg["user"], cfg["password"])
            else:
                with smtplib.SMTP(cfg["host"], int(cfg["port"])) as s:
                    s.ehlo(); s.starttls(); s.ehlo()
                    s.login(cfg["user"], cfg["password"])
            st.success("Connexion SMTP OK : authentification acceptée ✅")
        except Exception as e:
            st.error(f"Échec d'authentification SMTP : {e}")

    # Envoi
    if st.button("✉️ Envoyer la demande de ramasse", type="primary", use_container_width=True):
        if pdf_bytes is None:
            st.error("Le PDF n’est pas prêt. Clique d’abord sur « Télécharger la version PDF ».")
        elif not to_list:
            st.error("Indique au moins un destinataire.")
        else:
            try:
                filename = f"Fiche_de_ramasse_{date_ramasse.strftime('%Y%m%d')}.pdf"
                size_mb = len(pdf_bytes) / (1024*1024)
                st.caption(f"Taille PDF : {size_mb:.2f} Mo")

                refused = send_mail_with_pdf(pdf_bytes, filename, total_palettes, to_list, date_ramasse, bcc_me=True)

                st.write("Destinataires envoyés :", ", ".join(to_list))
                if refused:
                    bad = ", ".join(f"{k} ({v[0]})" for k, v in refused.items())
                    st.warning(f"E-mail refusé pour : {bad} — adresse ou politique du domaine.")
                else:
                    st.success("Serveur SMTP : OK ✅ — message remis au transport. "
                               "Si le destinataire ne le voit pas, il est probablement en quarantaine/filtre.")
            except Exception as e:
                st.error(f"Échec de l’envoi : {e}")
