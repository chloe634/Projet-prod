# app.py — Accueil + préflight syntaxe des pages
import pathlib, traceback
import streamlit as st
import pandas as pd

# ---------- PRE-FLIGHT : détecte les erreurs de syntaxe dans pages/*.py ----------
def _preflight_pages():
    root = pathlib.Path(__file__).resolve().parent
    pages = sorted((root / "pages").glob("*.py"))
    bad = []
    for p in pages:
        code = p.read_text(encoding="utf-8", errors="replace")
        try:
            compile(code, str(p), "exec")
        except SyntaxError as e:
            st.set_page_config(page_title="Erreur de syntaxe", page_icon="🛑", layout="wide")
            st.title("🛑 Erreur de syntaxe dans une page Streamlit")
            st.error(f"Fichier : `{p.name}` — ligne **{e.lineno}**, colonne **{e.offset}**")
            st.code("".join(traceback.format_exception_only(e)), language="text")
            # extrait de code : 2 lignes avant/après
            lines = code.splitlines()
            i = max(0, (e.lineno or 1) - 1)
            snippet = "\n".join(lines[max(0, i-2): i+3])
            st.code(snippet, language="python")
            st.info("Corrige ce fichier dans GitHub → Commit → recharge l’app.")
            bad.append(p)
    if bad:
        st.stop()

_preflight_pages()
# ---------- FIN PRE-FLIGHT ------------------------------------------------------

# --- Accueil “Uploader unique” (ton code d’origine) ---
from common.design import apply_theme, section
from core.optimizer import read_input_excel_and_period_from_upload

apply_theme("Ferment Station — Accueil", "🥤")
section("Accueil", "🏠")
st.caption("Dépose ici ton fichier Excel. Il sera utilisé automatiquement dans tous les onglets.")

uploaded = st.file_uploader("Dépose un Excel (.xlsx / .xls)", type=["xlsx", "xls"])
col1, col2 = st.columns([1,1])
with col1:
    clear = st.button("♻️ Réinitialiser le fichier chargé", use_container_width=True)
with col2:
    show_head = st.toggle("Afficher un aperçu (20 premières lignes)", value=True)

if clear:
    for k in ("df_raw", "window_days", "file_name"):
        if k in st.session_state:
            del st.session_state[k]
    st.success("Fichier déchargé. Dépose un nouvel Excel pour continuer.")

if uploaded is not None:
    try:
        df_raw, window_days = read_input_excel_and_period_from_upload(uploaded)
        st.session_state.df_raw = df_raw
        st.session_state.window_days = window_days
        st.session_state.file_name = uploaded.name
        st.success(f"Fichier chargé ✅ : **{uploaded.name}** · Fenêtre détectée (B2) : **{window_days} jours**")
    except Exception as e:
        st.error(f"Erreur de lecture de l'Excel : {e}")

if "df_raw" in st.session_state:
    st.info(f"Fichier en mémoire : **{st.session_state.get('file_name','(sans nom)')}** — fenêtre : **{st.session_state.get('window_days', '—')} jours**")
    if show_head:
        st.dataframe(st.session_state.df_raw.head(20), use_container_width=True)
else:
    st.warning("Aucun fichier en mémoire. Dépose un Excel ci-dessus pour activer les autres onglets.")
