# pages/99_Debug.py
import pathlib, traceback
import streamlit as st

st.set_page_config(page_title="Debug pages", page_icon="🛠️", layout="wide")
st.title("🛠️ Debug des pages Streamlit")

root = pathlib.Path(__file__).resolve().parents[1]  # racine du projet
pages = sorted((root / "pages").glob("*.py"))

bad = []
for p in pages:
    code = p.read_text(encoding="utf-8", errors="replace")
    try:
        compile(code, str(p), "exec")
        st.success(f"OK: {p.name}")
    except SyntaxError as e:
        st.error(f"SYNTAX ERROR dans {p.name} — ligne {e.lineno}, colonne {e.offset}")
        st.code("".join(traceback.format_exception_only(e)), language="text")
        # Montre la ligne incriminée
        lines = code.splitlines()
        i = max(0, (e.lineno or 1) - 1)
        snippet = "\n".join(lines[max(0, i-2): i+3])
        st.code(snippet, language="python")
        bad.append(p.name)

if not bad:
    st.info("✅ Toutes les pages compilent correctement.")
else:
    st.warning("Corrige les pages en erreur ci-dessus puis rafraîchis.")
