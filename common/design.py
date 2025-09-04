import re, unicodedata, os
from io import BytesIO
from PIL import Image
import streamlit as st

COLORS = {
    "bg": "#F7F4EF", "ink": "#2D2A26", "green": "#2F7D5A",
    "sage": "#8BAA8B", "lemon": "#EEDC5B", "card": "#FFFFFF",
}

def apply_theme(page_title="Ferment Station", icon="🥤"):
    st.set_page_config(page_title=page_title, page_icon=icon, layout="wide")
    st.markdown(f"""
    <style>
      .block-container {{ max-width: 1400px; padding-top: 1rem; padding-bottom: 3rem; }}
      h1,h2,h3,h4,h5 {{ color:{COLORS['ink']}; letter-spacing:.2px; }}
      .section-title {{
        display:flex; align-items:center; gap:.5rem; padding:.4rem .8rem;
        background:{COLORS['sage']}22; border-left:6px solid {COLORS['sage']};
        border-radius:14px; margin:.2rem 0 1rem 0;
      }}
      .kpi {{
        background:{COLORS['card']}; border:1px solid #0001;
        border-left:6px solid {COLORS['green']}; border-radius:14px; padding:16px;
      }}
      .kpi .t {{ font-size:.9rem; color:#555; margin-bottom:6px; }}
      .kpi .v {{ font-size:1.5rem; font-weight:700; color:{COLORS['ink']}; }}
      div.stButton > button:first-child {{ background:{COLORS['green']}; color:#fff; border:none; border-radius:12px; }}
    </style>
    """, unsafe_allow_html=True)

def section(title: str, emoji=""):
    t = f"{emoji} {title}" if emoji else title
    st.markdown(f'<div class="section-title"><h2 style="margin:0">{t}</h2></div>', unsafe_allow_html=True)

def kpi(title: str, value: str):
    st.markdown(f'<div class="kpi"><div class="t">{title}</div><div class="v">{value}</div></div>', unsafe_allow_html=True)

# ---------- Images helpers ----------
IMG_EXTS = (".png", ".jpg", ".jpeg", ".webp", ".gif")

def slugify(s: str) -> str:
    s = str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"[^a-zA-Z0-9]+", "-", s).strip("-").lower()
    return s

def find_image_path(images_dir: str, sku: str = None, flavor: str = None):
    # 1) priorité SKU exact (CITR-33.png)
    if sku:
        base_sku = sku
        for ext in IMG_EXTS:
            p = os.path.join(images_dir, f"{base_sku}{ext}")
            if os.path.exists(p): return p
        # fallback racine sans format (CITR.png)
        import re as _re
        base_root = _re.sub(r"-\d+$", "", base_sku)
        for ext in IMG_EXTS:
            p = os.path.join(images_dir, f"{base_root}{ext}")
            if os.path.exists(p): return p
    # 2) slug du goût canonique
    if flavor:
        s = slugify(flavor)
        for ext in IMG_EXTS:
            p = os.path.join(images_dir, f"{s}{ext}")
            if os.path.exists(p): return p
    return None

def load_image_bytes(path: str):
    if not path or not os.path.exists(path): return None
    im = Image.open(path).convert("RGBA")
    buf = BytesIO()
    im.save(buf, format="PNG")
    return buf.getvalue()

