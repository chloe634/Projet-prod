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
    """
    Ordre de recherche:
      0) assets/image_map.csv si présent (canonical -> filename), insensible aux accents/espaces/casse
      1) Fichier nommé par SKU (CITR-33.webp) puis racine (CITR.webp)
      2) Fichier nommé par slug du goût (mangue-passion.webp)
    """

    import os, csv, unicodedata, re as _re

    def _norm_key(s: str) -> str:
        # supprime accents, met en minuscules, condense espaces
        s = str(s or "")
        s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
        s = _re.sub(r"\s+", " ", s).strip().lower()
        return s

    # 0) mapping optionnel assets/image_map.csv
    map_csv = os.path.join(images_dir, "image_map.csv")
    if os.path.exists(map_csv) and flavor:
        # on tente séparateur virgule puis point-virgule
        loaded = False
        for sep in (",", ";"):
            try:
                d = {}
                with open(map_csv, "r", encoding="utf-8") as f:
                    rdr = csv.DictReader(f, delimiter=sep)
                    if "canonical" in (c.lower() for c in rdr.fieldnames or []) and "filename" in (c.lower() for c in rdr.fieldnames or []):
                        for row in rdr:
                            cano = (row.get("canonical") or row.get("Canonical") or "").strip()
                            fn   = (row.get("filename")  or row.get("Filename")  or "").strip()
                            if cano and fn:
                                d[_norm_key(cano)] = fn
                        loaded = True
                        break
            except Exception:
                pass
        if loaded:
            fn = d.get(_norm_key(flavor))
            if fn:
                p = os.path.join(images_dir, fn)
                if os.path.exists(p):
                    return p  # mapping trouvé

    # 1) priorité SKU exact (ex: CITR-33.webp), puis racine SKU (CITR.webp)
    if sku:
        for ext in IMG_EXTS:
            p = os.path.join(images_dir, f"{sku}{ext}")
            if os.path.exists(p):
                return p
        base_root = _re.sub(r"-\d+$", "", sku)
        for ext in IMG_EXTS:
            p = os.path.join(images_dir, f"{base_root}{ext}")
            if os.path.exists(p):
                return p

    # 2) slug du goût canonique (mangue-passion.webp)
    if flavor:
        s = slugify(flavor)
        for ext in IMG_EXTS:
            p = os.path.join(images_dir, f"{s}{ext}")
            if os.path.exists(p):
                return p

    return None


import os, base64
from io import BytesIO
from PIL import Image

def load_image_bytes(path: str):
    """
    Retourne une valeur affichable par ImageColumn :
    - Si possible : bytes PNG (compat universelle).
    - Sinon (ex. WEBP sans plugin Pillow) : data-URL 'data:image/...;base64,...'
    - Sinon : None.
    """
    if not path or not os.path.exists(path):
        return None
    ext = os.path.splitext(path)[1].lower()
    # 1) tentative conversion PNG via Pillow
    try:
        im = Image.open(path)
        im = im.convert("RGBA")
        buf = BytesIO()
        im.save(buf, format="PNG")
        return buf.getvalue()            # ✅ ImageColumn sait afficher des bytes PNG
    except Exception:
        # 2) fallback : on renvoie une data-URL que le navigateur sait décoder
        try:
            with open(path, "rb") as f:
                raw = f.read()
            # mime basique selon l’extension
            mime = {
                ".webp": "image/webp",
                ".jpg": "image/jpeg",
                ".jpeg": "image/jpeg",
                ".png": "image/png",
                ".gif": "image/gif",
            }.get(ext, "image/octet-stream")
            b64 = base64.b64encode(raw).decode("ascii")
            return f"data:{mime};base64,{b64}"   # ✅ URL affichable sans Pillow
        except Exception:
            return None
