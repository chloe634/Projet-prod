# common/xlsx_fill.py
from __future__ import annotations

import io
import os
import re
import unicodedata
from datetime import date, datetime
from typing import Optional, Dict, List, Tuple

from dateutil.relativedelta import relativedelta
import pandas as pd
import openpyxl
from openpyxl.utils import coordinate_to_tuple, get_column_letter

# ======================================================================
#                         Utilitaires généraux
# ======================================================================

VOL_TOL = 0.02

def _is_close(a: float, b: float, tol: float = VOL_TOL) -> bool:
    try:
        return abs(float(a) - float(b)) <= tol
    except Exception:
        return False

# ----------- parse format depuis la colonne "Stock" (df_min) -----------
def _parse_format_from_stock(stock: str):
    s = str(stock or "")
    m_nb = re.search(r'(Carton|Pack)\s+de\s+(\d+)\s+Bouteilles?', s, flags=re.I)
    nb = int(m_nb.group(2)) if m_nb else None
    m_l = re.search(r'(\d+(?:[.,]\d+)?)\s*[lL]\b', s)
    vol = float(m_l.group(1).replace(",", ".")) if m_l else None
    if vol is None:
        m_cl = re.search(r'(\d+(?:[.,]\d+)?)\s*c[lL]\b', s)
        vol = float(m_cl.group(1).replace(",", "."))/100.0 if m_cl else None
    return nb, vol

# ----------- Agrégat STRICT depuis df_min (tableau affiché) -----------
def _agg_from_dfmin(df_min: pd.DataFrame, gout: str) -> Dict[str, Dict[str, int]]:
    out = {
        "33_fr":  {"cartons": 0, "bouteilles": 0},
        "33_niko":{"cartons": 0, "bouteilles": 0},
        "75x6":   {"cartons": 0, "bouteilles": 0},
        "75x4":   {"cartons": 0, "bouteilles": 0},
    }
    if df_min is None or not isinstance(df_min, pd.DataFrame) or df_min.empty:
        return out
    req = {"Produit","Stock","GoutCanon","Cartons à produire (arrondi)","Bouteilles à produire (arrondi)"}
    if any(c not in df_min.columns for c in req):
        return out

    df = df_min.copy()
    df = df[df["GoutCanon"].astype(str).str.strip() == str(gout).strip()]
    if df.empty:
        return out

    for _, r in df.iterrows():
        nb, vol = _parse_format_from_stock(r["Stock"])
        if nb is None or vol is None:
            continue
        ct = int(pd.to_numeric(r["Cartons à produire (arrondi)"], errors="coerce") or 0)
        bt = int(pd.to_numeric(r["Bouteilles à produire (arrondi)"], errors="coerce") or 0)
        prod_up = str(r["Produit"]).upper()

        if nb == 12 and _is_close(vol, 0.33):
            key = "33_niko" if "NIKO" in prod_up else "33_fr"
        elif nb == 6 and _is_close(vol, 0.75):
            key = "75x6"
        elif nb == 4 and _is_close(vol, 0.75):
            key = "75x4"
        else:
            continue

        out[key]["cartons"]    += ct
        out[key]["bouteilles"] += bt

    return out

# ======================================================================
#                   Outils sûrs d’écriture Excel (fusions)
# ======================================================================

def _safe_set_cell(ws, row: int, col: int, value, number_format: str | None = None):
    """
    Écrit une valeur *même si* (row,col) tombe dans une cellule fusionnée.
    - Si (row,col) est l'ancre: écrit directement.
    - Si c'est à l'intérieur d'une fusion (pas l'ancre): on unmerge -> écrit à l'ancre -> remerge.
    Neutralise 'MergedCell ... value is read-only'.
    """
    hit = None
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            hit = rng
            break

    if hit is None:
        cell = ws.cell(row=row, column=col)
        cell.value = value
        if number_format:
            cell.number_format = number_format
        return

    a_row, a_col = hit.min_row, hit.min_col
    coord = hit.coord  # ex: "C12:H14"
    # si on est déjà sur l'ancre, pas besoin de dé-fusionner
    if row == a_row and col == a_col:
        cell = ws.cell(row=a_row, column=a_col)
        cell.value = value
        if number_format:
            cell.number_format = number_format
        return

    # sinon on force
    ws.unmerge_cells(coord)
    cell = ws.cell(row=a_row, column=a_col)
    cell.value = value
    if number_format:
        cell.number_format = number_format
    ws.merge_cells(coord)

# Ecrit via adresse A1 ("D10" …) en gérant les fusions
def _set(ws, addr: str, value, number_format: str | None = None):
    row, col = coordinate_to_tuple(addr)
    _safe_set_cell(ws, row, col, value, number_format)
    return f"{get_column_letter(col)}{row}"

def _addr(col: int, row: int) -> str:
    return f"{get_column_letter(col)}{row}"

# ======================================================================
#     Fiche de prod 7000L — (laisse tel quel si tu l’utilises)
# ======================================================================

def fill_fiche_7000L_xlsx(
    template_path: str,
    semaine_du: date,
    ddm: date,
    gout1: str,
    gout2: Optional[str],
    df_calc,
    sheet_name: str | None = None,
    df_min=None,
) -> bytes:
    wb = openpyxl.load_workbook(template_path, data_only=False, keep_vba=False)

    targets = [sheet_name] if sheet_name else ["Fiche de production 7000 L", "Fiche de production 7000L"]
    ws = None
    for nm in targets:
        if nm and nm in wb.sheetnames:
            ws = wb[nm]
            break
    if ws is None:
        raise KeyError(f"Feuille cible introuvable. Feuilles présentes : {wb.sheetnames}")

    # En-têtes (exemple)
    _set(ws, "D8", gout1 or "")
    _set(ws, "T8", gout2 or "")
    _set(ws, "D10", ddm, number_format="DD/MM/YYYY")
    _set(ws, "O10", ddm.strftime("%d%m%Y"))
    ferment_date = ddm - relativedelta(years=1)
    _set(ws, "A20", ferment_date, number_format="DD/MM/YYYY")

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()

# ======================================================================
#                   Remplissage BL enlèvements Sofripa
# ======================================================================

def _iter_cells(ws):
    for r in ws.iter_rows(values_only=False):
        for c in r:
            yield c

def _find_cell_by_regex(ws, pattern: str) -> Tuple[int, int] | Tuple[None, None]:
    rx = re.compile(pattern, flags=re.I)
    for cell in _iter_cells(ws):
        v = cell.value
        if isinstance(v, str) and rx.search(v):
            return cell.row, cell.column
    return None, None

def _write_right_of(ws, row: int, col: int, value):
    """Écrit dans la cellule immédiatement à droite (gère fusions)."""
    _safe_set_cell(ws, row, col + 1, value)

def _write_cell(ws, row: int, col: int, value):
    """Écrit (row,col) en gérant les fusions."""
    _safe_set_cell(ws, row, col, value)

def _normalize_header_text(s: str) -> str:
    s = str(s or "").strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.replace("’", "'")
    for ch in ["(", ")", ":", ";", ","]:
        s = s.replace(ch, " ")
    s = " ".join(s.split())
    return s

def _first_data_row_after_header(ws, hdr_row: int, cols: List[int]) -> int:
    """
    Si l'en-tête est fusionné sur 2+ lignes, commence à la 1ère ligne
    *après* ces fusions sur n'importe laquelle des colonnes du tableau.
    """
    start = hdr_row + 1
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= hdr_row <= rng.max_row:
            if any(rng.min_col <= c <= rng.max_col for c in cols if c is not None):
                start = max(start, rng.max_row + 1)
    return start

def fill_bl_enlevements_xlsx(
    template_path: str,
    date_creation: date,
    date_ramasse: date,
    destinataire_title: str,
    destinataire_lines: List[str],
    df_lines: pd.DataFrame,   # colonnes attendues (ordre libre) cf. ci-dessous
) -> bytes:
    """
    Remplit le modèle XLSX 'LOG_EN_001_01 BL enlèvements Sofripa-2.xlsx'.

    df_lines doit contenir les colonnes (noms exacts ou équivalents) :
      - 'Référence'
      - 'Produit (goût + format)' ou 'Produit'
      - 'DDM'
      - 'Quantité cartons'
      - 'Quantité palettes'
      - 'Poids palettes (kg)'
    """
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Modèle Excel introuvable: {template_path}")

    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    # ----- 1) Dates -----
    r, c = _find_cell_by_regex(ws, r"date\s+de\s+cr[eé]ation")
    if r and c:
        _write_right_of(ws, r, c, date_creation.strftime("%d/%m/%Y"))

    r, c = _find_cell_by_regex(ws, r"date\s+de\s+rammasse|date\s+de\s+ramasse")
    if r and c:
        _write_right_of(ws, r, c, date_ramasse.strftime("%d/%m/%Y"))

    # ----- 2) Destinataire (adresse DANS l'encadré, multi-lignes) -----
    from openpyxl.styles import Alignment

    r, c = _find_cell_by_regex(ws, r"destinataire")
    if r and c:
        # cherche une fusion à droite du libellé
        target_rng = None
        for rng in ws.merged_cells.ranges:
            if rng.min_row <= r <= rng.max_row and rng.min_col > c:
                if target_rng is None or rng.min_col < target_rng.min_col:
                    target_rng = rng

        if target_rng:
            rr, cc = target_rng.min_row, target_rng.min_col
            rr_end, cc_end = target_rng.max_row, target_rng.max_col
        else:
            # crée une fusion (3 lignes x 6 colonnes) à droite du libellé
            rr, cc = r, c + 1
            rr_end = min(r + 2, ws.max_row)
            cc_end = min(c + 6, ws.max_column)
            try:
                ws.merge_cells(start_row=rr, start_column=cc, end_row=rr_end, end_column=cc_end)
            except Exception:
                pass  # si déjà partiellement fusionné, on écrira en haut-gauche

        text_lines = [destinataire_title] + (destinataire_lines or [])
        text = "\n".join([str(x).strip() for x in text_lines if str(x).strip()])
        _safe_set_cell(ws, rr, cc, text)
        anchor = ws.cell(row=rr, column=cc)
        anchor.alignment = Alignment(wrap_text=True, vertical="top")

        # ajuste la hauteur des lignes de la zone
        n_lines = max(1, text.count("\n") + 1)
        span_rows = max(1, (rr_end - rr + 1))
        per_row_height = 14 * n_lines / span_rows
        for rset in range(rr, rr_end + 1):
            cur = ws.row_dimensions[rset].height or 0
            ws.row_dimensions[rset].height = max(cur, per_row_height)

        # nettoyage éventuel d'une ligne parasite ailleurs
        zr, zc = _find_cell_by_regex(ws, r"zac\s+du\s+haut\s+de\s+wissous")
        if zr and zc:
            _safe_set_cell(ws, zr, zc, "")

    # ----- 3) En-têtes du tableau (détection robuste) -----
    def _norm(x): return _normalize_header_text(x)

    SYN = {
        "ref":   ["référence", "reference"],
        "prod":  ["produit", "produit (gout + format)", "produit gout format"],
        "ddm":   ["ddm", "date de durabilite", "date de durabilité"],
        "q_cart":["quantité cartons", "quantite cartons", "n° cartons", "nb cartons", "no cartons"],
        "q_pal": ["quantité palettes", "quantite palettes", "n° palettes", "nb palettes", "no palettes"],
        "poids": ["poids palettes (kg)", "poids palettes", "poids (kg)"],
    }

    def _row_tokens(rw):
        maxc = min(ws.max_column, 120)
        return [_norm(ws.cell(row=rw, column=j).value) for j in range(1, maxc+1)]

    def _find_header_row():
        best = (0, None, None)  # hits, row, tokens
        for rw in range(1, min(ws.max_row, 200)+1):
            toks = _row_tokens(rw)
            has_ref = any(t in SYN["ref"] for t in toks)
            has_ddm = any(t in SYN["ddm"] for t in toks)
            if has_ref and has_ddm:
                bonus = sum(any(any(s in t for s in SYN[k]) for t in toks) for k in ("q_cart","q_pal","poids"))
                return rw, toks, bonus
            hit = sum(any(any(s in t for s in SYN[k]) for t in toks) for k in ("ref","prod","ddm","q_cart","q_pal","poids"))
            if hit > best[0]:
                best = (hit, rw, toks)
        return best[1], best[2], 0

    hdr_row, hdr_toks, _ = _find_header_row()
    if not hdr_row:
        raise KeyError("Ligne d’en-têtes du tableau introuvable dans le modèle Excel.")

    def _find_col(targ_keys):
        """colonne dont le texte correspond à l’un des synonymes"""
        maxc = min(ws.max_column, 120)
        wanted = [w for k in targ_keys for w in SYN[k]]
        for j in range(1, maxc+1):
            hv = _norm(ws.cell(row=hdr_row, column=j).value)
            if hv in wanted or any(w in hv for w in wanted if len(w) >= 3):
                return j
        return None

    c_ref   = _find_col(["ref"])
    c_prod  = _find_col(["prod"])
    c_ddm   = _find_col(["ddm"])
    c_qc    = _find_col(["q_cart"])
    c_qp    = _find_col(["q_pal"])
    c_poids = _find_col(["poids"])

    # Fallbacks positionnels autour de "Produit"
    def _clamp_col(j: int | None) -> int | None:
        if j is None: return None
        return max(1, min(int(j), ws.max_column))

    if c_prod is not None:
        if c_ref is None:   c_ref = c_prod - 1
        if c_ddm is None:   c_ddm = c_prod + 1
        if c_qc  is None:   c_qc  = (c_ddm or (c_prod + 1)) + 1
        if c_qp  is None:   c_qp  = (c_qc or ((c_ddm or (c_prod + 1)) + 1)) + 1
        if c_poids is None: c_poids = (c_qp or (((c_ddm or (c_prod + 1)) + 1) + 1)) + 1

    # clamp dans les bornes
    c_ref   = _clamp_col(c_ref)
    c_prod  = _clamp_col(c_prod)
    c_ddm   = _clamp_col(c_ddm)
    c_qc    = _clamp_col(c_qc)
    c_qp    = _clamp_col(c_qp)
    c_poids = _clamp_col(c_poids)

    need = {
        "Référence": c_ref, "Produit": c_prod, "DDM": c_ddm,
        "Quantité cartons": c_qc, "Quantité palettes": c_qp, "Poids palettes (kg)": c_poids
    }
    if any(v is None for v in need.values()):
        raise ValueError(f"Colonnes incomplètes dans le modèle Excel: {need}")


    # ----- 4) Normalisation DF d'entrée -----
    df = df_lines.copy()
    if "Produit" not in df.columns and "Produit (goût + format)" in df.columns:
        df = df.rename(columns={"Produit (goût + format)": "Produit"})

    def _to_ddm_val(x):
        if isinstance(x, (date, )):
            return x.strftime("%d/%m/%Y")
        s = str(x or "").strip()
        if not s:
            return ""
        try:
            if "-" in s and len(s.split("-")[0]) == 4:
                return datetime.strptime(s, "%Y-%m-%d").strftime("%d/%m/%Y")
            return datetime.strptime(s, "%d/%m/%Y").strftime("%d/%m/%Y")
        except Exception:
            return s

    # ----- 5) Première ligne de données = après les fusions d'en-tête -----
    data_start = _first_data_row_after_header(ws, hdr_row, [c_ref, c_prod, c_ddm, c_qc, c_qp, c_poids])

    # ----- 6) Écriture des lignes (via _safe_set_cell) -----
    row = data_start

    def _as_int(v) -> int:
        try:
            f = float(v)
            return int(round(f))
        except Exception:
            return 0

    for _, r in df.iterrows():
        _write_cell(ws, row, c_ref,   str(r.get("Référence", "")))
        _write_cell(ws, row, c_prod,  str(r.get("Produit", "")))
        _write_cell(ws, row, c_ddm,   _to_ddm_val(r.get("DDM", "")))
        _write_cell(ws, row, c_qc,    _as_int(r.get("Quantité cartons", 0)))
        _write_cell(ws, row, c_qp,    _as_int(r.get("Quantité palettes", 0)))
        _write_cell(ws, row, c_poids, _as_int(r.get("Poids palettes (kg)", 0)))
        row += 1

    # ----- 7) Sauvegarde -----
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()

# =======================  PDF BL enlèvements (fpdf2)  =======================

def build_bl_enlevements_pdf(
    date_creation: date,
    date_ramasse: date,
    destinataire_title: str,
    destinataire_lines: List[str],
    df_lines: pd.DataFrame,
    *,
    logo_path: str | None = "assets/logo_symbiose.png",
    issuer_name: str = "FERMENT STATION",
    issuer_lines: List[str] | None = None,
    issuer_footer: str | None = "Produits issus de l'Agriculture Biologique certifié par FR-BIO-01",
) -> bytes:
    """PDF BL au look Excel : encadré, tableau gris, totaux. (Helvetica/latin-1)."""
    import os
    from fpdf import FPDF

    # ---------- helpers texte latin-1 ----------
    def _latin1_safe(s: str) -> str:
        s = str(s or "")
        repl = {"—": "-", "–": "-", "‒": "-", "’": "'", "‘": "'", "“": '"', "”": '"', "…": "...",
                "\u00A0": " ", "\u202F": " ", "\u2009": " ", "œ": "oe", "Œ": "OE", "€": "EUR"}
        for k, v in repl.items():
            s = s.replace(k, v)
        return s.encode("latin-1", "replace").decode("latin-1")

    def _txt(x) -> str:
        return _latin1_safe(x)

    # ---------- data ----------
    df = df_lines.copy()
    if "Produit" not in df.columns and "Produit (goût + format)" in df.columns:
        df = df.rename(columns={"Produit (goût + format)": "Produit"})

    def _ival(x):
        try:
            return int(round(float(x)))
        except Exception:
            return 0

    # ---------- PDF ----------
    pdf = FPDF("P", "mm", "A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    left, right = 15, 195
    width = right - left

    # ---- Logo + coordonnées expéditeur
    y = 18
    if logo_path and os.path.exists(logo_path):
        pdf.image(logo_path, x=left, y=y - 2, w=28)
        x_text = left + 34
    else:
        x_text = left

    pdf.set_xy(x_text, y)
    pdf.set_font("Helvetica", "B", 12); pdf.cell(0, 6, _txt(issuer_name), ln=1)
    pdf.set_x(x_text); pdf.set_font("Helvetica", "", 11)
    if issuer_lines is None:
        issuer_lines = [
            "Carré Ivry Bâtiment D2",
            "47 rue Ernest Renan",
            "94200 Ivry-sur-Seine - FRANCE",
            "Tél : 0967504647",
            "Site : https://www.symbiose-kefir.fr",
        ]
    for line in issuer_lines:
        pdf.set_x(x_text); pdf.cell(0, 5, _txt(line), ln=1)
    if issuer_footer:
        pdf.ln(2); pdf.set_x(x_text); pdf.set_font("Helvetica", "", 9)
        pdf.cell(0, 4, _txt(issuer_footer), ln=1)
    pdf.ln(2)

    # ---- Encadré "BON DE LIVRAISON"
    x_box, w_box = left, width * 0.70
    w_lbl, w_val = w_box * 0.55, w_box * 0.45
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_xy(x_box, pdf.get_y() + 2)
    pdf.cell(w_box, 8, _txt("BON DE LIVRAISON"), border=1, ln=1)

    pdf.set_font("Helvetica", "", 11)

    def _row_simple(label: str, value: str):
        pdf.set_x(x_box)
        pdf.cell(w_lbl, 8, _txt(label), border=1)
        pdf.cell(w_val, 8, _txt(value), border=1, ln=1, align="R")

    def _row_dest(label: str, title: str, lines: List[str]):
        val_text = "\n".join([title] + (lines or []))
        n_lines = len(pdf.multi_cell(w_val, 6, _txt(val_text), split_only=True)) or 1
        row_h = max(8, 6 * n_lines)
        y0 = pdf.get_y()
        pdf.set_xy(x_box, y0); pdf.cell(w_lbl, row_h, _txt(label), border=1)
        pdf.set_xy(x_box + w_lbl, y0); pdf.multi_cell(w_val, 6, _txt(val_text), border=1)
        pdf.set_xy(x_box, y0 + row_h)

    _row_simple("DATE DE CREATION :", date_creation.strftime("%d/%m/%Y"))
    _row_simple("DATE DE RAMASSE :", date_ramasse.strftime("%d/%m/%Y"))
    _row_dest("DESTINATAIRE :", destinataire_title, destinataire_lines)

    # ---- Tableau
    pdf.ln(6)
    pdf.set_fill_color(230, 230, 230)

    headers = [
        "Référence",
        "Produit",
        "DDM",
        "Nb cartons",
        "Nb palettes",
        "Poids (kg)",
    ]

    # Largeurs (somme 180) — Référence & DDM élargies
    widths_base = [30, 66, 26, 24, 22, 12]
    widths = widths_base[:]
    header_h = 8
    line_h = 6

    # Auto-ajustement des titres sur 1 ligne
    pdf.set_font("Helvetica", "B", 10)
    margin_mm = 2.5
    min_w = {0: 30.0, 1: 58.0, 2: 26.0, 3: 22.0, 4: 20.0, 5: 18.0}
    extra_needed = 0.0
    for j, h in enumerate(headers):
        if j == 1:
            continue
        need = pdf.get_string_width(_txt(h)) + 2 * margin_mm
        new_w = max(widths[j], need, min_w.get(j, widths[j]))
        extra_needed += max(0.0, new_w - widths_base[j])
        widths[j] = new_w
    widths[1] = max(min_w[1], widths[1] - extra_needed)
    total = sum(widths)
    if total > 180.0:
        overflow = total - 180.0
        take = min(overflow, max(0.0, widths[1] - min_w[1]))
        widths[1] -= take; overflow -= take
        for j in (3, 4, 5, 0, 2):
            if overflow <= 0: break
            free = max(0.0, widths[j] - min_w[j])
            d = min(free, overflow)
            widths[j] -= d; overflow -= d

    # En-tête (1 ligne)
    x = left; y = pdf.get_y()
    for h, w in zip(headers, widths):
        pdf.set_xy(x, y); pdf.cell(w, header_h, _txt(h), border=1, align="C", fill=True); x += w
    pdf.set_xy(left, y + header_h)

    # Lignes
    pdf.set_font("Helvetica", "", 10)
    tot_cart = tot_pal = tot_poids = 0

    def _maybe_break(h):
        if pdf.will_page_break(h + header_h):
            pdf.add_page()
            pdf.set_fill_color(230, 230, 230)
            pdf.set_font("Helvetica", "B", 10)
            xh = left; yh = pdf.get_y()
            for hh, ww in zip(headers, widths):
                pdf.set_xy(xh, yh); pdf.cell(ww, header_h, _txt(hh), border=1, align="C", fill=True); xh += ww
            pdf.set_xy(left, yh + header_h)
            pdf.set_font("Helvetica", "", 10)

    for _, r in df.iterrows():
        ref = _txt(r.get("Référence", ""))
        prod = _txt(r.get("Produit", ""))
        ddm = _txt(r.get("DDM", ""))
        qc = _ival(r.get("Nb cartons", r.get("Quantité cartons", 0)));   tot_cart += qc
        qp = _ival(r.get("Nb palettes", r.get("Quantité palettes", 0)));  tot_pal  += qp
        po = _ival(r.get("Poids (kg)",  r.get("Poids palettes (kg)", 0))); tot_poids += po

        prod_lines = pdf.multi_cell(widths[1], line_h, prod, split_only=True)
        row_h = max(line_h, line_h * len(prod_lines))
        _maybe_break(row_h)

        xrow = left; yrow = pdf.get_y()
        pdf.set_xy(xrow, yrow); pdf.multi_cell(widths[0], row_h, ref, border=1, align="C"); xrow += widths[0]
        pdf.set_xy(xrow, yrow); pdf.multi_cell(widths[1], line_h, prod, border=1, align="L", max_line_height=line_h); xrow += widths[1]
        pdf.set_xy(xrow, yrow); pdf.multi_cell(widths[2], row_h, ddm, border=1, align="C"); xrow += widths[2]
        pdf.set_xy(xrow, yrow); pdf.multi_cell(widths[3], row_h, str(qc), border=1, align="C"); xrow += widths[3]
        pdf.set_xy(xrow, yrow); pdf.multi_cell(widths[4], row_h, str(qp), border=1, align="C"); xrow += widths[4]
        pdf.set_xy(xrow, yrow); pdf.multi_cell(widths[5], row_h, str(po), border=1, align="C")
        pdf.set_xy(left, yrow + row_h)

    # Totaux
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(widths[0] + widths[1] + widths[2], 8, _txt("Totaux"), border=1, align="R")
    pdf.cell(widths[3], 8, _txt(f"{tot_cart:,}".replace(",", " ")), border=1, align="C")
    pdf.cell(widths[4], 8, _txt(f"{tot_pal:,}".replace(",", " ")),  border=1, align="C")
    pdf.cell(widths[5], 8, _txt(f"{tot_poids:,}".replace(",", " ")), border=1, align="C")

    return bytes(pdf.output(dest="S"))
