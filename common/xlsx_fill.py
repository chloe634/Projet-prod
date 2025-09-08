# common/xlsx_fill.py
from __future__ import annotations
import io, re, os 
from datetime import date
from typing import Optional, Dict
from dateutil.relativedelta import relativedelta
import pandas as pd
import openpyxl
from openpyxl.utils import coordinate_to_tuple, get_column_letter

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

# ----------- Helper écriture tolérante aux fusions -----------
def _set(ws, addr: str, value, number_format: str | None = None):
    row, col = coordinate_to_tuple(addr)
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            row, col = rng.min_row, rng.min_col
            break
    cell = ws.cell(row=row, column=col)
    cell.value = value
    if number_format:
        cell.number_format = number_format
    return f"{get_column_letter(col)}{row}"

# ----------- Détection auto des blocs Quantité (paire du BAS) -----------
def _norm(s) -> str:
    return str(s).strip().lower()

def _locate_quantity_blocks(ws) -> Dict[str, Dict[str, int]]:
    """
    Le modèle contient 2 paires de blocs (haut = résumé, bas = zone d'entrée).
    On retourne **la paire du BAS** pour la saisie.
    """
    labels = {"france", "niko", "x6", "x4"}
    row_hits: Dict[int, Dict[str, int]] = {}

    for r in ws.iter_rows(values_only=False):
        for c in r:
            v = c.value
            if isinstance(v, str):
                nv = _norm(v)
                if nv in labels:
                    row_hits.setdefault(c.row, {})[nv] = c.column

    candidates = [(row, cols) for row, cols in row_hits.items() if len(cols) >= 3]
    if len(candidates) < 2:
        raise KeyError("En-têtes 'France/NIKO/X6/X4' introuvables (paire du bas non détectée).")

    # On prend les 2 lignes les plus basses (bas de page)
    candidates.sort(key=lambda x: x[0])
    bottom_pair = candidates[-2:]

    def _avg_col(cols: Dict[str, int]) -> float:
        return sum(cols.values()) / len(cols)

    # gauche / droite
    bottom_pair.sort(key=lambda x: _avg_col(x[1]))
    (left_row, left_cols), (right_row, right_cols) = bottom_pair

    def _fill_missing(cols: Dict[str, int]) -> Dict[str, int]:
        out = cols.copy()
        for k in ["france", "niko", "x6", "x4"]:
            out.setdefault(k, next(iter(out.values())))
        return out

    left_cols  = _fill_missing(left_cols)
    right_cols = _fill_missing(right_cols)

    return {
        "left":  {"header_row": left_row,  "bouteilles_row": left_row + 1, "cartons_row": left_row + 2, **left_cols},
        "right": {"header_row": right_row, "bouteilles_row": right_row + 1, "cartons_row": right_row + 2, **right_cols},
    }

def _addr(col: int, row: int) -> str:
    return f"{get_column_letter(col)}{row}"

# ----------- Filler principal -----------
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

    # En-têtes
    _set(ws, "D8", gout1 or "")
    _set(ws, "T8", gout2 or "")
    _set(ws, "D10", ddm, number_format="DD/MM/YYYY")
    _set(ws, "O10", ddm.strftime("%d%m%Y"))
    ferment_date = ddm - relativedelta(years=1)
    _set(ws, "A20", ferment_date, number_format="DD/MM/YYYY")

    # Localisation des blocs
    blocks = _locate_quantity_blocks(ws)
    L = blocks["left"];  R = blocks["right"]

    P1 = {
        "33_fr":  {"b": _addr(L["france"], L["bouteilles_row"]), "c": _addr(L["france"], L["cartons_row"])},
        "33_niko":{"b": _addr(L["niko"],   L["bouteilles_row"]), "c": _addr(L["niko"],   L["cartons_row"])},
        "75x6":   {"b": _addr(L["x6"],     L["bouteilles_row"]), "c": _addr(L["x6"],     L["cartons_row"])},
        "75x4":   {"b": _addr(L["x4"],     L["bouteilles_row"]), "c": _addr(L["x4"],     L["cartons_row"])},
    }
    P2 = {
        "33_fr":  {"b": _addr(R["france"], R["bouteilles_row"]), "c": _addr(R["france"], R["cartons_row"])},
        "33_niko":{"b": _addr(R["niko"],   R["bouteilles_row"]), "c": _addr(R["niko"],   R["cartons_row"])},
        "75x6":   {"b": _addr(R["x6"],     R["bouteilles_row"]), "c": _addr(R["x6"],     R["cartons_row"])},
        "75x4":   {"b": _addr(R["x4"],     R["bouteilles_row"]), "c": _addr(R["x4"],     R["cartons_row"])},
    }

    # --- Agrégats : df_min uniquement (copie EXACTE du tableau affiché)
    agg1 = _agg_from_dfmin(df_min, gout1)
    agg2 = _agg_from_dfmin(df_min, gout2) if gout2 else None

    # N'écrit rien si 0 → on laisse les pointillés du modèle
    def _write_if_pos(addr: str, val):
        v = int(pd.to_numeric(val, errors="coerce") or 0)
        if v > 0:
            _set(ws, addr, v)

    # Gauche (Produit 1)
    for k, dest in P1.items():
        _write_if_pos(dest["b"], agg1[k]["bouteilles"])
        _write_if_pos(dest["c"], agg1[k]["cartons"])

    # Droite (Produit 2) si présent (sinon on ne touche pas aux pointillés)
    if agg2 is not None:
        for k, dest in P2.items():
            _write_if_pos(dest["b"], agg2[k]["bouteilles"])
            _write_if_pos(dest["c"], agg2[k]["cartons"])

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()

# --- AJOUTER EN BAS DE common/xlsx_fill.py ---

from typing import Dict, List, Tuple, Optional
import re
import openpyxl
from openpyxl.utils.cell import get_column_letter

def _find_cell_by_regex(ws, pattern: str):
    """Retourne (row, col) de la 1ère cellule qui match le regex (insensible à la casse)."""
    rx = re.compile(pattern, flags=re.IGNORECASE)
    for r in ws.iter_rows(values_only=True):
        pass  # force lazy eval
    for row in ws.iter_rows(values_only=False):
        for c in row:
            v = str(c.value or "").strip()
            if v and rx.search(v):
                return c.row, c.column
    return None, None

def _find_table_headers(ws, headers: List[str]) -> Tuple[Optional[int], Dict[str, int]]:
    """
    Retourne (header_row, mapping header_name -> col index) en détectant la ligne
    contenant les en-têtes (ordre quelconque).
    """
    want = {h.lower(): None for h in headers}
    header_row = None
    for row in ws.iter_rows(values_only=False):
        found = {}
        for c in row:
            v = str(c.value or "").strip().lower()
            if v in want and v not in found:
                found[v] = c.column
        # ligne valide si on a au moins 4-5 en-têtes clés
        if len(found) >= max(4, len(headers) - 2):
            header_row = row[0].row
            colmap = {}
            for h in headers:
                colmap[h] = found.get(h.lower(), None)
            return header_row, colmap
    return None, {}

def _write_right_of(ws, row: int, col_start: int, value, max_scan: int = 12):
    """
    Ecrit la valeur dans la 1ère cellule vide à droite de col_start (sur la même ligne).
    Si rien de vide n'est trouvé dans la fenêtre, écrit à col_start+1.
    """
    c = col_start + 1
    for j in range(col_start + 1, col_start + 1 + max_scan):
        cell = ws.cell(row=row, column=j)
        if cell.value in (None, ""):
            c = j
            break
    ws.cell(row=row, column=c).value = value
    return c

def fill_bl_enlevements_xlsx(
    template_path: str,
    date_creation: "date",
    date_ramasse: "date",
    destinataire_title: str,
    destinataire_lines: List[str],
    df_lines,  # DataFrame avec colonnes: Référence, Produit (goût + format), DDM, Quantité cartons, Quantité palettes, Poids palettes (kg)
) -> bytes:
    """
    Remplit le modèle Excel Sofripa automatiquement (détection des en-têtes + libellés).
    Renvoie les bytes XLSX prêts à être téléchargés.
    """
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Modèle introuvable: {template_path}")

    wb = openpyxl.load_workbook(template_path)
    ws = wb.active  # 1er onglet

    # 1) Dates : détecte les libellés et écrit la valeur à droite
    r, c = _find_cell_by_regex(ws, r"date\s+de\s+cr[eé]ation")
    if r and c:
        _write_right_of(ws, r, c, date_creation.strftime("%d/%m/%Y"))
    r, c = _find_cell_by_regex(ws, r"date\s+de\s+rammasse|date\s+de\s+ramasse")
    if r and c:
        _write_right_of(ws, r, c, date_ramasse.strftime("%d/%m/%Y"))

    # 2) Destinataire
    r, c = _find_cell_by_regex(ws, r"destinataire")
    if r and c:
        _write_right_of(ws, r, c, destinataire_title)
        # lignes suivantes (adresse)
        for i, line in enumerate(destinataire_lines[:3], start=1):
            ws.cell(row=r + i, column=c + 1).value = line

   # 3) Tableau principal : repère la ligne d’en-têtes (tolérant)
import unicodedata

def _norm_header(s: str) -> str:
    s = str(s or "").strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    for ch in ["(", ")", ":", ";", ","]:
        s = s.replace(ch, " ")
    s = " ".join(s.split())
    return s

def _match_header(name: str, header_cells: list[str], contains: bool=False) -> int | None:
    # retourne l’index 0-based dans la ligne d’en-tête
    target = _norm_header(name)
    for i, cell in enumerate(header_cells):
        h = _norm_header(cell)
        if (contains and target in h) or (not contains and h == target):
            return i
    return None

# on localise la ligne d’en-tête (on réutilise ton helper pour la ligne)
hdr_row, _ = _find_table_headers(
    ws,
    ["Référence", "Produit", "DDM", "Quantité cartons", "Quantité palettes", "Poids palettes (kg)"]
)
if not hdr_row:
    raise KeyError("Ligne d’en-têtes introuvable dans le modèle Excel.")

# valeurs brutes de la ligne d’en-tête
header_vals = [ws.cell(row=hdr_row, column=j).value for j in range(1, ws.max_column + 1)]

def _idx(label: str, contains: bool=False) -> int | None:
    i = _match_header(label, header_vals, contains=contains)
    return (i + 1) if i is not None else None   # 1-based (colonnes Excel)

# Synonymes / variantes acceptés
c_ref   = _idx("référence") or _idx("reference")
c_prod  = (
    _idx("produit")
    or _idx("produit (gout + format)", contains=True)
    or _idx("produit gout format", contains=True)
)
c_ddm   = _idx("ddm") or _idx("date de durabilite", contains=True)
c_qc    = _idx("quantité cartons") or _idx("quantite cartons")
c_qp    = _idx("quantité palettes") or _idx("quantite palettes")
c_poids = _idx("poids palettes (kg)") or _idx("poids palettes") or _idx("poids (kg)")

# Fallback : si Produit manquant mais structure Réf | Produit | DDM respectée
if c_prod is None and c_ref is not None and c_ddm is not None and c_ddm > c_ref:
    if (c_ddm - c_ref) >= 2:   # il y a une colonne entre Réf et DDM
        c_prod = c_ref + 1

# Vérification finale
need = {
    "Référence": c_ref,
    "Produit": c_prod,
    "DDM": c_ddm,
    "Quantité cartons": c_qc,
    "Quantité palettes": c_qp,
    "Poids palettes (kg)": c_poids,
}
missing = [k for k, v in need.items() if v is None]
if missing:
    raise KeyError(f"Colonnes incomplètes dans le modèle Excel: {need}")

    # 4) Ecriture des lignes
    row_out = hdr_row + 1
    for _, r in df_lines.iterrows():
        ws.cell(row=row_out, column=c_ref).value   = str(r["Référence"])
        ws.cell(row=row_out, column=c_prod).value  = str(r["Produit (goût + format)"]).replace(" — ", " - ")
        # DDM
        ddm_str = str(r["DDM"])
        # accepte "dd/mm/yyyy" en string ; si besoin, on peut parser datetime
        ws.cell(row=row_out, column=c_ddm).value   = ddm_str
        # nombres
        def _int(v): 
            try: return int(float(v))
            except: return 0
        ws.cell(row=row_out, column=c_qc).value    = _int(r["Quantité cartons"])
        ws.cell(row=row_out, column=c_qp).value    = _int(r["Quantité palettes"])
        ws.cell(row=row_out, column=c_poids).value = _int(r["Poids palettes (kg)"])
        row_out += 1

    # 5) Ligne TOTAL : cherche la cellule "TOTAL" sur la ligne suivante (sinon, on l'écrit nous-mêmes)
    total_row = None
    for j in range(0, 6):
        cell = ws.cell(row=hdr_row + 1 + len(df_lines) + j, column=c_ref)
        if str(cell.value or "").strip().upper() == "TOTAL":
            total_row = cell.row
            break

    if total_row is None:
        total_row = hdr_row + 1 + len(df_lines)
        ws.cell(row=total_row, column=c_ref).value = "TOTAL"

    # Totaux (côté chiffres)
    tot_qc = int(pd.to_numeric(df_lines["Quantité cartons"], errors="coerce").fillna(0).sum())
    tot_qp = int(pd.to_numeric(df_lines["Quantité palettes"], errors="coerce").fillna(0).sum())
    tot_pd = int(pd.to_numeric(df_lines["Poids palettes (kg)"], errors="coerce").fillna(0).sum())

    ws.cell(row=total_row, column=c_qc).value    = tot_qc
    ws.cell(row=total_row, column=c_qp).value    = tot_qp
    ws.cell(row=total_row, column=c_poids).value = tot_pd

    # 6) Sauvegarde en mémoire
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()
