# common/xlsx_fill.py
from __future__ import annotations
import io, re
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

# ----------- Agrégat depuis df_calc (fallback) -----------
def _agg_counts_by_format_and_brand(df_calc: pd.DataFrame, gout: str) -> Dict[str, Dict[str, int]]:
    out = {
        "33_fr":  {"cartons": 0, "bouteilles": 0},
        "33_niko":{"cartons": 0, "bouteilles": 0},
        "75x6":   {"cartons": 0, "bouteilles": 0},
        "75x4":   {"cartons": 0, "bouteilles": 0},
    }
    if df_calc is None or not isinstance(df_calc, pd.DataFrame) or df_calc.empty:
        return out

    req = {
        "GoutCanon","Produit","Bouteilles/carton","Volume bouteille (L)",
        "Cartons à produire (arrondi)","Bouteilles à produire (arrondi)"
    }
    if any(c not in df_calc.columns for c in req):
        return out

    df = df_calc.copy()
    df = df[df["GoutCanon"].astype(str).str.strip() == str(gout).strip()]
    if df.empty:
        return out

    def _add(where: str, ct, bt):
        out[where]["cartons"]    += int(pd.to_numeric(ct, errors="coerce").fillna(0).sum())
        out[where]["bouteilles"] += int(pd.to_numeric(bt, errors="coerce").fillna(0).sum())

    # 33 cL x12 -> France/NIKO selon libellé
    m33 = (df["Bouteilles/carton"] == 12) & (_is_close(df["Volume bouteille (L)"], 0.33))
    if m33.any():
        part = df.loc[m33, ["Produit","Cartons à produire (arrondi)","Bouteilles à produire (arrondi)"]].copy()
        up = part["Produit"].astype(str).str.upper()
        is_niko  = up.str.contains("NIKO", na=False)
        is_kefir = up.str.contains("KÉFIR|KEFIR", na=False)

        _add("33_niko",
             part.loc[is_niko, "Cartons à produire (arrondi)"],
             part.loc[is_niko, "Bouteilles à produire (arrondi)"])

        fr_mask = (~is_niko) | is_kefir
        _add("33_fr",
             part.loc[fr_mask, "Cartons à produire (arrondi)"],
             part.loc[fr_mask, "Bouteilles à produire (arrondi)"])

    # 75 cL x6
    m75x6 = (df["Bouteilles/carton"] == 6) & (_is_close(df["Volume bouteille (L)"], 0.75))
    if m75x6.any():
        _add("75x6",
             df.loc[m75x6, "Cartons à produire (arrondi)"],
             df.loc[m75x6, "Bouteilles à produire (arrondi)"])

    # 75 cL x4
    m75x4 = (df["Bouteilles/carton"] == 4) & (_is_close(df["Volume bouteille (L)"], 0.75))
    if m75x4.any():
        _add("75x4",
             df.loc[m75x4, "Cartons à produire (arrondi)"],
             df.loc[m75x4, "Bouteilles à produire (arrondi)"])

    return out

# ----------- Agrégat STRICT depuis df_min (tableau affiché) -----------
def _agg_from_dfmin(df_min, gout: str) -> Dict[str, Dict[str, int]]:
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

# ----------- Détection auto des blocs Quantité -----------
def _norm(s) -> str:
    return str(s).strip().lower()

def _locate_quantity_blocks(ws) -> Dict[str, Dict[str, int]]:
    """
    Le modèle contient 2 paires de blocs (haut = résumé, bas = zone d'entrée).
    On retourne volontairement **la paire du BAS** pour la saisie.
    """
    labels = {"france", "niko", "x6", "x4"}
    row_hits: Dict[int, Dict[str, int]] = {}

    for r in ws.iter_rows(values_only=False):
        for c in r:
            v = c.value
            if isinstance(v, str):
                nv = str(v).strip().lower()
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


    def _avg_col(cols: Dict[str,int]) -> float:
        return sum(cols.values()) / len(cols)

    candidates.sort(key=lambda x: _avg_col(x[1]))
    left_row, left_cols   = candidates[0]
    right_row, right_cols = candidates[-1]

    def _fill_missing(cols: Dict[str,int]) -> Dict[str,int]:
        out = cols.copy()
        keys = ["france","niko","x6","x4"]
        if keys[0] in out:
            first_c = out[keys[0]]
            for k in keys:
                out.setdefault(k, first_c)
        return out

    left_cols  = _fill_missing(left_cols)
    right_cols = _fill_missing(right_cols)

    return {
        "left":  {"header_row": left_row,  "bouteilles_row": left_row+1,  "cartons_row": left_row+2,  **left_cols},
        "right": {"header_row": right_row, "bouteilles_row": right_row+1, "cartons_row": right_row+2, **right_cols},
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
