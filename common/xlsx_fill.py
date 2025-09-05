# common/xlsx_fill.py
from __future__ import annotations
import io
from datetime import date
from typing import Optional, Dict, Tuple, List
from dateutil.relativedelta import relativedelta
import pandas as pd
import openpyxl
from openpyxl.utils import coordinate_to_tuple, get_column_letter

import re

def _parse_format_from_stock(stock: str):
    """
    Extrait (nb_bouteilles_par_carton, volume_L) depuis la chaîne Stock visible
    ex: 'Carton de 12 Bouteilles - 0.33L' -> (12, 0.33)
        'Carton de 6 Bouteilles 75cl ...' -> (6, 0.75)
        'Pack de 4 Bouteilles 75cl ...'   -> (4, 0.75)
    Retourne (None, None) si introuvable.
    """
    s = str(stock or "")
    # nb bouteilles
    m_nb = re.search(r'(Carton|Pack)\s+de\s+(\d+)\s+Bouteilles?', s, flags=re.I)
    nb = int(m_nb.group(2)) if m_nb else None
    # volume
    m_l = re.search(r'(\d+(?:[.,]\d+)?)\s*[lL]\b', s)
    vol = float(m_l.group(1).replace(",", ".")) if m_l else None
    if vol is None:
        m_cl = re.search(r'(\d+(?:[.,]\d+)?)\s*c[lL]\b', s)
        vol = float(m_cl.group(1).replace(",", "."))/100.0 if m_cl else None
    return nb, vol

VOL_TOL = 0.02

def _is_close(a: float, b: float, tol: float = VOL_TOL) -> bool:
    try:
        return abs(float(a) - float(b)) <= tol
    except Exception:
        return False

# ---------- Agrégat depuis df_calc (cartons & bouteilles) ----------
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
def _agg_from_dfmin(df_min, gout: str):
    """
    Lit *exactement* les nombres visibles dans le tableau sauvegardé (df_min):
      colonnes attendues : Produit, Stock, GoutCanon,
        Cartons à produire (arrondi), Bouteilles à produire (arrondi)
    Reconnaît 33cl x12, 75cl x6, 75cl x4 via la colonne Stock
    Ventile 33cl : NIKO si 'NIKO' dans Produit, sinon FRANCE.
    """
    out = {
        "33_fr":  {"cartons": 0, "bouteilles": 0},
        "33_niko":{"cartons": 0, "bouteilles": 0},
        "75x6":   {"cartons": 0, "bouteilles": 0},
        "75x4":   {"cartons": 0, "bouteilles": 0},
    }
    import pandas as pd
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

        if nb == 12 and abs(vol-0.33) <= 0.02:
            key = "33_niko" if "NIKO" in prod_up else "33_fr"
        elif nb == 6 and abs(vol-0.75) <= 0.02:
            key = "75x6"
        elif nb == 4 and abs(vol-0.75) <= 0.02:
            key = "75x4"
        else:
            continue

        out[key]["cartons"]    += ct
        out[key]["bouteilles"] += bt

    return out



    def _add(where: str, ct, bt):
        out[where]["cartons"]    += int(pd.to_numeric(ct, errors="coerce").fillna(0).sum())
        out[where]["bouteilles"] += int(pd.to_numeric(bt, errors="coerce").fillna(0).sum())

    # 33 cL x12 -> France ou NIKO
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

# ---------- Helper : écrire même si la cellule est fusionnée ----------
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

# ---------- Détection auto des blocs “Quantité / France / NIKO / X6 / X4” ----------
def _norm(s) -> str:
    return str(s).strip().lower()

def _locate_quantity_blocks(ws) -> Dict[str, Dict[str, int]]:
    """
    Repère automatiquement les 2 blocs (gauche/droite) :
      - headers "France", "NIKO", "X6", "X4" sur une même ligne
      - 'bouteilles_row' = header_row + 1
      - 'cartons_row'    = header_row + 2
    Retourne pour chaque côté les colonnes des 4 en-têtes.
    """
    # map: row -> {label: col}
    row_hits: Dict[int, Dict[str,int]] = {}
    labels = {"france","niko","x6","x4"}

    for r in ws.iter_rows(values_only=False):
        for c in r:
            v = c.value
            if isinstance(v, str):
                nv = _norm(v)
                if nv in labels:
                    row_hits.setdefault(c.row, {})[nv] = c.column

    # On garde les lignes qui ont au moins 3 des 4 libellés (tolérant)
    candidates = [(row, cols) for row, cols in row_hits.items() if len(cols) >= 3]
    if not candidates:
        raise KeyError("En-têtes 'France/NIKO/X6/X4' introuvables.")

    # Gauche = ligne avec plus petite moyenne de colonnes; Droite = la plus grande
    def _avg_col(cols: Dict[str,int]) -> float:
        return sum(cols.values()) / len(cols)

    candidates.sort(key=lambda x: _avg_col(x[1]))
    left_row, left_cols = candidates[0]
    right_row, right_cols = candidates[-1]

    # Compléter les colonnes manquantes si besoin (par proximité)
    def _fill_missing(cols: Dict[str,int]) -> Dict[str,int]:
        out = cols.copy()
        # ordre attendu visuel : France < NIKO < X6 < X4
        # si un manque, on tente d’estimer par interpolation (très tolérant)
        keys = ["france","niko","x6","x4"]
        present = [k for k in keys if k in out]
        if present:
            first_c = out[present[0]]
            for k in keys:
                out.setdefault(k, first_c)
            # ré-ordonner grossièrement
            # pas critique : on écrira au minimum France, NIKO et X6
        return out

    left_cols  = _fill_missing(left_cols)
    right_cols = _fill_missing(right_cols)

    return {
        "left":  {"header_row": left_row,  "bouteilles_row": left_row+1,  "cartons_row": left_row+2,  **left_cols},
        "right": {"header_row": right_row, "bouteilles_row": right_row+1, "cartons_row": right_row+2, **right_cols},
    }

def _addr(col: int, row: int) -> str:
    return f"{get_column_letter(col)}{row}"

def fill_fiche_7000L_xlsx(
    template_path: str,
    semaine_du: date,
    ddm: date,
    gout1: str,
    gout2: Optional[str],
    df_calc,                      # inchangé
    sheet_name: str | None = None,
    df_min=None,                  # ✅ NOUVEAU (optionnel)
) -> bytes:
    """
    Remplit la fiche de production :
      - Produit 1 = bloc gauche ; Produit 2 = bloc droite (si présent)
      - DDM (JJ/MM/AAAA) + LOT = DDM sans '/'
      - Date fermentation = DDM - 1 an
      - Quantités détectées automatiquement (on ne dépend pas d'adresses en dur)
      - Si df_min est fourni, on lit STRICTEMENT les nombres affichés dans le tableau sauvegardé
        (sinon fallback sur df_calc).
    """
    wb = openpyxl.load_workbook(template_path, data_only=False, keep_vba=False)

    # Choix de l’onglet (garde la logique souple que tu avais)
    targets = [sheet_name] if sheet_name else ["Fiche de production 7000 L", "Fiche de production 7000L"]
    ws = None
    for nm in targets:
        if nm and nm in wb.sheetnames:
            ws = wb[nm]
            break
    if ws is None:
        raise KeyError(f"Feuille cible introuvable. Feuilles présentes : {wb.sheetnames}")

    # En-tête
    _set(ws, "D8", gout1 or "")
    _set(ws, "T8", gout2 or "")

    _set(ws, "D10", ddm, number_format="DD/MM/YYYY")
    _set(ws, "O10", ddm.strftime("%d%m%Y"))

    ferment_date = ddm - relativedelta(years=1)
    _set(ws, "A20", ferment_date, number_format="DD/MM/YYYY")

    # Repère dynamiquement les blocs "France / NIKO / X6 / X4"
    blocks = _locate_quantity_blocks(ws)
    L = blocks["left"];  R = blocks["right"]

    # Construit les adresses bouteilles/cartons à partir des en-têtes détectés
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

    # 👉 AGRÉGATS : on privilégie df_min (tableau affiché et sauvegardé)
    agg1 = _agg_from_dfmin(df_min, gout1) if df_min is not None else _agg_counts_by_format_and_brand(df_calc, gout1)
    for k, dest in P1.items():
        _set(ws, dest["b"], int(agg1[k]["bouteilles"]))
        _set(ws, dest["c"], int(agg1[k]["cartons"]))

    if gout2:
        agg2 = _agg_from_dfmin(df_min, gout2) if df_min is not None else _agg_counts_by_format_and_brand(df_calc, gout2)
        for k, dest in P2.items():
            _set(ws, dest["b"], int(agg2[k]["bouteilles"]))
            _set(ws, dest["c"], int(agg2[k]["cartons"]))
    else:
        for k, dest in P2.items():
            _set(ws, dest["b"], 0); _set(ws, dest["c"], 0)

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()

