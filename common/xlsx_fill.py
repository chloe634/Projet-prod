# common/xlsx_fill.py
from __future__ import annotations
import io
from datetime import date
from typing import Optional, Tuple, Dict
from dateutil.relativedelta import relativedelta
import pandas as pd
import openpyxl

from openpyxl.utils import get_column_letter

def _set(ws, addr: str, value, number_format: str | None = None):
    """
    Ecrit `value` dans `addr`. Si `addr` appartient à une zone fusionnée,
    écrit dans la cellule *top-left* de cette zone (sinon openpyxl lève MergedCell read-only).
    Retourne la coordonnée effectivement utilisée.
    """
    for rng in ws.merged_cells.ranges:
        if addr in rng:  # addr est contenu dans la zone fusionnée
            tl = f"{get_column_letter(rng.min_col)}{rng.min_row}"
            cell = ws[tl]
            cell.value = value
            if number_format:
                cell.number_format = number_format
            return tl
    cell = ws[addr]
    cell.value = value
    if number_format:
        cell.number_format = number_format
    return addr

VOL_TOL = 0.02

def _is_close(a: float, b: float, tol: float = VOL_TOL) -> bool:
    try:
        return abs(float(a) - float(b)) <= tol
    except Exception:
        return False

def _agg_counts_by_format_and_brand(df_calc: pd.DataFrame, gout: str) -> Dict[str, Dict[str, int]]:
    """
    Agrège CARTONS et BOUTEILLES à produire pour un goût donné, ventilés ainsi :
      - 33cl x12 FRANCE -> key "33_fr"
      - 33cl x12 NIKO   -> key "33_niko"
      - 75cl x6         -> key "75x6"
      - 75cl x4         -> key "75x4"

    Règle 33cl : si libellé produit contient 'NIKO' => NIKO, sinon FRANCE.
                 Tout libellé contenant 'Kéfir/Kefir' est rangé FRANCE par défaut.
    """
    out = {
        "33_fr":  {"cartons": 0, "bouteilles": 0},
        "33_niko":{"cartons": 0, "bouteilles": 0},
        "75x6":   {"cartons": 0, "bouteilles": 0},
        "75x4":   {"cartons": 0, "bouteilles": 0},
    }
    if df_calc is None or not isinstance(df_calc, pd.DataFrame) or df_calc.empty:
        return out

    req = {
        "GoutCanon", "Produit", "Bouteilles/carton", "Volume bouteille (L)",
        "Cartons à produire (arrondi)", "Bouteilles à produire (arrondi)"
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

    # 33 cL x12 -> France ou NIKO
    m33 = (df["Bouteilles/carton"] == 12) & (_is_close(df["Volume bouteille (L)"], 0.33))
    if m33.any():
        part = df.loc[m33, ["Produit", "Cartons à produire (arrondi)", "Bouteilles à produire (arrondi)"]].copy()
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

def fill_fiche_7000L_xlsx(
    template_path: str,
    semaine_du: date,      # utilisé pour le nom de fichier côté appelant
    ddm: date,
    gout1: str,
    gout2: Optional[str],
    df_calc: pd.DataFrame,
) -> bytes:
    """
    Remplit la feuille 'Fiche de production 7000 L' (ou '...7000L') du modèle :
      - D8 = Produit 1 (goût), T8 = Produit 2
      - D10 = DDM (format JJ/MM/AAAA)
      - O10 = LOT = DDM sans '/'
      - A20 = Date = DDM - 1 an
      - D/F/H/J 15-16 et T/V/X/Z 15-16 = cartons & bouteilles par format
    Retourne les bytes XLSX du classeur rempli (formules préservées).
    """
    wb = openpyxl.load_workbook(template_path, data_only=False, keep_vba=False)

    # Feuille tolérante (avec/sans espace)
    target_names = ["Fiche de production 7000 L", "Fiche de production 7000L"]
    ws = None
    for nm in target_names:
        if nm in wb.sheetnames:
            ws = wb[nm]
            break
    if ws is None:
        raise KeyError(f"Feuille cible introuvable. Feuilles présentes : {wb.sheetnames}")

    # Produits
    _set(ws, "D8", gout1 or "")
    _set(ws, "T8", gout2 or "")

    # DDM & LOT
    _set(ws, "D10", ddm, number_format="DD/MM/YYYY")
    _set(ws, "O10", ddm.strftime("%d%m%Y"))

    # Fermentation > Date (DDM - 1 an)
    ferment_date = ddm - relativedelta(years=1)
    _set(ws, "A20", ferment_date, number_format="DD/MM/YYYY")


    # Cell mapping : adapte si ta maquette a d'autres lignes
    CELLS_P1 = {  # Produit 1
        "33_fr":  {"cartons": "D15", "bouteilles": "D16"},
        "33_niko":{"cartons": "F15", "bouteilles": "F16"},
        "75x6":   {"cartons": "H15", "bouteilles": "H16"},
        "75x4":   {"cartons": "J15", "bouteilles": "J16"},
    }
    CELLS_P2 = {  # Produit 2
        "33_fr":  {"cartons": "T15", "bouteilles": "T16"},
        "33_niko":{"cartons": "V15", "bouteilles": "V16"},
        "75x6":   {"cartons": "X15", "bouteilles": "X16"},
        "75x4":   {"cartons": "Z15", "bouteilles": "Z16"},
    }

    # Produit 1
    agg1 = _agg_counts_by_format_and_brand(df_calc, gout1)
    for key, dest in CELLS_P1.items():
        _set(ws, dest["cartons"],    int(agg1[key]["cartons"]))
        _set(ws, dest["bouteilles"], int(agg1[key]["bouteilles"]))

    # Produit 2 (ou zéros)
    if gout2:
        agg2 = _agg_counts_by_format_and_brand(df_calc, gout2)
        for key, dest in CELLS_P2.items():
            _set(ws, dest["cartons"],    int(agg2[key]["cartons"]))
            _set(ws, dest["bouteilles"], int(agg2[key]["bouteilles"]))
    else:
        for key, dest in CELLS_P2.items():
            _set(ws, dest["cartons"], 0)
            _set(ws, dest["bouteilles"], 0)


    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()
