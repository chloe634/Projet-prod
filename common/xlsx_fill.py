# common/xlsx_fill.py
from __future__ import annotations
import io
from datetime import date
from dateutil.relativedelta import relativedelta
from typing import Optional, Tuple
import numpy as np
import pandas as pd
import openpyxl

VOL_TOL = 0.02

def _is_close(a: float, b: float, tol: float = VOL_TOL) -> bool:
    try:
        return abs(float(a) - float(b)) <= tol
    except Exception:
        return False

def _agg_counts_by_format_and_brand(df_calc: pd.DataFrame, gout: str):
    """
    Retourne un dict avec (cartons, bouteilles) pour un goût donné :
      - 33cl x12 France -> key "33_fr"
      - 33cl x12 NIKO   -> key "33_niko"
      - 75cl x6         -> key "75x6"
      - 75cl x4         -> key "75x4"
    Règle : en 33cl, si le nom produit contient 'NIKO' -> NIKO, sinon FRANCE.
            (Et tout libellé contenant 'Kéfir' est rangé FRANCE par défaut.)
    """
    out = {
        "33_fr":  {"cartons": 0, "bouteilles": 0},
        "33_niko":{"cartons": 0, "bouteilles": 0},
        "75x6":   {"cartons": 0, "bouteilles": 0},
        "75x4":   {"cartons": 0, "bouteilles": 0},
    }
    if df_calc is None or not isinstance(df_calc, pd.DataFrame) or df_calc.empty:
        return out

    req = {"GoutCanon","Produit","Bouteilles/carton","Volume bouteille (L)",
           "Cartons à produire (arrondi)","Bouteilles à produire (arrondi)"}
    if any(c not in df_calc.columns for c in req):
        return out

    df = df_calc.copy()
    df = df[df["GoutCanon"].astype(str).str.strip() == str(gout).strip()]
    if df.empty:
        return out

    def _add(where: str, ct, bt):
        out[where]["cartons"]    += int(pd.to_numeric(ct, errors="coerce").fillna(0).sum())
        out[where]["bouteilles"] += int(pd.to_numeric(bt, errors="coerce").fillna(0).sum())

    # 33 cL x12 -> France ou NIKO selon libellé produit
    mask_33x12 = (df["Bouteilles/carton"]==12) & (_is_close(df["Volume bouteille (L)"], 0.33))
    if mask_33x12.any():
        part = df.loc[mask_33x12, ["Produit","Cartons à produire (arrondi)","Bouteilles à produire (arrondi)"]].copy()
        up = part["Produit"].astype(str).str.upper()
        is_niko   = up.str.contains("NIKO", na=False)
        is_kefir  = up.str.contains("KÉFIR|KEFIR", na=False)

        # NIKO
        _add("33_niko",
             part.loc[is_niko, "Cartons à produire (arrondi)"],
             part.loc[is_niko, "Bouteilles à produire (arrondi)"])
        # FRANCE (tout le reste, et on force Kéfir en France)
        fr_mask = (~is_niko) | is_kefir
        _add("33_fr",
             part.loc[fr_mask, "Cartons à produire (arrondi)"],
             part.loc[fr_mask, "Bouteilles à produire (arrondi)"])

    # 75 cL x6
    mask_75x6 = (df["Bouteilles/carton"]==6) & (_is_close(df["Volume bouteille (L)"], 0.75))
    if mask_75x6.any():
        _add("75x6",
             df.loc[mask_75x6, "Cartons à produire (arrondi)"],
             df.loc[mask_75x6, "Bouteilles à produire (arrondi)"])

    # 75 cL x4
    mask_75x4 = (df["Bouteilles/carton"]==4) & (_is_close(df["Volume bouteille (L)"], 0.75))
    if mask_75x4.any():
        _add("75x4",
             df.loc[mask_75x4, "Cartons à produire (arrondi)"],
             df.loc[mask_75x4, "Bouteilles à produire (arrondi)"])

    return out


    # --- Quantités à produire : mapping de cellules (ajuste si besoin)
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
        ws[dest["cartons"]].value    = int(agg1[key]["cartons"])
        ws[dest["bouteilles"]].value = int(agg1[key]["bouteilles"])

    # Produit 2 (si présent), sinon zéros
    if gout2:
        agg2 = _agg_counts_by_format_and_brand(df_calc, gout2)
        for key, dest in CELLS_P2.items():
            ws[dest["cartons"]].value    = int(agg2[key]["cartons"])
            ws[dest["bouteilles"]].value = int(agg2[key]["bouteilles"])
    else:
        for key, dest in CELLS_P2.items():
            ws[dest["cartons"]].value    = 0
            ws[dest["bouteilles"]].value = 0

