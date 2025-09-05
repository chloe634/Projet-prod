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


def fill_fiche_7000L_xlsx(
    template_path: str,
    semaine_du: date,
    ddm: date,
    gout1: str,
    gout2: Optional[str],
    df_calc: pd.DataFrame,
) -> bytes:
    """
    Remplit la feuille 'Fiche de production 7000 L' du modèle.
    - D8  = Produit 1 (Goût)
    - T8  = Produit 2 (Goût) si présent
    - D10 = DDM (écrase la formule pour respecter la saisie manuelle)
    - O10 = LOT = DDM sans '/' (écrase la formule)
    - A20 = Date Fermentation = DDM - 1 an
    - D15/F15/H15/J15  et  T15/V15/X15/Z15 = cartons par format
      (par défaut on met tout le 33cl x12 en D15/T15, F15/V15 à 0)
    Retourne les bytes du classeur XLSX rempli.
    """
    wb = openpyxl.load_workbook(template_path, data_only=False, keep_vba=False)
    ws = wb["Fiche de production 7000 L"]  # ⚠️ nom exact du modèle

    # Produits
    ws["D8"].value = gout1 or ""
    ws["T8"].value = gout2 or ""

    # DDM & LOT
    ws["D10"].value = ddm
    ws["D10"].number_format = "DD/MM/YYYY"
    lot = ddm.strftime("%d%m%Y")
    ws["O10"].value = lot

    # Fermentation > Date = DDM - 1 an
    ferment_date = ddm - relativedelta(years=1)
    ws["A20"].value = ferment_date
    ws["A20"].number_format = "DD/MM/YYYY"

    # Quantités à produire (cartons)
    c33_1, c75_6_1, c75_4_1 = _agg_cartons_by_format(df_calc, gout1)
    ws["D15"].value = int(c33_1)      # 33cl x12 (France)
    ws["F15"].value = 0               # 33cl x12 (NIKO) -> on laisse à 0 par défaut
    ws["H15"].value = int(c75_6_1)    # 75cl x6
    ws["J15"].value = int(c75_4_1)    # 75cl x4

    if gout2:
        c33_2, c75_6_2, c75_4_2 = _agg_cartons_by_format(df_calc, gout2)
        ws["T15"].value = int(c33_2)    # 33cl x12 (France)
        ws["V15"].value = 0             # 33cl x12 (NIKO)
        ws["X15"].value = int(c75_6_2)  # 75cl x6
        ws["Z15"].value = int(c75_4_2)  # 75cl x4
    else:
        ws["T15"].value = 0; ws["V15"].value = 0; ws["X15"].value = 0; ws["Z15"].value = 0

    # Laisse toutes les autres formules du modèle telles quelles.
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()
