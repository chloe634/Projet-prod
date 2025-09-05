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

def _agg_cartons_by_format(df_calc: pd.DataFrame, gout: str) -> Tuple[int,int,int]:
    """
    Retourne les cartons (arrondis) pour un goût donné par format :
      - 33cl x12  → pour D15/T15
      - 75cl x6   → pour H15/X15
      - 75cl x4   → pour J15/Z15
    """
    if df_calc is None or not isinstance(df_calc, pd.DataFrame) or df_calc.empty:
        return 0, 0, 0

    df = df_calc.copy()
    df = df[df["GoutCanon"].astype(str).str.strip() == str(gout).strip()]
    if df.empty:
        return 0, 0, 0

    # colonnes attendues
    for c in ["Bouteilles/carton", "Volume bouteille (L)", "Cartons à produire (arrondi)"]:
        if c not in df.columns:
            return 0, 0, 0

    # somme par triplet
    c33x12 = int(pd.to_numeric(
        df[(df["Bouteilles/carton"]==12) & (_is_close(df["Volume bouteille (L)"], 0.33))]["Cartons à produire (arrondi)"],
        errors="coerce"
    ).fillna(0).sum())

    c75x6  = int(pd.to_numeric(
        df[(df["Bouteilles/carton"]==6)  & (_is_close(df["Volume bouteille (L)"], 0.75))]["Cartons à produire (arrondi)"],
        errors="coerce"
    ).fillna(0).sum())

    c75x4  = int(pd.to_numeric(
        df[(df["Bouteilles/carton"]==4)  & (_is_close(df["Volume bouteille (L)"], 0.75))]["Cartons à produire (arrondi)"],
        errors="coerce"
    ).fillna(0).sum())

    return c33x12, c75x6, c75x4

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
