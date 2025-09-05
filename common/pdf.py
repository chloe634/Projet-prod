# common/pdf.py
from io import BytesIO
from datetime import date
from dateutil.relativedelta import relativedelta
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

def _fmt_dmy(d: date) -> str:
    return d.strftime("%d/%m/%Y")

def _agg_by_gout_format(df_calc):
    """
    Regroupe par (Goût, nb bouteilles/carton, volume L) et additionne
    les cartons/bouteilles arrondis. Retourne une liste de lignes.
    """
    import pandas as pd
    if df_calc is None or not isinstance(df_calc, pd.DataFrame) or df_calc.empty:
        return []

    needed = {
        "GoutCanon", "Bouteilles/carton", "Volume bouteille (L)",
        "Cartons à produire (arrondi)", "Bouteilles à produire (arrondi)"
    }
    if any(c not in df_calc.columns for c in needed):
        return []

    grp = (df_calc.groupby(["GoutCanon", "Bouteilles/carton", "Volume bouteille (L)"], dropna=False)[
        ["Cartons à produire (arrondi)", "Bouteilles à produire (arrondi)"]
    ].sum(min_count=1).reset_index())

    rows = []
    for _, r in grp.iterrows():
        gout = str(r["GoutCanon"])
        nb = int(r["Bouteilles/carton"]) if pd.notna(r["Bouteilles/carton"]) else 0
        vol = float(r["Volume bouteille (L)"]) if pd.notna(r["Volume bouteille (L)"]) else 0.0
        ct = int(r["Cartons à produire (arrondi)"]) if pd.notna(r["Cartons à produire (arrondi)"]) else 0
        bt = int(r["Bouteilles à produire (arrondi)"]) if pd.notna(r["Bouteilles à produire (arrondi)"]) else 0
        rows.append([gout, f"{nb} × {vol:.2f} L", ct, bt])

    rows.sort(key=lambda x: (x[0].lower(), x[1]))
    return rows

def generate_production_pdf(
    semaine_du: date,
    ddm: date,
    produit_1: str,
    produit_2: str | None,
    df_calc,
    entreprise: str = "Ferment Station",
    titre_modele: str = "Fiche de production 7000L",
) -> bytes:
    """
    Génère un PDF A4 reprenant l’esprit de la feuille 'Fiche de production 7000L'
    avec les champs variables + un tableau récapitulatif des quantités à produire.
    """
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    W, H = A4
    styles = getSampleStyleSheet()
    style_h = styles["Heading1"]
    style_h.fontSize = 16
    style_h.spaceAfter = 6
    style_p = styles["Normal"]

    # Marges
    margin_x, margin_y = 2*cm, 2*cm
    x = margin_x
    y = H - margin_y

    # En-tête
    c.setFont("Helvetica-Bold", 16)
    c.drawString(x, y, entreprise)
    y -= 12
    c.setFont("Helvetica", 11)
    c.drawString(x, y, f"{titre_modele} — semaine du {_fmt_dmy(semaine_du)}")
    y -= 18
    c.line(x, y, W - margin_x, y)
    y -= 14

    # Bloc champs variables
    lot = _fmt_dmy(ddm).replace("/", "")
    ferment_date = ddm - relativedelta(years=1)

    c.setFont("Helvetica-Bold", 12)
    c.drawString(x, y, "Paramètres de production")
    y -= 12

    c.setFont("Helvetica", 10)
    c.drawString(x, y, f"Produit 1 : {produit_1}")
    y -= 12
    if produit_2:
        c.drawString(x, y, f"Produit 2 : {produit_2}")
        y -= 12
    c.drawString(x, y, f"DDM : {_fmt_dmy(ddm)}")
    y -= 12
    c.drawString(x, y, f"Lot : {lot}")
    y -= 12
    c.drawString(x, y, f"Fermentation — Date : {_fmt_dmy(ferment_date)}  (DDM - 1 an)")
    y -= 16

    # Tableau quantités (agrégation Goût × format)
    rows = _agg_by_gout_format(df_calc)
    if not rows:
        rows = [["—", "—", 0, 0]]

    data = [["Goût", "Format", "Cartons", "Bouteilles"]] + rows

    table = Table(data, colWidths=[7*cm, 4*cm, 3*cm, 3*cm])
    table.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#EFEFEF")),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("ALIGN", (2,1), (-1,-1), "RIGHT"),
        ("ALIGN", (0,0), (-1,0), "CENTER"),
        ("BOTTOMPADDING", (0,0), (-1,0), 6),
        ("TOPPADDING", (0,0), (-1,0), 6),
    ]))

    # Dessin du tableau
    # Descend la position si nécessaire
    max_table_height = 18 * cm
    table.wrapOn(c, W - 2*margin_x, max_table_height)
    table.drawOn(c, x, y - table._height)
    y -= table._height + 10

    # Pied de page
    c.setFont("Helvetica-Oblique", 8)
    c.drawRightString(W - margin_x, margin_y - 6, "Document généré automatiquement")
    c.showPage()
    c.save()
    return buf.getvalue()
