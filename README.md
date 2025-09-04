# Ferment Station — Streamlit (multi-pages)

L'app lit **uniquement** les fichiers du repo (`/data`, `/assets`).  
Aucune importation locale n'est nécessaire.

## Structure
- `app.py` (accueil)
- `pages/01_Production.py`, `pages/02_Optimisation.py`, `pages/03_Fiche_de_ramasse.py`
- `common/design.py` (thème & UI)
- `common/data.py` (config & chemins)
- `core/optimizer.py` (algorithmes)
- `data/production.xlsx`, `data/flavor_map.csv`
- `assets/` (images produits)

## Lancer en local (optionnel)
```bash
pip install -r requirements.txt
streamlit run app.py
