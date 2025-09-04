import os, yaml, pandas as pd
from functools import lru_cache

CONFIG_DEFAULT = {
    "data_files": {
        "main_table": "data/production.xlsx",
        "flavor_map": "data/flavor_map.csv",
    },
    "images_dir": "assets",
}

def load_config() -> dict:
    path = "config.yaml"
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return {**CONFIG_DEFAULT, **(yaml.safe_load(f) or {})}
    return CONFIG_DEFAULT

@lru_cache(maxsize=1)
def get_paths():
    cfg = load_config()
    return (
        cfg["data_files"]["main_table"],
        cfg["data_files"]["flavor_map"],
        cfg["images_dir"],
    )

@lru_cache(maxsize=2)
def read_table():
    main_table, _, _ = get_paths()
    if not os.path.exists(main_table):
        return pd.DataFrame()
    if main_table.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(main_table, header=None)  # brut (pour détecter l'en-tête)
    return pd.read_csv(main_table, sep=";", engine="python", header=None)

@lru_cache(maxsize=2)
def read_flavor_map():
    _, flavor_map, _ = get_paths()
    if not os.path.exists(flavor_map):
        return pd.DataFrame(columns=["name","canonical"])
    # essaie différents séparateurs si besoin
    try:
        return pd.read_csv(flavor_map, encoding="utf-8")
    except Exception:
        return pd.read_csv(flavor_map, encoding="utf-8", sep=";")

