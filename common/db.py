import os
from sqlalchemy import create_engine, text

def _build_url() -> str:
    # On récupère d'abord la variable DB_URL que Kinsta fournit automatiquement.
    url = os.getenv("DB_URL")
    if url and url.startswith("postgresql"):
        # Si la connexion n'a pas d'option SSL (sécurité), on ajoute "sslmode=prefer".
        return url if "sslmode=" in url else (url + ("&" if "?" in url else "?") + "sslmode=prefer")

    # Si Kinsta ne donne pas une URL complète, on la reconstruit avec les morceaux :
    host = os.getenv("DB_HOST")
    port = os.getenv("DB_PORT", "5432")
    db   = os.getenv("DB_DATABASE")
    user = os.getenv("DB_USERNAME")
    pwd  = os.getenv("DB_PASSWORD")
    ssl  = os.getenv("DB_SSLMODE", "prefer")  # "prefer" fonctionne en interne sur Kinsta
    return f"postgresql+psycopg2://{user}:{pwd}@{host}:{port}/{db}?sslmode={ssl}"

# On crée le moteur une seule fois pour tout le site
_ENGINE = None

def engine():
    """Renvoie un moteur de connexion SQLAlchemy prêt à l'emploi"""
    global _ENGINE
    if _ENGINE is None:
        _ENGINE = create_engine(_build_url(), pool_pre_ping=True)
    return _ENGINE

def run_sql(sql: str, params: dict | None = None):
    """Exécute une requête SQL directement depuis le code Python"""
    with engine().begin() as conn:
        return conn.execute(text(sql), params or {})
