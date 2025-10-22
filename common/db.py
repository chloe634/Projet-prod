# db.py
import os
from urllib.parse import urlparse, parse_qsl, urlencode, urlunparse
from sqlalchemy import create_engine, text

def _is_internal(host: str | None) -> bool:
    # Kinsta (réseau privé) => host interne Kubernetes
    return bool(host) and host.endswith(".svc.cluster.local")

def _with_sslmode(url: str, sslmode: str) -> str:
    u = urlparse(url)
    qs = dict(parse_qsl(u.query, keep_blank_values=True))
    qs["sslmode"] = sslmode
    new_query = urlencode(qs)
    return urlunparse((u.scheme, u.netloc, u.path, u.params, new_query, u.fragment))

def _build_url() -> str:
    # 1) Si Kinsta te fournit une URL complète
    db_url = os.getenv("DB_URL") or os.getenv("DATABASE_URL")  # fallback classique
    if db_url:
        host = urlparse(db_url).hostname
        if _is_internal(host):
            # Endpoint interne Kinsta -> PAS d’SSL
            return _with_sslmode(db_url, "disable")
        # Endpoint public -> SSL recommandé
        if "sslmode=" in db_url:
            return db_url
        return _with_sslmode(db_url, "require")

    # 2) Sinon, reconstruire à partir des morceaux (on gère plusieurs conventions de noms)
    host = os.getenv("DB_HOST") or os.getenv("POSTGRES_HOST")
    port = os.getenv("DB_PORT") or os.getenv("POSTGRES_PORT") or "5432"
    name = os.getenv("DB_DATABASE") or os.getenv("DB_NAME") or os.getenv("POSTGRES_DB")
    user = os.getenv("DB_USERNAME") or os.getenv("DB_USER") or os.getenv("POSTGRES_USER")
    pwd  = os.getenv("DB_PASSWORD") or os.getenv("POSTGRES_PASSWORD")

    # Choix du sslmode :
    sslmode = os.getenv("DB_SSLMODE")
    if not sslmode:
        sslmode = "disable" if _is_internal(host) else "require"

    return f"postgresql+psycopg2://{user}:{pwd}@{host}:{port}/{name}?sslmode={sslmode}"

_ENGINE = None

def engine():
    """Renvoie un moteur SQLAlchemy prêt à l'emploi."""
    global _ENGINE
    if _ENGINE is None:
        _ENGINE = create_engine(_build_url(), pool_pre_ping=True)
    return _ENGINE

def run_sql(sql: str, params: dict | None = None):
    """Exécute une requête SQL et renvoie le résultat."""
    with engine().begin() as conn:
        return conn.execute(text(sql), params or {})

def ping():
    """Petit test de santé : SELECT 1."""
    try:
        _ = run_sql("SELECT 1;")
        return True, "✅ DB OK (SELECT 1)"
    except Exception as e:
        return False, f"❌ Erreur de connexion : {e}"
