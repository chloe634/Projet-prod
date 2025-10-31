# common/auth_reset.py
from __future__ import annotations
import os, secrets, hashlib, datetime as dt
from typing import Optional, Tuple, Dict, Any, List

from sqlalchemy import text
from db.conn import run_sql

# Réutilise ton hasher de mots de passe existant
from common.auth import get_user_by_email, hash_password  # adapte si besoin

BASE_URL = os.getenv("BASE_URL", "https://ton-domaine.app")  # Kinsta env var
RESET_TTL_MINUTES = int(os.getenv("RESET_TTL_MINUTES", "60"))

def _hash_token(token: str) -> str:
    # On stocke uniquement le hash (pas le token en clair)
    return hashlib.sha256(token.encode()).hexdigest()

def _now_utc() -> dt.datetime:
    return dt.datetime.now(dt.timezone.utc)

def create_password_reset(email: str, meta: Optional[Dict[str, str]]=None) -> Optional[str]:
    """
    Crée un token de reset pour l'utilisateur (si l'email existe).
    Retourne l'URL de reset (avec token) OU None si pas d'utilisateur.
    Ne divulgue pas l'existence ou non de l'email côté UI.
    """
    user = get_user_by_email(email)
    if not user:
        return None

    # Petit rate-limit "maison": max 3 tokens actifs non utilisés, et 1 req/min
    rows = run_sql(text("""
        SELECT created_at FROM password_resets
        WHERE user_id=:uid AND used_at IS NULL AND expires_at > now()
        ORDER BY created_at DESC
        LIMIT 3
    """), {"uid": str(user["id"])})
    if rows:
        last = rows[0]["created_at"]
        if _now_utc() - last < dt.timedelta(seconds=60):
            # Trop fréquent: on ne crée pas de nouveau token, mais on renvoie None (UI dira "email envoyé")
            return None
        if len(rows) >= 3:
            # Trop de tokens actifs, on ne crée pas un nouveau
            return None

    token = secrets.token_urlsafe(32)
    token_hash = _hash_token(token)
    expires_at = _now_utc() + dt.timedelta(minutes=RESET_TTL_MINUTES)

    run_sql(text("""
        INSERT INTO password_resets (user_id, token_hash, expires_at, request_ip, request_ua)
        VALUES (:uid, :th, :exp, :ip, :ua)
    """), {
        "uid": str(user["id"]),
        "th": token_hash,
        "exp": expires_at,
        "ip": (meta or {}).get("ip"),
        "ua": (meta or {}).get("ua"),
    })

    # Lien vers une page dédiée Streamlit
    reset_url = f"{BASE_URL}/_01_Reset_password?token={token}"
    return reset_url

def verify_token(token: str) -> Optional[Dict[str, Any]]:
    """Retourne {user_id, reset_id} si token valide, sinon None."""
    token_hash = _hash_token(token)
    row = run_sql(text("""
        SELECT pr.id as reset_id, pr.user_id
        FROM password_resets pr
        WHERE pr.token_hash=:th AND pr.used_at IS NULL AND pr.expires_at > now()
        ORDER BY pr.id DESC
        LIMIT 1
    """), {"th": token_hash})
    return dict(row[0]) if row else None

def consume_token_and_set_password(reset_id: int, user_id: str, new_password: str) -> None:
    pwd_hash = hash_password(new_password)
    run_sql(text("""
        UPDATE users SET password_hash=:ph WHERE id=:uid;
        UPDATE password_resets SET used_at=now() WHERE id=:rid;
    """), {"ph": pwd_hash, "uid": user_id, "rid": reset_id})
