# common/auth.py
from __future__ import annotations
import os, base64, secrets, hashlib, re
from typing import Optional, Dict, Any
from datetime import datetime

import pandas as pd
from sqlalchemy import text
from db.conn import run_sql  # votre helper SQL

# ------------------------------------------------------------------------------
# Helpers DataFrame / SQL
# ------------------------------------------------------------------------------
def _to_df(res) -> pd.DataFrame:
    """Compat: transforme CursorResult (SQLAlchemy 2.x) en DataFrame."""
    try:
        if isinstance(res, pd.DataFrame):
            return res
        return pd.DataFrame(list(res.mappings()))
    except Exception:
        return pd.DataFrame()

# ------------------------------------------------------------------------------
# PBKDF2 (format: pbkdf2_sha256$<iters>$<salt_b64>$<hash_b64>)
# ------------------------------------------------------------------------------
PBKDF2_ALGO = "sha256"
PBKDF2_ITERS = 310_000
SALT_BYTES = 16

def hash_password(password: str) -> str:
    salt = secrets.token_bytes(SALT_BYTES)
    dk = hashlib.pbkdf2_hmac(PBKDF2_ALGO, password.encode("utf-8"), salt, PBKDF2_ITERS)
    return f"pbkdf2_sha256${PBKDF2_ITERS}${base64.b64encode(salt).decode()}${base64.b64encode(dk).decode()}"

def verify_password(password: str, stored: str) -> bool:
    try:
        scheme, iters_s, salt_b64, hash_b64 = stored.split("$", 3)
        if scheme != "pbkdf2_sha256":
            return False
        iters = int(iters_s)
        salt = base64.b64decode(salt_b64)
        expected = base64.b64decode(hash_b64)
        dk = hashlib.pbkdf2_hmac(PBKDF2_ALGO, password.encode("utf-8"), salt, iters)
        return secrets.compare_digest(dk, expected)
    except Exception:
        return False

# ------------------------------------------------------------------------------
# Tenants (résolution nom <-> UUID)
# ------------------------------------------------------------------------------
_UUID_RE = re.compile(
    r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[1-5][0-9a-fA-F]{3}-[89abAB][0-9a-fA-F]{3}-[0-9a-fA-F]{12}$"
)

def get_tenant_by_name(name: str) -> Optional[Dict[str, Any]]:
    """Recherche insensible à la casse."""
    q = text("SELECT id, name, created_at FROM tenants WHERE lower(name)=lower(:n) LIMIT 1")
    df = _to_df(run_sql(q, {"n": name}))
    return None if df.empty else df.iloc[0].to_dict()

def create_tenant(name: str) -> Dict[str, Any]:
    """Crée le tenant et renvoie (id, name, created_at)."""
    q = text("INSERT INTO tenants(name) VALUES (:n) RETURNING id, name, created_at")
    df = _to_df(run_sql(q, {"n": name}))
    return df.iloc[0].to_dict()

def get_or_create_tenant(name: str) -> Dict[str, Any]:
    t = get_tenant_by_name(name)
    return t if t else create_tenant(name)

def ensure_tenant_id(tenant_name_or_id: str) -> str:
    """
    Accepte un UUID ou un nom ; renvoie toujours l'UUID.
    - 'Ferment Station' -> crée/trouve puis renvoie tenants.id (uuid)
    - 'f32b3c7e-....'   -> renvoie tel quel
    """
    t = (tenant_name_or_id or "").strip()
    if not t:
        raise ValueError("Tenant requis.")
    if _UUID_RE.match(t):
        return t
    return get_or_create_tenant(t)["id"]

# ------------------------------------------------------------------------------
# Users
# ------------------------------------------------------------------------------
def find_user_by_email(email: str) -> Optional[Dict[str, Any]]:
    q = text("SELECT * FROM users WHERE lower(email)=lower(:e) LIMIT 1")
    df = _to_df(run_sql(q, {"e": email}))
    return None if df.empty else df.iloc[0].to_dict()

def count_users_in_tenant(tenant_id: str) -> int:
    q = text("SELECT COUNT(*) AS n FROM users WHERE tenant_id=:t")
    df = _to_df(run_sql(q, {"t": tenant_id}))
    return int(df.iloc[0]["n"]) if not df.empty else 0

def create_user(email: str, password: str, tenant_name_or_id: str, role: str = "user") -> Dict[str, Any]:
    """
    Crée un utilisateur (active=True). tenant_name_or_id peut être un nom ou un UUID.
    """
    tenant_id = ensure_tenant_id(tenant_name_or_id)
    q = text("""
        INSERT INTO users(tenant_id, email, password_hash, role, is_active)
        VALUES (:t, lower(:e), :ph, :r, TRUE)
        RETURNING id, tenant_id, email, role, is_active, created_at
    """)
    df = _to_df(run_sql(q, {"t": tenant_id, "e": email, "ph": hash_password(password), "r": role or "user"}))
    return df.iloc[0].to_dict()

def authenticate(email: str, password: str) -> Optional[Dict[str, Any]]:
    q = text("SELECT * FROM users WHERE lower(email)=lower(:e) LIMIT 1")
    df = _to_df(run_sql(q, {"e": email}))
    if df.empty:
        return None
    user = df.iloc[0].to_dict()
    return user if verify_password(password, user["password_hash"]) else None

def set_user_role(user_id: str, role: str) -> None:
    run_sql(text("UPDATE users SET role=:r WHERE id=:id"), {"r": role, "id": user_id})

def change_password(user_id: str, new_password: str) -> None:
    run_sql(text("UPDATE users SET password_hash=:ph WHERE id=:id"),
            {"ph": hash_password(new_password), "id": user_id})
