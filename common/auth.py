# common/auth.py
from __future__ import annotations
import os, base64, secrets, hashlib
from typing import Optional, Dict, Any
from datetime import datetime
import pandas as pd
from sqlalchemy import text
from db.conn import run_sql  # tu l'utilises déjà ailleurs

PBKDF2_ALGO = "sha256"
PBKDF2_ITERS = 310_000
SALT_BYTES = 16

# --------------------------------------------------------------------------
# Helpers PBKDF2 (pas de dépendance externe, OK pour Kinsta)
# format stocké: "pbkdf2_sha256$<iters>$<salt_b64>$<hash_b64>"
# --------------------------------------------------------------------------
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
        # compare constant-time
        return secrets.compare_digest(dk, expected)
    except Exception:
        return False

# --------------------------------------------------------------------------
# Tenants
# --------------------------------------------------------------------------
def get_tenant_by_name(name: str) -> Optional[Dict[str, Any]]:
    sql = text("SELECT id, name, created_at FROM tenants WHERE name=:n")
    df = run_sql(sql, {"n": name})
    return None if df.empty else df.iloc[0].to_dict()

def create_tenant(name: str) -> Dict[str, Any]:
    sql = text("""
        INSERT INTO tenants(name)
        VALUES (:n)
        ON CONFLICT(name) DO UPDATE SET name=EXCLUDED.name
        RETURNING id, name, created_at
    """)
    df = run_sql(sql, {"n": name})
    return df.iloc[0].to_dict()

def get_or_create_tenant(name: str) -> Dict[str, Any]:
    return get_tenant_by_name(name) or create_tenant(name)

# --------------------------------------------------------------------------
# Users
# --------------------------------------------------------------------------
def find_user_by_email(email: str) -> Optional[Dict[str, Any]]:
    sql = text("""
        SELECT id, tenant_id, email, password_hash, role, is_active, created_at
        FROM users
        WHERE lower(email)=lower(:e)
        LIMIT 1
    """)
    df = run_sql(sql, {"e": email})
    return None if df.empty else df.iloc[0].to_dict()

def count_users_in_tenant(tenant_id: str) -> int:
    sql = text("SELECT count(*) AS c FROM users WHERE tenant_id=:t")
    df = run_sql(sql, {"t": tenant_id})
    return int(df.iloc[0]["c"])

def create_user(email: str, password: str, tenant_name: str, role: Optional[str] = None) -> Dict[str, Any]:
    tenant = get_or_create_tenant(tenant_name)
    tenant_id = tenant["id"]

    # premier utilisateur du tenant => admin par défaut (sinon 'user')
    final_role = role or ("admin" if count_users_in_tenant(tenant_id)==0 else "user")

    sql = text("""
        INSERT INTO users(tenant_id, email, password_hash, role, is_active)
        VALUES (:t, lower(:e), :ph, :r, TRUE)
        RETURNING id, tenant_id, email, role, is_active, created_at
    """)
    df = run_sql(sql, {"t": tenant_id, "e": email, "ph": hash_password(password), "r": final_role})
    return df.iloc[0].to_dict()

def authenticate(email: str, password: str) -> Optional[Dict[str, Any]]:
    u = find_user_by_email(email)
    if not u or not u.get("is_active"):
        return None
    if not verify_password(password, u.get("password_hash","")):
        return None
    # on ne renvoie pas le hash
    u.pop("password_hash", None)
    return u

def set_user_role(user_id: str, role: str) -> None:
    sql = text("UPDATE users SET role=:r WHERE id=:id")
    run_sql(sql, {"r": role, "id": user_id})

def change_password(user_id: str, new_password: str) -> None:
    sql = text("UPDATE users SET password_hash=:ph WHERE id=:id")
    run_sql(sql, {"ph": hash_password(new_password), "id": user_id})
