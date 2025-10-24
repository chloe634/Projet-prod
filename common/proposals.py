# common/proposals.py (facultatif – remplacement)
import json
from typing import Dict, List, Optional
from sqlalchemy import text
from common.db import run_sql

DEFAULT_TENANT = "default"
SYSTEM_EMAIL = "system@symbiose.local"

def _tenant_id(tenant_name: str = DEFAULT_TENANT) -> str:
    row = run_sql(text("""
        INSERT INTO tenants (name)
        VALUES (:n)
        ON CONFLICT (name) DO UPDATE SET name = EXCLUDED.name
        RETURNING id;
    """), {"n": tenant_name}).mappings().first()
    return row["id"]

def _user_id(email: str, tenant_id: str) -> str:
    row = run_sql(text("""
        INSERT INTO users (tenant_id, email, password_hash, role, is_active)
        VALUES (:t, :e, '$local$disabled', 'admin', true)
        ON CONFLICT (email) DO UPDATE SET tenant_id = EXCLUDED.tenant_id
        RETURNING id;
    """), {"t": tenant_id, "e": email}).mappings().first()
    return row["id"]

# ========= Fonctions d'origine (inchangées de signature) =========
def save_proposal(tenant_id: str, user_id: str | None, payload: Dict):
    run_sql("""
        insert into production_proposals (tenant_id, created_by, payload)
        values (:t, :u, :p::jsonb)
    """, {"t": tenant_id, "u": user_id, "p": json.dumps(payload)})

def list_proposals(tenant_id: str, limit: int = 50) -> List[Dict]:
    rows = run_sql("""
        select id, created_at, status, payload
        from production_proposals
        where tenant_id = :t
        order by created_at desc
        limit :l
    """, {"t": tenant_id, "l": limit})
    return [dict(r._mapping) for r in rows]

# ========= Utilitaires pratiques =========
def proposals_create(payload: Dict, tenant_name: str = DEFAULT_TENANT,
                     created_by_email: str = SYSTEM_EMAIL, status: str = "draft") -> str:
    t = _tenant_id(tenant_name)
    u = _user_id(created_by_email, t)
    row = run_sql(text("""
        INSERT INTO production_proposals (tenant_id, created_by, payload, status)
        VALUES (:t, :u, CAST(:p AS JSONB), :s)
        RETURNING id;
    """), {"t": t, "u": u, "p": json.dumps(payload), "s": status}).mappings().first()
    return row["id"]

def proposals_update_payload(id: str, payload: Dict) -> None:
    run_sql(text("""
        UPDATE production_proposals
        SET payload = CAST(:p AS JSONB), updated_at = NOW()
        WHERE id = :id
    """), {"id": id, "p": json.dumps(payload)})

def proposals_update_status(id: str, status: str) -> None:
    run_sql(text("""
        UPDATE production_proposals
        SET status=:s, updated_at=NOW()
        WHERE id=:id
    """), {"id": id, "s": status})

def proposals_get(id: str) -> Optional[Dict]:
    row = run_sql(text("""
        SELECT id, created_at, updated_at, status, payload
        FROM production_proposals WHERE id=:id
    """), {"id": id}).mappings().first()
    return dict(row) if row else None

def proposals_delete(id: str) -> None:
    run_sql(text("DELETE FROM production_proposals WHERE id=:id"), {"id": id})
