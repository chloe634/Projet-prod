# common/storage.py — VERSION DB (drop-in replacement)
from __future__ import annotations
import json
from datetime import datetime, timezone
from typing import Dict, List, Tuple, Any

import pandas as pd
from sqlalchemy import text
from db.conn import run_sql

# Limite "mémoire longue" par tenant (nombre max de NOMS distincts)
MAX_SLOTS = 6

# Identité par défaut (tu peux changer via variables d'env si tu veux)
DEFAULT_TENANT_NAME = "default"
SYSTEM_EMAIL = "system@symbiose.local"

# ---------- Helpers encodage DataFrame ----------
def _encode_sp(sp: Dict[str, Any]) -> Dict[str, Any]:
    def _df(x):
        return x.to_json(orient="split") if isinstance(x, pd.DataFrame) else None
    return {
        "semaine_du": sp.get("semaine_du"),
        "ddm": sp.get("ddm"),
        "gouts": list(sp.get("gouts", [])),
        "df_min": _df(sp.get("df_min")),
        "df_calc": _df(sp.get("df_calc")),
    }

def _decode_sp(obj: Dict[str, Any]) -> Dict[str, Any]:
    def _df(s):
        return pd.read_json(s, orient="split") if isinstance(s, str) and s.strip() else None
    return {
        "semaine_du": obj.get("semaine_du"),
        "ddm": obj.get("ddm"),
        "gouts": obj.get("gouts") or [],
        "df_min": _df(obj.get("df_min")),
        "df_calc": _df(obj.get("df_calc")),
    }

# ---------- Helpers DB (tenant/user) ----------
def _ensure_tenant(tenant_name: str = DEFAULT_TENANT_NAME) -> str:
    row = run_sql(text("""
        INSERT INTO tenants (name)
        VALUES (:n)
        ON CONFLICT (name) DO UPDATE SET name = EXCLUDED.name
        RETURNING id;
    """), {"n": tenant_name}).mappings().first()
    return row["id"]

def _ensure_user(email: str, tenant_id: str) -> str:
    row = run_sql(text("""
        INSERT INTO users (tenant_id, email, password_hash, role, is_active)
        VALUES (:t, :e, '$local$disabled', 'admin', true)
        ON CONFLICT (email) DO UPDATE SET tenant_id = EXCLUDED.tenant_id
        RETURNING id;
    """), {"t": tenant_id, "e": email}).mappings().first()
    return row["id"]

def _tenant_id() -> str:
    return _ensure_tenant(DEFAULT_TENANT_NAME)

def _system_user_id(tenant_id: str) -> str:
    return _ensure_user(SYSTEM_EMAIL, tenant_id)

# ---------- API publique (identique à l’ancienne) ----------
def list_saved() -> List[Dict[str, Any]]:
    """Retourne [{name, ts, gouts, semaine_du}] triés du plus récent au plus ancien (DB)."""
    t_id = _tenant_id()
    rows = run_sql(text("""
        SELECT id, created_at, updated_at, payload
        FROM production_proposals
        WHERE tenant_id = :t
        ORDER BY created_at DESC
    """), {"t": t_id})
    out: List[Dict[str, Any]] = []
    for r in rows.mappings().all():
        payload = r["payload"] or {}
        meta = payload.get("_meta", {})
        out.append({
            "name": meta.get("name"),
            "ts": meta.get("ts") or (r["created_at"].isoformat() if r.get("created_at") else None),
            "gouts": (payload.get("gouts") or [])[:],
            "semaine_du": payload.get("semaine_du"),
        })
    # déjà trié par created_at DESC ; si tu préfères par meta.ts :
    out.sort(key=lambda x: (x.get("ts") or ""), reverse=True)
    return out

def save_snapshot(name: str, sp: Dict[str, Any]) -> Tuple[bool, str]:
    """Crée / remplace une proposition (MAX_SLOTS par tenant basé sur les NOMS distincts)."""
    name = (name or "").strip()
    if not name:
        return False, "Nom vide."

    t_id = _tenant_id()
    u_id = _system_user_id(t_id)

    # construit le payload applicatif + meta
    payload = _encode_sp(sp)
    ts = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    payload["_meta"] = {"name": name, "ts": ts, "source": "app-db"}

    # Existe déjà ? (match sur _meta.name)
    row = run_sql(text("""
        SELECT id FROM production_proposals
        WHERE tenant_id = :t
          AND payload->'_meta'->>'name' = :n
        ORDER BY updated_at DESC, created_at DESC
        LIMIT 1
    """), {"t": t_id, "n": name}).first()

    if row:
        pid = dict(row._mapping)["id"]
        run_sql(text("""
            UPDATE production_proposals
            SET payload = CAST(:p AS JSONB), updated_at = NOW()
            WHERE id = :id
        """), {"p": json.dumps(payload), "id": pid})
        return True, "Proposition mise à jour."

    # Vérifie la limite MAX_SLOTS sur NOM distinct
    cnt_row = run_sql(text("""
        SELECT COUNT(DISTINCT payload->'_meta'->>'name') AS c
        FROM production_proposals
        WHERE tenant_id = :t
    """), {"t": t_id}).first()
    count = int(dict(cnt_row._mapping)["c"]) if cnt_row else 0
    if count >= MAX_SLOTS:
        return False, f"Limite atteinte ({MAX_SLOTS}). Supprime ou renomme une entrée."

    # Insert (nouvelle entrée)
    run_sql(text("""
        INSERT INTO production_proposals (tenant_id, created_by, payload, status)
        VALUES (:t, :u, CAST(:p AS JSONB), 'draft')
    """), {"t": t_id, "u": u_id, "p": json.dumps(payload)})

    return True, "Proposition enregistrée."

def load_snapshot(name: str) -> Dict[str, Any] | None:
    t_id = _tenant_id()
    row = run_sql(text("""
        SELECT payload FROM production_proposals
        WHERE tenant_id = :t
          AND payload->'_meta'->>'name' = :n
        ORDER BY updated_at DESC, created_at DESC
        LIMIT 1
    """), {"t": t_id, "n": name}).first()
    if not row:
        return None
    payload = dict(row._mapping)["payload"] or {}
    return _decode_sp(payload)

def delete_snapshot(name: str) -> bool:
    t_id = _tenant_id()
    res = run_sql(text("""
        DELETE FROM production_proposals
        WHERE tenant_id = :t
          AND payload->'_meta'->>'name' = :n
        RETURNING id
    """), {"t": t_id, "n": name})
    rows = res.fetchall()
    return len(rows) > 0  # True si au moins une ligne supprimée

def rename_snapshot(old: str, new: str) -> Tuple[bool, str]:
    new = (new or "").strip()
    if not new:
        return False, "Nouveau nom vide."
    t_id = _tenant_id()

    # existe déjà ?
    exists = run_sql(text("""
        SELECT 1 FROM production_proposals
        WHERE tenant_id = :t AND payload->'_meta'->>'name' = :n
        LIMIT 1
    """), {"t": t_id, "n": new}).first()
    if exists:
        return False, "Ce nom existe déjà."

    res = run_sql(text("""
        UPDATE production_proposals
        SET payload = jsonb_set(payload, '{_meta,name}', to_jsonb(:new_name::text), true),
            updated_at = NOW()
        WHERE tenant_id = :t
          AND payload->'_meta'->>'name' = :old_name
        RETURNING id
    """), {"t": t_id, "old_name": old, "new_name": new})
    rows = res.fetchall()
    if len(rows) == 0:
        return False, "Entrée introuvable."
    return True, "Renommée."
