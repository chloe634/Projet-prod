# pages/98_Import_JSON.py
import json
import pathlib
import streamlit as st
from sqlalchemy import text
from db.conn import run_sql  # tu as remplacé common.db par db.conn, parfait

st.set_page_config(page_title="Import mémoire JSON → DB", page_icon="⬆️", layout="wide")
st.title("⬆️ Importer l'ancienne mémoire (JSON) vers PostgreSQL")

DATA_PATH = pathlib.Path(__file__).resolve().parents[1] / "data" / "memoire_longue.json"

st.info(f"Fichier attendu : {DATA_PATH}")

if not DATA_PATH.exists():
    st.error("Fichier introuvable. Vérifie que data/memoire_longue.json est bien dans le repo.")
    st.stop()

TENANT_NAME = "default"
SYSTEM_EMAIL = "system@symbiose.local"
SYSTEM_PWD_HASH = "$local$disabled"  # placeholder non utilisable pour login

if st.button("Créer/assurer tenant & user système"):
    # upsert tenant
    row_t = run_sql(text("""
        INSERT INTO tenants (name)
        VALUES (:name)
        ON CONFLICT (name) DO UPDATE SET name = EXCLUDED.name
        RETURNING id;
    """), {"name": TENANT_NAME}).mappings().first()
    tenant_id = row_t["id"]

    # upsert user système
    row_u = run_sql(text("""
        INSERT INTO users (tenant_id, email, password_hash, role, is_active)
        VALUES (:tenant_id, :email, :pwd, 'admin', true)
        ON CONFLICT (email) DO UPDATE SET tenant_id = EXCLUDED.tenant_id
        RETURNING id;
    """), {"tenant_id": tenant_id, "email": SYSTEM_EMAIL, "pwd": SYSTEM_PWD_HASH}).mappings().first()

    st.success(f"OK — tenant '{TENANT_NAME}' et user '{SYSTEM_EMAIL}' prêts.")

st.divider()

if st.button("Importer maintenant le JSON → production_proposals"):
    data = json.load(open(DATA_PATH, "r", encoding="utf-8"))

    # Récupérer tenant_id & user_id
    tenant = run_sql(text("SELECT id FROM tenants WHERE name=:n"), {"n": TENANT_NAME}).mappings().first()
    user   = run_sql(text("SELECT id FROM users WHERE email=:e"),  {"e": SYSTEM_EMAIL}).mappings().first()
    if not tenant or not user:
        st.error("Assure d'abord le tenant & le user (bouton au-dessus).")
        st.stop()
    tenant_id = tenant["id"]
    user_id   = user["id"]

    inserted = 0
    for item in data:
        # Prépare le payload et garde les métadonnées d’origine dans _meta
        payload = dict(item.get("payload") or {})
        payload["_meta"] = {
            "name": item.get("name"),
            "ts": item.get("ts"),
            "source": "legacy-json"
        }

        # ⬇️ PATCH appliqué : CAST(:ts AS timestamptz) au lieu de :ts::timestamptz
        sql = text("""
            INSERT INTO production_proposals
                (tenant_id, created_by, payload, status, created_at, updated_at)
            VALUES
                (:tenant_id, :created_by, CAST(:payload AS JSONB), 'draft',
                 COALESCE(CAST(:ts AS timestamptz), NOW()),
                 COALESCE(CAST(:ts AS timestamptz), NOW()))
            RETURNING id;
        """)
        run_sql(sql, {
            "tenant_id": tenant_id,
            "created_by": user_id,
            "payload": json.dumps(payload),
            "ts": item.get("ts")  # ex: "2025-09-24T13:36:38Z"
        })
        inserted += 1

    st.success(f"✅ Import terminé : {inserted} propositions insérées.")
    st.info("Tu peux supprimer cette page ensuite (one-shot).")
