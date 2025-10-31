-- Extensions nécessaires
CREATE EXTENSION IF NOT EXISTS pgcrypto;  -- pour gen_random_uuid()

-- =========================
-- Tables de base
-- =========================
CREATE TABLE IF NOT EXISTS tenants (
  id         UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  name       TEXT NOT NULL UNIQUE,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS users (
  id            UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  tenant_id     UUID REFERENCES tenants(id) ON DELETE SET NULL,
  email         TEXT NOT NULL UNIQUE,
  password_hash TEXT NOT NULL,
  role          TEXT NOT NULL DEFAULT 'user',
  is_active     BOOLEAN NOT NULL DEFAULT TRUE,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS production_proposals (
  id          UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  tenant_id   UUID REFERENCES tenants(id) ON DELETE CASCADE,
  created_by  UUID REFERENCES users(id) ON DELETE SET NULL,
  payload     JSONB NOT NULL,
  status      TEXT NOT NULL DEFAULT 'draft',
  created_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at  TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- =========================
-- Fonctions utilitaires
-- =========================

-- Normalise les e-mails en minuscules pour garantir l’unicité réelle
CREATE OR REPLACE FUNCTION normalize_email_lower()
RETURNS trigger AS $$
BEGIN
  NEW.email := lower(NEW.email);
  RETURN NEW;
END $$ LANGUAGE plpgsql;

-- Met à jour automatiquement updated_at
CREATE OR REPLACE FUNCTION touch_updated_at()
RETURNS trigger AS $$
BEGIN
  NEW.updated_at = now();
  RETURN NEW;
END $$ LANGUAGE plpgsql;

-- =========================
-- Triggers
-- =========================
DROP TRIGGER IF EXISTS trg_users_email_lower ON users;
CREATE TRIGGER trg_users_email_lower
BEFORE INSERT OR UPDATE ON users
FOR EACH ROW EXECUTE FUNCTION normalize_email_lower();

DROP TRIGGER IF EXISTS trg_pp_touch ON production_proposals;
CREATE TRIGGER trg_pp_touch
BEFORE UPDATE ON production_proposals
FOR EACH ROW EXECUTE FUNCTION touch_updated_at();

-- =========================
-- Index & contraintes
-- =========================

-- Accélère les accès multi-tenant
CREATE INDEX IF NOT EXISTS idx_users_tenant      ON users(tenant_id);
CREATE INDEX IF NOT EXISTS idx_pp_tenant         ON production_proposals(tenant_id);
CREATE INDEX IF NOT EXISTS idx_pp_created_by     ON production_proposals(created_by);

-- Recherche performante sur JSONB
CREATE INDEX IF NOT EXISTS idx_pp_payload_gin    ON production_proposals USING GIN (payload);
-- Recherche par nom (_meta.name) à l’intérieur du JSON
CREATE INDEX IF NOT EXISTS idx_pp_meta_name      ON production_proposals ((payload->'_meta'->>'name'));

-- Recherche par e-mail (déjà normalisé en lowercase par trigger)
CREATE INDEX IF NOT EXISTS idx_users_email_lower ON users (lower(email));

-- Contrainte de rôle autorisé
ALTER TABLE users
  ADD CONSTRAINT users_role_check
  CHECK (role IN ('user','admin'));

-- =========================
-- Password reset tokens
-- =========================
CREATE TABLE IF NOT EXISTS password_resets (
  id          BIGSERIAL PRIMARY KEY,
  user_id     UUID NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  token_hash  TEXT NOT NULL,
  expires_at  TIMESTAMPTZ NOT NULL,
  used_at     TIMESTAMPTZ,
  request_ip  TEXT,
  request_ua  TEXT,
  created_at  TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_password_resets_user  ON password_resets(user_id);
CREATE INDEX IF NOT EXISTS idx_password_resets_token ON password_resets(token_hash);
