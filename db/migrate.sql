create extension if not exists pgcrypto;

create table if not exists tenants (
  id uuid primary key default gen_random_uuid(),
  name text not null unique,
  created_at timestamptz not null default now()
);

create table if not exists users (
  id uuid primary key default gen_random_uuid(),
  tenant_id uuid references tenants(id) on delete set null,
  email text not null unique,
  password_hash text not null,
  role text not null default 'user',
  is_active boolean not null default true,
  created_at timestamptz not null default now()
);

create table if not exists production_proposals (
  id uuid primary key default gen_random_uuid(),
  tenant_id uuid references tenants(id) on delete cascade,
  created_by uuid references users(id) on delete set null,
  payload jsonb not null,
  status text not null default 'draft',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists idx_pp_tenant on production_proposals(tenant_id);

-- Accélère les filtres sur payload JSON
CREATE INDEX IF NOT EXISTS idx_pp_payload_gin ON production_proposals USING GIN (payload);
-- Accélère la recherche par nom (_meta.name)
CREATE INDEX IF NOT EXISTS idx_pp_meta_name ON production_proposals ((payload->'_meta'->>'name'));


-- === AUTH & MULTI-TENANT HARDENING ========================================

-- Emails en minuscule pour l’unicité "vraie"
CREATE OR REPLACE FUNCTION normalize_email_lower()
RETURNS trigger AS $$
BEGIN
  NEW.email := lower(NEW.email);
  RETURN NEW;
END $$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS trg_users_email_lower ON users;
CREATE TRIGGER trg_users_email_lower
BEFORE INSERT OR UPDATE ON users
FOR EACH ROW EXECUTE FUNCTION normalize_email_lower();

-- Index utiles
CREATE INDEX IF NOT EXISTS idx_users_tenant    ON users(tenant_id);
CREATE INDEX IF NOT EXISTS idx_users_email     ON users(lower(email));
CREATE INDEX IF NOT EXISTS idx_pp_created_by   ON production_proposals(created_by);
CREATE INDEX IF NOT EXISTS idx_pp_tenant       ON production_proposals(tenant_id);

-- Contrainte rôle
ALTER TABLE users
  ADD CONSTRAINT users_role_check
  CHECK (role IN ('user','admin'));

-- Auto mise à jour updated_at
CREATE OR REPLACE FUNCTION touch_updated_at() RETURNS trigger AS $$
BEGIN
  NEW.updated_at = now();
  RETURN NEW;
END $$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS trg_pp_touch ON production_proposals;
CREATE TRIGGER trg_pp_touch
BEFORE UPDATE ON production_proposals
FOR EACH ROW EXECUTE FUNCTION touch_updated_at();

-- Password reset tokens
CREATE TABLE IF NOT EXISTS password_resets (
  id BIGSERIAL PRIMARY KEY,
  user_id UUID NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  token_hash TEXT NOT NULL,
  expires_at TIMESTAMPTZ NOT NULL,
  used_at TIMESTAMPTZ,
  request_ip TEXT,
  request_ua TEXT,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS idx_password_resets_user ON password_resets(user_id);
CREATE INDEX IF NOT EXISTS idx_password_resets_token ON password_resets(token_hash);
