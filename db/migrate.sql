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

