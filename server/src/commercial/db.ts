import type { Pool, PoolClient } from "pg";
import { sha256, transaction } from "./core.js";

const schema = `
CREATE SCHEMA IF NOT EXISTS commercial;
CREATE TABLE commercial.users (
 id uuid PRIMARY KEY, external_subject text UNIQUE NOT NULL, active boolean NOT NULL DEFAULT true
);
CREATE TABLE commercial.organizations (id uuid PRIMARY KEY, name text NOT NULL, active boolean NOT NULL DEFAULT true);
CREATE TABLE commercial.memberships (
 org_id uuid REFERENCES commercial.organizations(id), user_id uuid REFERENCES commercial.users(id),
 role text NOT NULL CHECK(role IN ('owner','admin','member','viewer')), active boolean NOT NULL DEFAULT true,
 PRIMARY KEY(org_id,user_id)
);
CREATE TABLE commercial.test_auth_tokens (
 token_hash text PRIMARY KEY CHECK(length(token_hash)=64), user_id uuid REFERENCES commercial.users(id), expires_at timestamptz NOT NULL
);
CREATE TABLE commercial.subscriptions (
 org_id uuid PRIMARY KEY REFERENCES commercial.organizations(id), plan text NOT NULL CHECK(plan IN ('Free','Pro','Business','Team')),
 status text NOT NULL CHECK(status IN ('active','suspended','cancelled')), valid_until timestamptz NOT NULL,
 allow_hosted boolean NOT NULL DEFAULT false, allow_byok boolean NOT NULL DEFAULT false, version integer NOT NULL DEFAULT 1
);
CREATE TABLE commercial.credentials (
 org_id uuid NOT NULL REFERENCES commercial.organizations(id), id uuid NOT NULL, owner_user_id uuid,
 provider text NOT NULL, ciphertext jsonb, hint text NOT NULL, version integer NOT NULL CHECK(version>0), revoked boolean NOT NULL DEFAULT false,
 PRIMARY KEY(org_id,id), FOREIGN KEY(org_id,owner_user_id) REFERENCES commercial.memberships(org_id,user_id)
);
CREATE TABLE commercial.budget_accounts (
 id uuid PRIMARY KEY, scope text NOT NULL CHECK(scope IN ('global','org','user')), org_id uuid REFERENCES commercial.organizations(id), user_id uuid,
 meter text NOT NULL CHECK(meter IN ('usd_micro','service_request')), currency text NOT NULL DEFAULT 'USD' CHECK(currency='USD'),
 period_start timestamptz NOT NULL, period_end timestamptz NOT NULL CHECK(period_end>period_start),
 cap bigint NOT NULL DEFAULT 0 CHECK(cap>=0), spent bigint NOT NULL DEFAULT 0 CHECK(spent>=0), reserved bigint NOT NULL DEFAULT 0 CHECK(reserved>=0),
 frozen boolean NOT NULL DEFAULT false, max_inflight integer NOT NULL DEFAULT 2 CHECK(max_inflight>0),
 CHECK(spent+reserved<=cap),
 CHECK((scope='global' AND org_id IS NULL AND user_id IS NULL) OR (scope='org' AND org_id IS NOT NULL AND user_id IS NULL) OR (scope='user' AND org_id IS NOT NULL AND user_id IS NOT NULL)),
 FOREIGN KEY(org_id,user_id) REFERENCES commercial.memberships(org_id,user_id)
);
CREATE UNIQUE INDEX budget_period_unique ON commercial.budget_accounts(scope,COALESCE(org_id::text,''),COALESCE(user_id::text,''),meter,period_start,period_end);
CREATE TABLE commercial.usage_requests (
 org_id uuid NOT NULL, id uuid NOT NULL, user_id uuid NOT NULL, idempotency_key text NOT NULL, body_hash text NOT NULL,
 state text NOT NULL CHECK(state IN ('reserved','dispatched','settled','released','uncertain')),
 provider text NOT NULL, model text NOT NULL, key_mode text NOT NULL CHECK(key_mode IN ('hosted','byok')),
 credential_id uuid, pricing_version text NOT NULL, bound bigint NOT NULL CHECK(bound>0), provider_bound bigint NOT NULL CHECK(provider_bound>=0),
 actual_provider_cost bigint, usage jsonb, created_at timestamptz NOT NULL DEFAULT now(), updated_at timestamptz NOT NULL DEFAULT now(),
 PRIMARY KEY(org_id,id), UNIQUE(org_id,idempotency_key), FOREIGN KEY(org_id,user_id) REFERENCES commercial.memberships(org_id,user_id),
 FOREIGN KEY(org_id,credential_id) REFERENCES commercial.credentials(org_id,id)
);
CREATE TABLE commercial.allocations (
 org_id uuid NOT NULL, request_id uuid NOT NULL, account_id uuid REFERENCES commercial.budget_accounts(id), amount bigint NOT NULL CHECK(amount>0),
 PRIMARY KEY(org_id,request_id,account_id), FOREIGN KEY(org_id,request_id) REFERENCES commercial.usage_requests(org_id,id)
);
CREATE TABLE commercial.usage_attempts (
 org_id uuid NOT NULL, request_id uuid NOT NULL, id uuid NOT NULL, state text NOT NULL CHECK(state IN ('dispatched','settled','uncertain')),
 dispatched_at timestamptz NOT NULL DEFAULT now(), finished_at timestamptz,
 PRIMARY KEY(org_id,id), UNIQUE(org_id,request_id), FOREIGN KEY(org_id,request_id) REFERENCES commercial.usage_requests(org_id,id)
);
CREATE TABLE commercial.usage_ledger (
 org_id uuid NOT NULL, request_id uuid NOT NULL, kind text NOT NULL CHECK(kind IN ('settled','released','overrun')),
 amount bigint NOT NULL CHECK(amount>=0), provider_cost bigint, created_at timestamptz NOT NULL DEFAULT now(),
 PRIMARY KEY(org_id,request_id,kind), FOREIGN KEY(org_id,request_id) REFERENCES commercial.usage_requests(org_id,id)
);
CREATE TABLE commercial.audit_events (
 id bigserial PRIMARY KEY, org_id uuid NOT NULL REFERENCES commercial.organizations(id), user_id uuid NOT NULL REFERENCES commercial.users(id),
 kind text NOT NULL, reference_id uuid NOT NULL, created_at timestamptz NOT NULL DEFAULT now()
);
CREATE INDEX usage_pending ON commercial.usage_requests(state,updated_at);
`;

/** Transactional migrations, a checksum and a DB advisory lock also work across replicas. */
export async function migrate(pool: Pool): Promise<void> {
  await transaction(pool, async c => {
    await c.query("SELECT pg_advisory_xact_lock(703124801)");
    await c.query("CREATE SCHEMA IF NOT EXISTS commercial");
    await c.query("CREATE TABLE IF NOT EXISTS commercial.schema_migrations(version integer PRIMARY KEY,checksum text NOT NULL,applied_at timestamptz NOT NULL DEFAULT now())");
    const applied = await c.query("SELECT checksum FROM commercial.schema_migrations WHERE version=1");
    if (applied.rowCount) {
      if (applied.rows[0].checksum !== sha256(schema)) throw new Error("migration_checksum_mismatch");
      return;
    }
    await c.query(schema);
    await c.query("INSERT INTO commercial.schema_migrations(version,checksum) VALUES(1,$1)", [sha256(schema)]);
  });
}

/** Only for an empty development/test database; never erase a database with usage. */
export async function rollbackEmptyDatabase(pool: Pool): Promise<void> {
  await transaction(pool, async c => {
    await c.query("SELECT pg_advisory_xact_lock(703124801)");
    const used = await c.query("SELECT 1 FROM commercial.usage_requests LIMIT 1");
    const orgs = await c.query("SELECT 1 FROM commercial.organizations LIMIT 1");
    if (used.rowCount || orgs.rowCount) throw new Error("rollback_requires_empty_database");
    await c.query("DROP SCHEMA commercial CASCADE");
  });
}

export async function audit(c: PoolClient, orgId: string, userId: string, kind: string, reference: string): Promise<void> {
  await c.query("INSERT INTO commercial.audit_events(org_id,user_id,kind,reference_id) VALUES($1,$2,$3,$4)", [orgId, userId, kind, reference]);
}
