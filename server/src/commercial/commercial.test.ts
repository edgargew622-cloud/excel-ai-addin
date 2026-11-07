import { before, beforeEach, after, test } from "node:test";
import assert from "node:assert/strict";
import { randomBytes } from "node:crypto";
import { spawnSync } from "node:child_process";
import { request as httpRequest, type Server } from "node:http";
import { poolFor, id, sha256, readConfig, type Context } from "./core.js";
import { migrate, rollbackEmptyDatabase } from "./db.js";
import { TestIdentityVerifier } from "./identity.js";
import { Ledger, type ReserveCommand } from "./ledger.js";
import { CredentialStore, DevKeyringProtector } from "./secrets.js";
import { ChatService } from "./service.js";
import { FakeProvider, parseChat, type ProviderAdapter } from "./providers.js";
import { createCommercialApp, type AuditMetric } from "./app.js";

const databaseUrl = process.env.COMMERCIAL_TEST_DATABASE_URL;
if (!databaseUrl || !new URL(databaseUrl).pathname.endsWith("_c1_test")) throw new Error("Set COMMERCIAL_TEST_DATABASE_URL to an isolated *_c1_test database; this suite resets only that test schema.");
const pool = poolFor(databaseUrl);
const master = randomBytes(32);
const credentials = new CredentialStore(pool, new DevKeyringProtector(new Map([["v1", master]]), "v1"));
let ledger = new Ledger(pool);
let a: Context, b: Context, member: Context, viewer: Context;
let tokenA: string, tokenB: string, tokenMember: string;
const input = () => ({ provider: "fake", model: "c1-test", keyMode: "hosted", messages: [{ role: "user", content: "synthetic request" }] });
const command = (key = id(), mode: "hosted" | "byok" = "hosted"): ReserveCommand => ({ idempotencyKey: key, bodyHash: sha256(key), provider: "fake", model: "c1-test", mode, providerBound: 10n, pricingVersion: "fake-v1" });

before(async () => { await migrate(pool); });
beforeEach(async () => {
  await pool.query("TRUNCATE commercial.users,commercial.organizations,commercial.budget_accounts,commercial.audit_events RESTART IDENTITY CASCADE");
  a = { orgId: id(), userId: id(), role: "owner" };
  b = { orgId: id(), userId: id(), role: "owner" };
  member = { orgId: a.orgId, userId: id(), role: "member" };
  viewer = { orgId: a.orgId, userId: id(), role: "viewer" };
  tokenA = randomBytes(32).toString("base64url"); tokenB = randomBytes(32).toString("base64url"); tokenMember = randomBytes(32).toString("base64url");
  for (const ctx of [a, b, member, viewer]) await pool.query("INSERT INTO commercial.users(id,external_subject) VALUES($1,$2)", [ctx.userId, `test:${ctx.userId}`]);
  for (const ctx of [a, b]) await pool.query("INSERT INTO commercial.organizations(id,name) VALUES($1,'synthetic')", [ctx.orgId]);
  for (const ctx of [a, b, member, viewer]) await pool.query("INSERT INTO commercial.memberships(org_id,user_id,role) VALUES($1,$2,$3)", [ctx.orgId, ctx.userId, ctx.role]);
  for (const [ctx, token] of [[a, tokenA], [b, tokenB], [member, tokenMember]] as const) await pool.query("INSERT INTO commercial.test_auth_tokens(token_hash,user_id,expires_at) VALUES($1,$2,now()+interval '1 hour')", [sha256(token), ctx.userId]);
  for (const ctx of [a, b]) await pool.query("INSERT INTO commercial.subscriptions(org_id,plan,status,valid_until,allow_hosted,allow_byok) VALUES($1,'Team','active',now()+interval '1 day',true,true)", [ctx.orgId]);
  for (const meter of ["usd_micro", "service_request"]) {
    await pool.query("INSERT INTO commercial.budget_accounts(id,scope,meter,period_start,period_end,cap,max_inflight) VALUES($1,'global',$2,now()-interval '1 day',now()+interval '1 day',10000,200)", [id(), meter]);
    for (const ctx of [a, b]) await pool.query("INSERT INTO commercial.budget_accounts(id,scope,org_id,meter,period_start,period_end,cap,max_inflight) VALUES($1,'org',$2,$3,now()-interval '1 day',now()+interval '1 day',1000,200)", [id(), ctx.orgId, meter]);
    for (const ctx of [a, b, member, viewer]) await pool.query("INSERT INTO commercial.budget_accounts(id,scope,org_id,user_id,meter,period_start,period_end,cap,max_inflight) VALUES($1,'user',$2,$3,$4,now()-interval '1 day',now()+interval '1 day',1000,200)", [id(), ctx.orgId, ctx.userId, meter]);
  }
  ledger = new Ledger(pool);
});
after(async () => { await pool.end(); });

async function rejectsCode(run: Promise<unknown>, code: string) {
  await assert.rejects(run, (e: unknown) => (e as { code?: string }).code === code);
}
const service = (provider: ProviderAdapter = new FakeProvider()) => new ChatService(ledger, credentials, new Map([[provider.id, provider]]));

test("C1 config is zero-spend by default, refuses production/test auth and public bind", () => {
  assert.equal(readConfig({}).enabled, false);
  const valid = { COMMERCIAL_ENABLED: "1", COMMERCIAL_AUTH_MODE: "test", COMMERCIAL_DATABASE_URL: databaseUrl, COMMERCIAL_DEV_MASTER_KEY: master.toString("hex") };
  assert.equal(readConfig(valid).allowFake, false);
  assert.throws(() => readConfig({ ...valid, NODE_ENV: "production" }));
  assert.throws(() => readConfig({ ...valid, COMMERCIAL_BIND: "0.0.0.0" }));
  assert.throws(() => readConfig({ ...valid, COMMERCIAL_DATABASE_URL: "postgresql://user:synthetic@remote.example/db" }));
  assert.throws(() => readConfig({ ...valid, COMMERCIAL_DEV_MASTER_KEY: "bad" }));
});

test("100 concurrent requests with one-request budget dispatch exactly one upstream", async () => {
  await pool.query("UPDATE commercial.budget_accounts SET cap=10 WHERE meter='usd_micro'");
  let calls = 0, rejected = 0;
  let finish!: () => void;
  const gate = new Promise<void>(r => { finish = r; });
  const fake = new FakeProvider();
  const provider: ProviderAdapter = { id: "fake", quote: i => fake.quote(i), async run(i, o) { calls++; await gate; return fake.run(i, o); } };
  const chat = service(provider);
  const all = Array.from({ length: 100 }, () => chat.chat(a, input(), id(), new AbortController().signal).catch(e => { assert.equal(e.code, "budget_exhausted"); rejected++; return null; }));
  try {
    const end = Date.now() + 20_000;
    while (rejected < 99 && Date.now() < end) await new Promise(r => setTimeout(r, 20));
    assert.equal(rejected, 99); assert.equal(calls, 1);
  } finally { finish(); }
  const results = await Promise.all(all);
  assert.equal(results.filter(Boolean).length, 1);
  assert.equal((await pool.query("SELECT count(*)::integer n FROM commercial.usage_attempts")).rows[0].n, 1);
  assert.equal((await pool.query("SELECT count(*)::integer n FROM commercial.budget_accounts WHERE spent+reserved>cap")).rows[0].n, 0);
});

test("100 duplicate IDs create one reservation, attempt and settlement; changed body conflicts", async () => {
  const fake = new FakeProvider(), chat = service(fake), key = id();
  const all = await Promise.all(Array.from({ length: 100 }, () => chat.chat(a, input(), key, new AbortController().signal)));
  assert.equal(fake.calls, 1); assert.equal(new Set(all.map(x => x.request.id)).size, 1);
  assert.equal((await pool.query("SELECT count(*)::integer n FROM commercial.usage_ledger WHERE kind='settled'")).rows[0].n, 1);
  await rejectsCode(chat.chat(a, { ...input(), messages: [{ role: "user", content: "different" }] }, key, new AbortController().signal), "idempotency_conflict");
});

test("quota checks apply independently to global, organization and user; absent and zero budgets fail closed", async () => {
  for (const scope of ["global", "org", "user"]) {
    await pool.query("UPDATE commercial.budget_accounts SET cap=1000");
    await pool.query("UPDATE commercial.budget_accounts SET cap=0 WHERE scope=$1 AND meter='usd_micro'", [scope]);
    await rejectsCode(ledger.reserve(a, command()), "budget_exhausted");
  }
  await pool.query("DELETE FROM commercial.budget_accounts WHERE scope='user' AND user_id=$1", [a.userId]);
  await rejectsCode(ledger.reserve(a, command()), "budget_not_configured");
});

test("active database membership, role, license expiry and suspension override supplied context", async () => {
  await rejectsCode(ledger.reserve(viewer, command()), "role_cannot_spend");
  await rejectsCode(ledger.reserve({ ...a, orgId: b.orgId }, command()), "organization_not_found");
  await pool.query("UPDATE commercial.subscriptions SET valid_until=now()-interval '1 second' WHERE org_id=$1", [a.orgId]);
  await rejectsCode(ledger.reserve(a, command()), "entitlement_required");
  await pool.query("UPDATE commercial.subscriptions SET valid_until=now()+interval '1 hour',status='suspended' WHERE org_id=$1", [a.orgId]);
  await rejectsCode(ledger.reserve(a, command()), "entitlement_required");
  await pool.query("UPDATE commercial.memberships SET active=false WHERE org_id=$1", [a.orgId]);
  await rejectsCode(ledger.reserve(a, command()), "organization_not_found");
});

test("dispatch fence disallows duplicate calls and checks revoked membership again", async () => {
  const r = (await ledger.reserve(a, command())).request;
  await ledger.dispatch(a, r.id);
  await rejectsCode(ledger.dispatch(a, r.id), "already_dispatched");
  await rejectsCode(ledger.release(a, r.id), "dispatched_budget_cannot_be_released");
  const next = (await ledger.reserve(a, command())).request;
  await pool.query("UPDATE commercial.memberships SET active=false WHERE org_id=$1", [a.orgId]);
  await rejectsCode(ledger.dispatch(a, next.id), "organization_not_found");
});

test("durable restart recovery releases only undispatched reservations and retains uncertain spend", async () => {
  const reserved = (await ledger.reserve(a, command())).request;
  const dispatched = (await ledger.reserve(a, command())).request;
  await ledger.dispatch(a, dispatched.id);
  const restartedPool = poolFor(databaseUrl!);
  try {
    const restarted = new Ledger(restartedPool);
    assert.equal(await restarted.recoverStale(0), 2);
    assert.equal((await restarted.get(a, reserved.id)).state, "released");
    assert.equal((await restarted.get(a, dispatched.id)).state, "uncertain");
    assert.equal((await pool.query("SELECT reserved::text FROM commercial.budget_accounts WHERE scope='org' AND org_id=$1 AND meter='usd_micro'", [a.orgId])).rows[0].reserved, "10");
    await restarted.settle(a, dispatched.id, 3n, { promptTokens: 2, completionTokens: 2, totalTokens: 4 });
    await restarted.settle(a, dispatched.id, 3n);
    assert.equal((await pool.query("SELECT spent::text,reserved::text FROM commercial.budget_accounts WHERE scope='org' AND org_id=$1 AND meter='usd_micro'", [a.orgId])).rows[0].spent, "3");
    await rejectsCode(restarted.settle(a, dispatched.id, 4n), "settlement_conflict");
  } finally { await restartedPool.end(); }
});

test("missing usage, cancellation and a provider failure keep reserve and never auto retry", async () => {
  for (const mode of ["missing", "throw", "abort"]) {
    const fake = new FakeProvider(); let calls = 0;
    const abort = new AbortController();
    const p: ProviderAdapter = { id: "fake", quote: i => fake.quote(i), async run(i, o) {
      calls++;
      if (mode === "throw") throw new Error("SYNTHETIC_SECRET_MUST_NOT_BE_LOGGED");
      if (mode === "abort") { abort.abort(); return new Promise(() => undefined); }
      return { ...(await fake.run(i, o)), usage: null };
    } };
    const chat = service(p), key = id();
    await chat.chat(a, input(), key, abort.signal).catch(() => undefined);
    const retry = await chat.chat(a, input(), key, new AbortController().signal);
    assert.equal(retry.request.state, "uncertain"); assert.equal(calls, 1);
  }
});

test("cost overrun freezes shared accounts, preserves reserve and emits an incident ledger entry", async () => {
  const r = (await ledger.reserve(a, command())).request;
  await ledger.dispatch(a, r.id);
  assert.equal((await ledger.settle(a, r.id, 11n)).state, "uncertain");
  assert.equal((await pool.query("SELECT count(*)::integer n FROM commercial.budget_accounts WHERE frozen")).rows[0].n, 3);
  await rejectsCode(ledger.reserve(b, command()), "budget_frozen");
  assert.equal((await pool.query("SELECT provider_cost::text FROM commercial.usage_ledger WHERE kind='overrun'")).rows[0].provider_cost, "11");
});

test("a real process crash after dispatch survives a fresh process/pool without releasing potentially billed spend", async () => {
  const coreUrl = new URL("./core.js", import.meta.url).href;
  const ledgerUrl = new URL("./ledger.js", import.meta.url).href;
  const code = `
    const { poolFor } = await import(${JSON.stringify(coreUrl)});
    const { Ledger } = await import(${JSON.stringify(ledgerUrl)});
    const db = poolFor(process.env.COMMERCIAL_TEST_DATABASE_URL);
    const ledger = new Ledger(db);
    const ctx = ${JSON.stringify(a)};
    const r = await ledger.reserve(ctx,{...${JSON.stringify({ ...command(), providerBound: "10" })},providerBound:10n});
    await ledger.dispatch(ctx,r.request.id);
    await new Promise(resolve => process.stdout.write(r.request.id, resolve));
    process.exit(23);
  `;
  const child = spawnSync(process.execPath, ["--input-type=module", "-e", code], { encoding: "utf8", timeout: 15_000 });
  assert.equal(child.status, 23);
  const requestId = child.stdout.trim();
  const restartedPool = poolFor(databaseUrl!);
  try {
    const restarted = new Ledger(restartedPool);
    assert.equal((await restarted.get(a, requestId)).state, "dispatched");
    await restarted.recoverStale(0);
    assert.equal((await restarted.get(a, requestId)).state, "uncertain");
    assert.equal((await pool.query("SELECT reserved::text FROM commercial.budget_accounts WHERE scope='org' AND org_id=$1 AND meter='usd_micro'", [a.orgId])).rows[0].reserved, "10");
  } finally { await restartedPool.end(); }
});

test("credentials are tenant/private scoped, encrypted, rotated and revoked without plaintext in metadata", async () => {
  const sentinel = "SYNTHETIC_PRIVATE_KEY_abcdef1234";
  const key = await credentials.create(a, "fake", sentinel);
  assert.ok(!JSON.stringify(key).includes(sentinel));
  assert.equal((await credentials.resolve(a, key.id, "fake")).value, sentinel);
  await rejectsCode(credentials.resolve(b, key.id, "fake"), "credential_not_found");
  await rejectsCode(credentials.resolve(member, key.id, "fake"), "credential_not_found");
  assert.equal((await credentials.list(b)).length, 0);
  const raw = (await pool.query("SELECT ciphertext FROM commercial.credentials WHERE id=$1", [key.id])).rows[0].ciphertext;
  assert.ok(!JSON.stringify(raw).includes(sentinel));
  await credentials.rotate(a, key.id, "SYNTHETIC_ROTATED_KEY_12345678");
  assert.equal((await credentials.resolve(a, key.id, "fake")).version, 2);
  await credentials.revoke(a, key.id);
  await rejectsCode(credentials.resolve(a, key.id, "fake"), "credential_not_found");
  assert.equal((await pool.query("SELECT ciphertext FROM commercial.credentials WHERE id=$1", [key.id])).rows[0].ciphertext, null);
});

test("encryption context prevents ciphertext swaps; shared keys require administration; key change fences dispatch", async () => {
  await rejectsCode(credentials.create(member, "fake", "synthetic-key-12345", true), "credential_permission_denied");
  const ka = await credentials.create(a, "fake", "synthetic-key-a-12345", true);
  const kb = await credentials.create(b, "fake", "synthetic-key-b-12345", true);
  await pool.query("UPDATE commercial.credentials SET ciphertext=(SELECT ciphertext FROM commercial.credentials WHERE org_id=$1 AND id=$2) WHERE org_id=$3 AND id=$4", [a.orgId, ka.id, b.orgId, kb.id]);
  await rejectsCode(credentials.resolve(b, kb.id, "fake"), "credential_unavailable");
  const cmd = { ...command(id(), "byok"), credentialId: ka.id };
  const r = (await ledger.reserve(a, cmd)).request;
  await credentials.rotate(a, ka.id, "synthetic-key-a-rotated");
  await rejectsCode(ledger.dispatch(a, r.id, 1), "credential_changed");
  await ledger.release(a, r.id);
});

test("BYOK records provider spend separately and consumes service quota, not hosted credits", async () => {
  const key = await credentials.create(a, "fake", "synthetic-byok-12345678", true);
  const result = await service().chat(a, { ...input(), keyMode: "byok", credentialId: key.id }, id(), new AbortController().signal);
  assert.equal(result.request.state, "settled");
  const rows = (await pool.query("SELECT meter,spent::text,reserved::text FROM commercial.budget_accounts WHERE scope='org' AND org_id=$1 ORDER BY meter", [a.orgId])).rows;
  assert.deepEqual(rows, [{ meter: "service_request", spent: "1", reserved: "0" }, { meter: "usd_micro", spent: "0", reserved: "0" }]);
  assert.equal(result.request.actual_provider_cost, "3");
});

test("pre-dispatch vault failure releases reservation without an upstream attempt", async () => {
  const key = await credentials.create(a, "fake", "synthetic-byok-before-dispatch", true);
  await pool.query("UPDATE commercial.credentials SET ciphertext=$1 WHERE org_id=$2 AND id=$3", [{ keyVersion: "missing", iv: "bad", data: "bad", tag: "bad" }, a.orgId, key.id]);
  const fake = new FakeProvider();
  await rejectsCode(service(fake).chat(a, { ...input(), keyMode: "byok", credentialId: key.id }, id(), new AbortController().signal), "credential_unavailable");
  assert.equal(fake.calls, 0);
  assert.equal((await pool.query("SELECT state FROM commercial.usage_requests")).rows[0].state, "released");
  assert.equal((await pool.query("SELECT count(*)::integer n FROM commercial.usage_attempts")).rows[0].n, 0);
  assert.equal((await pool.query("SELECT sum(reserved)::text n FROM commercial.budget_accounts")).rows[0].n, "0");
});

test("period rollover settles against original accounts, concurrency caps cannot be bypassed", async () => {
  await pool.query("UPDATE commercial.budget_accounts SET max_inflight=1 WHERE scope='org' AND org_id=$1", [a.orgId]);
  const r = (await ledger.reserve(a, command())).request;
  await rejectsCode(ledger.reserve(a, command()), "concurrency_limit");
  await ledger.dispatch(a, r.id);
  await pool.query("UPDATE commercial.budget_accounts SET period_start=now()-interval '2 days',period_end=now()-interval '1 day'");
  await ledger.settle(a, r.id, 3n);
  assert.equal((await pool.query("SELECT sum(spent)::text n FROM commercial.budget_accounts WHERE meter='usd_micro'")).rows[0].n, "9");
  await rejectsCode(ledger.reserve(a, command()), "budget_not_configured");
});

test("HTTP API rejects foreign tenants, unsafe host/origin, unauthenticated bodies and forged prices; errors/logs redact secrets", async () => {
  const logs: AuditMetric[] = [];
  const fake = new FakeProvider();
  const { app } = createCommercialApp({ pool, identity: new TestIdentityVerifier(pool), credentials, allowFake: true, adapters: new Map([["fake", fake]]), log: m => logs.push(m) });
  const server: Server = app.listen(0, "127.0.0.1");
  await new Promise<void>(r => server.once("listening", r));
  const port = (server.address() as { port: number }).port;
  const request = (path: string, token = tokenA, method = "GET", body?: unknown, extra?: Record<string, string>) => fetch(`http://127.0.0.1:${port}/api/v1${path}`, { method, headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...extra }, ...(body === undefined ? {} : { body: JSON.stringify(body) }) });
  try {
    assert.equal((await request(`/orgs/${b.orgId}/keys`)).status, 404);
    assert.equal((await request(`/orgs/${a.orgId}/keys`, tokenB)).status, 404);
    assert.equal((await request("/me", "invalid")).status, 401);
    assert.equal((await request("/me", tokenA, "GET", undefined, { Origin: "https://evil.example" })).status, 403);
    const foreignHostStatus = await new Promise<number>(resolve => {
      const r = httpRequest({ host: "127.0.0.1", port, path: "/api/v1/me", headers: { Host: "evil.example", Authorization: `Bearer ${tokenA}` } }, res => { res.resume(); resolve(res.statusCode ?? 0); });
      r.end();
    });
    assert.equal(foreignHostStatus, 403);
    assert.equal((await request(`/orgs/${a.orgId}/chat`, tokenA, "POST", { ...input(), price: "0", role: "owner" }, { "Idempotency-Key": id() })).status, 400);
    const keyRes = await request(`/orgs/${a.orgId}/keys`, tokenA, "POST", { provider: "fake", key: "HTTP_SECRET_SENTINEL_12345678", shared: true });
    assert.equal(keyRes.status, 201);
    const key = await keyRes.json() as { id: string };
    assert.ok(!JSON.stringify(key).includes("HTTP_SECRET_SENTINEL"));
    assert.equal((await request(`/orgs/${b.orgId}/keys/${key.id}`, tokenB, "DELETE")).status, 404);
    assert.equal((await request(`/orgs/${a.orgId}/keys/${key.id}/rotate`, tokenMember, "POST", { key: "other-secret-1234567" })).status, 403);
    const keyId = id();
    const good = await request(`/orgs/${a.orgId}/chat`, tokenA, "POST", input(), { "Idempotency-Key": keyId });
    assert.equal(good.status, 200); assert.match(await good.text(), /event: delta/);
    const duplicate = await request(`/orgs/${a.orgId}/chat`, tokenA, "POST", input(), { "Idempotency-Key": keyId });
    assert.equal(duplicate.status, 200); assert.doesNotMatch(await duplicate.text(), /event: delta/);
    assert.equal(fake.calls, 1);
    fake.run = async () => { throw new Error("PROVIDER_ERROR_SECRET_SENTINEL"); };
    const failure = await request(`/orgs/${a.orgId}/chat`, tokenA, "POST", { ...input(), messages: [{ role: "user", content: "PROMPT_SECRET_SENTINEL" }] }, { "Idempotency-Key": id() });
    assert.equal(failure.status, 500);
    assert.doesNotMatch(await failure.text(), /SECRET_SENTINEL/);
    const persisted = JSON.stringify((await pool.query("SELECT * FROM commercial.usage_requests")).rows);
    assert.ok(!persisted.includes("synthetic request"));
    await pool.query("UPDATE commercial.memberships SET role='viewer' WHERE org_id=$1 AND user_id=$2", [a.orgId, a.userId]);
    assert.equal((await request(`/orgs/${a.orgId}/chat`, tokenA, "POST", input(), { "Idempotency-Key": id() })).status, 403);
    assert.ok(!JSON.stringify(logs).includes("HTTP_SECRET_SENTINEL"));
    assert.ok(!JSON.stringify(logs).includes(tokenA));
    assert.doesNotMatch(JSON.stringify(logs), /SECRET_SENTINEL/);
  } finally { await new Promise<void>(r => server.close(() => r())); }
});

test("migration checksum is idempotent and destructive rollback refuses populated databases", async () => {
  await migrate(pool); await migrate(pool);
  await assert.rejects(rollbackEmptyDatabase(pool), /rollback_requires_empty_database/);
  assert.equal((await pool.query("SELECT count(*)::integer n FROM commercial.schema_migrations")).rows[0].n, 1);
  const checksum = (await pool.query("SELECT checksum FROM commercial.schema_migrations WHERE version=1")).rows[0].checksum;
  await pool.query("UPDATE commercial.schema_migrations SET checksum='changed'");
  await assert.rejects(migrate(pool), /migration_checksum_mismatch/);
  await pool.query("UPDATE commercial.schema_migrations SET checksum=$1", [checksum]);
  await pool.query("TRUNCATE commercial.users,commercial.organizations,commercial.budget_accounts RESTART IDENTITY CASCADE");
  await rollbackEmptyDatabase(pool);
  await migrate(pool);
});

test("encryption keyring rotates the master key while preserving explicitly retained versions", async () => {
  const v1 = randomBytes(32), v2 = randomBytes(32);
  const old = new DevKeyringProtector(new Map([["v1", v1]]), "v1");
  const encrypted = await old.seal("SYNTHETIC_KMS_ROTATION", "tenant-and-key-version");
  const rotated = new DevKeyringProtector(new Map([["v1", v1], ["v2", v2]]), "v2");
  assert.equal(await rotated.open(encrypted, "tenant-and-key-version"), "SYNTHETIC_KMS_ROTATION");
  assert.equal((await rotated.seal("new-synthetic-value", "context")).keyVersion, "v2");
  await rejectsCode(rotated.open(encrypted, "wrong-tenant"), "credential_unavailable");
  await rejectsCode(new DevKeyringProtector(new Map([["v2", v2]]), "v2").open(encrypted, "tenant-and-key-version"), "credential_unavailable");
});

test("unsupported models, custom URLs and arbitrary prompt properties are rejected without network", async () => {
  assert.throws(() => parseChat({ ...input(), baseURL: "http://169.254.169.254" }));
  await rejectsCode(service().chat(a, { ...input(), provider: "qwen" }, id(), new AbortController().signal), "provider_not_allowed");
  await rejectsCode(service().chat(a, { ...input(), model: "unbounded" }, id(), new AbortController().signal), "model_not_allowed");
  assert.equal((await pool.query("SELECT count(*)::integer n FROM commercial.usage_attempts")).rows[0].n, 0);
});
