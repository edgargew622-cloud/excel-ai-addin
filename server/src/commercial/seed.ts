import { randomBytes } from "node:crypto";
import { existsSync, writeFileSync } from "node:fs";
import { pathToFileURL } from "node:url";
import type { Pool } from "pg";
import { id, readConfig, poolFor, sha256, transaction } from "./core.js";
import { migrate } from "./db.js";

/** Explicit synthetic fixture only: no real providers or automatic trial credits. */
export async function seedDevelopment(pool: Pool) {
  const orgId = id(), userId = id(), token = randomBytes(32).toString("base64url");
  await transaction(pool, async c => {
    await c.query("SELECT pg_advisory_xact_lock(703124801)");
    await c.query("SELECT pg_advisory_xact_lock(703124802)");
    await c.query("INSERT INTO commercial.users(id,external_subject) VALUES($1,$2)", [userId, `c1-fixture:${userId}`]);
    await c.query("INSERT INTO commercial.organizations(id,name) VALUES($1,'C1 synthetic organization')", [orgId]);
    await c.query("INSERT INTO commercial.memberships(org_id,user_id,role) VALUES($1,$2,'owner')", [orgId, userId]);
    await c.query("INSERT INTO commercial.test_auth_tokens(token_hash,user_id,expires_at) VALUES($1,$2,now()+interval '1 day')", [sha256(token), userId]);
    await c.query("INSERT INTO commercial.subscriptions(org_id,plan,status,valid_until,allow_hosted,allow_byok) VALUES($1,'Free','active',now()+interval '1 day',true,true)", [orgId]);
    for (const meter of ["usd_micro", "service_request"]) {
      // A single shared global period: seeding another organization must not reset it.
      const exists = await c.query("SELECT 1 FROM commercial.budget_accounts WHERE scope='global' AND meter=$1 AND period_start<=now() AND period_end>now() FOR UPDATE", [meter]);
      if (!exists.rowCount) await c.query(`INSERT INTO commercial.budget_accounts(id,scope,meter,period_start,period_end,cap,max_inflight)
        VALUES($1,'global',$2,date_trunc('day',now()),date_trunc('day',now())+interval '1 day',100,4)`, [id(), meter]);
      for (const scope of ["org", "user"]) await c.query(`INSERT INTO commercial.budget_accounts(id,scope,org_id,user_id,meter,period_start,period_end,cap)
        VALUES($1,$2,$3,$4,$5,date_trunc('day',now()),date_trunc('day',now())+interval '1 day',30)`, [id(), scope, orgId, scope === "user" ? userId : null, meter]);
    }
  });
  return { orgId, userId, token, model: "c1-test", provider: "fake", warning: "Synthetic non-production credentials; expire in one day." };
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  void (async () => {
    const config = readConfig(process.env);
    if (!config.enabled || process.env.COMMERCIAL_DEV_SEED !== "1" || !config.allowFake) throw new Error("explicit_dev_seed_required");
    if (existsSync(".commercial-dev.json")) throw new Error("development_fixture_already_exists");
    const pool = poolFor(config.databaseUrl);
    try {
      await migrate(pool);
      const fixture = await seedDevelopment(pool);
      // Tokens are written only into an ignored local file, never stdout or logs.
      writeFileSync(".commercial-dev.json", JSON.stringify(fixture, null, 2), { mode: 0o600, flag: "wx" });
      console.log("Synthetic fixture saved to .commercial-dev.json; keep it out of Git.");
    } finally { await pool.end(); }
  })().catch(() => { console.error("Development seed failed; no credentials printed."); process.exitCode = 1; });
}
