import type { Pool, PoolClient } from "pg";
import { fail, id, transaction, type Context } from "./core.js";
import { context, entitlement } from "./identity.js";

export interface Reservation {
  id: string; state: "reserved" | "dispatched" | "settled" | "released" | "uncertain";
  bound: string; provider_bound: string; pricing_version: string; key_mode: "hosted" | "byok";
  actual_provider_cost: string | null; usage: Record<string, number> | null;
  task_id?: string | null;
}
export interface ReserveCommand {
  idempotencyKey: string; bodyHash: string; provider: string; model: string; mode: "hosted" | "byok";
  credentialId?: string; providerBound: bigint; pricingVersion: string;
  taskId?: string;
}

const publicColumns = "id,state,bound::text,provider_bound::text,pricing_version,key_mode,actual_provider_cost::text,usage,task_id";
const safeUsage = (usage: Record<string, unknown> | undefined): Record<string, number> | null => {
  if (!usage) return null;
  const out: Record<string, number> = {};
  for (const k of ["promptTokens", "completionTokens", "totalTokens", "cachedTokens", "reasoningTokens", "cacheWriteTokens"]) {
    const v = usage[k];
    if (v !== undefined) {
      if (typeof v !== "number" || !Number.isSafeInteger(v) || v < 0) return fail(502, "invalid_provider_usage");
      out[k] = v;
    }
  }
  return out;
};

export class Ledger {
  constructor(readonly pool: Pool) {}

  async reserve(principal: Context, cmd: ReserveCommand): Promise<{ request: Reservation; created: boolean }> {
    if (cmd.providerBound <= 0n || cmd.providerBound > 999_999_999_999_999n) return fail(503, "invalid_quote");
    return transaction(this.pool, async c => {
      const ctx = await context(c, principal.orgId, principal.userId, true);
      await entitlement(c, ctx, cmd.mode);
      await c.query("SELECT pg_advisory_xact_lock(hashtextextended($1,0))", [`${ctx.orgId}:chat:${cmd.idempotencyKey}`]);
      const old = await c.query(`SELECT ${publicColumns},body_hash,user_id FROM commercial.usage_requests WHERE org_id=$1 AND idempotency_key=$2`, [ctx.orgId, cmd.idempotencyKey]);
      if (old.rowCount) {
        if (old.rows[0].body_hash !== cmd.bodyHash || old.rows[0].user_id !== ctx.userId) return fail(409, "idempotency_conflict");
        const { body_hash, user_id, ...request } = old.rows[0];
        return { request, created: false };
      }
      if (cmd.mode === "byok") {
        const key = await c.query(`SELECT id FROM commercial.credentials WHERE org_id=$1 AND id=$2 AND provider=$3
          AND NOT revoked AND (owner_user_id IS NULL OR owner_user_id=$4) FOR SHARE`, [ctx.orgId, cmd.credentialId, cmd.provider, ctx.userId]);
        if (!key.rowCount) return fail(404, "credential_not_found");
      }
      const meter = cmd.mode === "byok" ? "service_request" : "usd_micro";
      const amount = cmd.mode === "byok" ? 1n : cmd.providerBound;
      if (cmd.taskId) {
        const task = await this.lockTask(c, ctx, cmd.taskId);
        if (task.state !== "active") return fail(409, "task_stopped");
        if (task.expired) return fail(402, "task_time_limit");
        if (task.key_mode !== cmd.mode) return fail(400, "task_mode_mismatch");
        if (task.steps >= task.max_steps) return fail(402, "task_step_limit");
        if (BigInt(task.spent) + BigInt(task.reserved) + amount > BigInt(task.cap)) return fail(402, "task_budget_exhausted");
        // A task is sequential: a later step cannot overtake an unfinished model/tool response.
        const pending = await c.query("SELECT 1 FROM commercial.usage_requests WHERE org_id=$1 AND task_id=$2 AND state IN ('reserved','dispatched','uncertain') LIMIT 1", [ctx.orgId, cmd.taskId]);
        if (pending.rowCount) return fail(409, "task_request_pending");
        await c.query("UPDATE commercial.tasks SET reserved=reserved+$3,steps=steps+1 WHERE org_id=$1 AND id=$2", [ctx.orgId, cmd.taskId, amount.toString()]);
      }
      const accounts = await c.query(`SELECT * FROM commercial.budget_accounts WHERE meter=$3 AND period_start<=now() AND period_end>now()
        AND (scope='global' OR (org_id=$1 AND (scope='org' OR (scope='user' AND user_id=$2)))) ORDER BY id FOR UPDATE`, [ctx.orgId, ctx.userId, meter]);
      if (accounts.rowCount !== 3 || new Set(accounts.rows.map(a => a.scope)).size !== 3) return fail(403, "budget_not_configured");
      for (const a of accounts.rows) {
        if (a.frozen) return fail(503, "budget_frozen");
        if (BigInt(a.spent) + BigInt(a.reserved) + amount > BigInt(a.cap)) return fail(402, "budget_exhausted");
        const active = await c.query(`SELECT count(*)::integer AS n FROM commercial.allocations x
          JOIN commercial.usage_requests r ON r.org_id=x.org_id AND r.id=x.request_id
          WHERE x.account_id=$1 AND r.state IN ('reserved','dispatched','uncertain')`, [a.id]);
        if (active.rows[0].n >= a.max_inflight) return fail(429, "concurrency_limit");
      }
      const requestId = id();
      const row = await c.query(`INSERT INTO commercial.usage_requests
        (org_id,id,user_id,idempotency_key,body_hash,state,provider,model,key_mode,credential_id,pricing_version,bound,provider_bound,task_id)
        VALUES($1,$2,$3,$4,$5,'reserved',$6,$7,$8,$9,$10,$11,$12,$13) RETURNING ${publicColumns}`,
      [ctx.orgId, requestId, ctx.userId, cmd.idempotencyKey, cmd.bodyHash, cmd.provider, cmd.model, cmd.mode, cmd.credentialId ?? null, cmd.pricingVersion, amount.toString(), cmd.providerBound.toString(), cmd.taskId ?? null]);
      for (const a of accounts.rows) {
        await c.query("UPDATE commercial.budget_accounts SET reserved=reserved+$2 WHERE id=$1", [a.id, amount.toString()]);
        await c.query("INSERT INTO commercial.allocations(org_id,request_id,account_id,amount) VALUES($1,$2,$3,$4)", [ctx.orgId, requestId, a.id, amount.toString()]);
      }
      return { request: row.rows[0], created: true };
    });
  }

  /** Durable dispatch fence commits before the provider is called; no implicit replay. */
  async dispatch(ctx: Context, requestId: string, credentialVersion?: number): Promise<string> {
    return transaction(this.pool, async c => {
      const live = await context(c, ctx.orgId, ctx.userId, true);
      const r = await this.lockRequest(c, ctx, requestId);
      await entitlement(c, live, r.key_mode);
      if (r.state !== "reserved") return fail(409, "already_dispatched");
      if (r.task_id) {
        const task = await this.lockTask(c, ctx, r.task_id);
        if (task.state !== "active" || task.expired) return fail(409, "task_stopped");
      }
      if (r.key_mode === "byok") {
        const key = await c.query(`SELECT k.version FROM commercial.credentials k JOIN commercial.usage_requests r
          ON r.org_id=k.org_id AND r.credential_id=k.id WHERE r.org_id=$1 AND r.id=$2 AND NOT k.revoked
          AND (k.owner_user_id IS NULL OR k.owner_user_id=$3) FOR SHARE OF k`, [ctx.orgId, requestId, ctx.userId]);
        if (!key.rowCount || key.rows[0].version !== credentialVersion) return fail(409, "credential_changed");
      }
      const attemptId = id();
      await c.query("UPDATE commercial.usage_requests SET state='dispatched',updated_at=now() WHERE org_id=$1 AND id=$2", [ctx.orgId, requestId]);
      await c.query("INSERT INTO commercial.usage_attempts(org_id,request_id,id,state) VALUES($1,$2,$3,'dispatched')", [ctx.orgId, requestId, attemptId]);
      return attemptId;
    });
  }

  /** This trusted path receives provider observations, never client-supplied prices. */
  async settle(ctx: Context, requestId: string, providerCost: bigint, usage?: Record<string, unknown>): Promise<Reservation> {
    if (providerCost < 0n || providerCost > 999_999_999_999_999n) return fail(502, "invalid_provider_cost");
    const safe = safeUsage(usage);
    return transaction(this.pool, async c => {
      const r = await this.lockRequest(c, ctx, requestId);
      if (r.state === "settled") {
        if (r.actual_provider_cost !== providerCost.toString()) return fail(409, "settlement_conflict");
        return r;
      }
      if (r.state !== "dispatched" && r.state !== "uncertain") return fail(409, "request_not_dispatched");
      if (r.task_id) await this.lockTask(c, ctx, r.task_id);
      const allocations = await this.lockAccounts(c, ctx.orgId, requestId);
      if (providerCost > BigInt(r.provider_bound)) {
        for (const a of allocations) await c.query("UPDATE commercial.budget_accounts SET frozen=true WHERE id=$1", [a.id]);
        await c.query("INSERT INTO commercial.usage_ledger(org_id,request_id,kind,amount,provider_cost) VALUES($1,$2,'overrun',0,$3) ON CONFLICT DO NOTHING", [ctx.orgId, requestId, providerCost.toString()]);
        await this.uncertainIn(c, ctx.orgId, requestId);
        return (await c.query(`SELECT ${publicColumns} FROM commercial.usage_requests WHERE org_id=$1 AND id=$2`, [ctx.orgId, requestId])).rows[0];
      }
      const charge = r.key_mode === "byok" ? 1n : providerCost;
      if (r.task_id) await c.query("UPDATE commercial.tasks SET reserved=reserved-$3,spent=spent+$4 WHERE org_id=$1 AND id=$2", [ctx.orgId, r.task_id, r.bound, charge.toString()]);
      for (const a of allocations) await c.query("UPDATE commercial.budget_accounts SET reserved=reserved-$2,spent=spent+$3 WHERE id=$1", [a.id, a.amount, charge.toString()]);
      await c.query("INSERT INTO commercial.usage_ledger(org_id,request_id,kind,amount,provider_cost) VALUES($1,$2,'settled',$3,$4)", [ctx.orgId, requestId, charge.toString(), providerCost.toString()]);
      await c.query("UPDATE commercial.usage_attempts SET state='settled',finished_at=now() WHERE org_id=$1 AND request_id=$2", [ctx.orgId, requestId]);
      const row = await c.query(`UPDATE commercial.usage_requests SET state='settled',actual_provider_cost=$3,usage=$4,updated_at=now() WHERE org_id=$1 AND id=$2 RETURNING ${publicColumns}`, [ctx.orgId, requestId, providerCost.toString(), safe]);
      return row.rows[0];
    });
  }

  async release(ctx: Context, requestId: string): Promise<void> {
    await transaction(this.pool, async c => {
      const r = await this.lockRequest(c, ctx, requestId);
      if (r.state === "released") return;
      if (r.state !== "reserved") return fail(409, "dispatched_budget_cannot_be_released");
      await this.releaseIn(c, ctx.orgId, requestId);
    });
  }
  async uncertain(ctx: Context, requestId: string): Promise<void> {
    await transaction(this.pool, async c => {
      const r = await this.lockRequest(c, ctx, requestId);
      if (r.state === "dispatched") await this.uncertainIn(c, ctx.orgId, requestId);
    });
  }

  /** Restart recovery never releases a possibly billed dispatched request. */
  async recoverStale(ageSeconds = 180): Promise<number> {
    if (!Number.isInteger(ageSeconds) || ageSeconds < 0) return fail(400, "invalid_recovery_age");
    return transaction(this.pool, async c => {
      const stale = await c.query(`SELECT org_id,id,state FROM commercial.usage_requests WHERE state IN ('reserved','dispatched')
        AND updated_at < now()-($1::integer*interval '1 second') ORDER BY org_id,id FOR UPDATE SKIP LOCKED LIMIT 100`, [ageSeconds]);
      for (const r of stale.rows) {
        if (r.state === "reserved") await this.releaseIn(c, r.org_id, r.id);
        else await this.uncertainIn(c, r.org_id, r.id);
      }
      return stale.rowCount ?? 0;
    });
  }
  async get(ctx: Context, requestId: string): Promise<Reservation> {
    const live = await context(this.pool, ctx.orgId, ctx.userId);
    const row = await this.pool.query(`SELECT ${publicColumns} FROM commercial.usage_requests WHERE org_id=$1 AND id=$2
      AND (user_id=$3 OR $4::boolean)`, [ctx.orgId, requestId, ctx.userId, live.role === "admin" || live.role === "owner"]);
    return row.rows[0] ?? fail(404, "request_not_found");
  }
  private async lockRequest(c: PoolClient, ctx: Context, requestId: string): Promise<Reservation> {
    const row = await c.query(`SELECT ${publicColumns} FROM commercial.usage_requests WHERE org_id=$1 AND id=$2 AND user_id=$3 FOR UPDATE`, [ctx.orgId, requestId, ctx.userId]);
    return row.rows[0] ?? fail(404, "request_not_found");
  }
  private async lockAccounts(c: PoolClient, orgId: string, requestId: string): Promise<Array<{ id: string; amount: string }>> {
    return (await c.query(`SELECT a.id,x.amount::text FROM commercial.budget_accounts a JOIN commercial.allocations x ON x.account_id=a.id
      WHERE x.org_id=$1 AND x.request_id=$2 ORDER BY a.id FOR UPDATE OF a`, [orgId, requestId])).rows;
  }
  private async releaseIn(c: PoolClient, orgId: string, requestId: string): Promise<void> {
    const request = (await c.query("SELECT task_id,user_id,bound::text FROM commercial.usage_requests WHERE org_id=$1 AND id=$2", [orgId, requestId])).rows[0];
    if (request.task_id) {
      await this.lockTask(c, { orgId, userId: request.user_id, role: "member" }, request.task_id);
      await c.query("UPDATE commercial.tasks SET reserved=reserved-$3 WHERE org_id=$1 AND id=$2", [orgId, request.task_id, request.bound]);
    }
    for (const a of await this.lockAccounts(c, orgId, requestId)) await c.query("UPDATE commercial.budget_accounts SET reserved=reserved-$2 WHERE id=$1", [a.id, a.amount]);
    await c.query("UPDATE commercial.usage_requests SET state='released',updated_at=now() WHERE org_id=$1 AND id=$2", [orgId, requestId]);
    await c.query("INSERT INTO commercial.usage_ledger(org_id,request_id,kind,amount) VALUES($1,$2,'released',0)", [orgId, requestId]);
  }
  private async uncertainIn(c: PoolClient, orgId: string, requestId: string): Promise<void> {
    await c.query("UPDATE commercial.usage_requests SET state='uncertain',updated_at=now() WHERE org_id=$1 AND id=$2", [orgId, requestId]);
    await c.query("UPDATE commercial.usage_attempts SET state='uncertain' WHERE org_id=$1 AND request_id=$2 AND state='dispatched'", [orgId, requestId]);
  }
  private async lockTask(c: PoolClient, ctx: Context, taskId: string) {
    const r = await c.query("SELECT *,deadline<=now() expired FROM commercial.tasks WHERE org_id=$1 AND id=$2 AND user_id=$3 FOR UPDATE", [ctx.orgId, taskId, ctx.userId]);
    return r.rows[0] ?? fail(404, "task_not_found");
  }
}
