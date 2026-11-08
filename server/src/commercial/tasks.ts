import type { Pool } from "pg";
import { context, entitlement } from "./identity.js";
import { fail, id, transaction, type Context } from "./core.js";

export class Tasks {
  constructor(private pool: Pool) {}
  async create(principal: Context, mode: "hosted" | "byok") {
    if (mode !== "hosted" && mode !== "byok") return fail(400, "invalid_key_mode");
    return transaction(this.pool, async c => {
      const ctx = await context(c, principal.orgId, principal.userId, true);
      await entitlement(c, ctx, mode);
      await c.query("SELECT pg_advisory_xact_lock(hashtextextended($1,0))", [`${ctx.orgId}:${ctx.userId}:tasks`]);
      const pending = await c.query("SELECT count(*)::int n FROM commercial.tasks WHERE org_id=$1 AND user_id=$2 AND state='active' AND deadline>now()", [ctx.orgId, ctx.userId]);
      if (pending.rows[0].n >= 4) return fail(429, "active_task_limit");
      // Fixed synthetic pilot policy, never a body-supplied quota or a production SKU.
      return (await c.query(`INSERT INTO commercial.tasks(org_id,id,user_id,key_mode,max_steps,cap,deadline)
        VALUES($1,$2,$3,$4,12,$5,now()+interval '15 minutes') RETURNING id,state,max_steps,steps,cap::text,spent::text,reserved::text,deadline,key_mode`,
      [ctx.orgId, id(), ctx.userId, mode, mode === "hosted" ? "200" : "12"])).rows[0];
    });
  }
  async get(principal: Context, taskId: string) {
    const ctx = await context(this.pool, principal.orgId, principal.userId);
    const row = (await this.pool.query(`SELECT id,state,max_steps,steps,cap::text,spent::text,reserved::text,deadline,key_mode,
      CASE WHEN deadline<now() THEN 'task_time_limit' WHEN steps>=max_steps THEN 'task_step_limit'
      WHEN spent+reserved>=cap THEN 'task_budget_exhausted' ELSE NULL END AS stop_reason
      FROM commercial.tasks WHERE org_id=$1 AND id=$2 AND user_id=$3`, [ctx.orgId, taskId, ctx.userId])).rows[0];
    if (!row) return fail(404, "task_not_found");
    const requests = await this.pool.query(`SELECT id,state,pricing_version,bound::text,actual_provider_cost::text FROM commercial.usage_requests
      WHERE org_id=$1 AND task_id=$2 AND user_id=$3 ORDER BY created_at,id LIMIT 50`, [ctx.orgId, taskId, ctx.userId]);
    return { ...row, synthetic: true, requests: requests.rows };
  }
  async finish(principal: Context, taskId: string, state: "completed" | "cancelled" | "stopped") {
    if (!["completed", "cancelled", "stopped"].includes(state)) return fail(400, "invalid_task_state");
    return transaction(this.pool, async c => {
      const ctx = await context(c, principal.orgId, principal.userId, true);
      const r = await c.query(`UPDATE commercial.tasks SET state=$4 WHERE org_id=$1 AND id=$2 AND user_id=$3 AND state='active' RETURNING id`, [ctx.orgId, taskId, ctx.userId, state]);
      if (!r.rowCount) {
        const old = await c.query("SELECT id FROM commercial.tasks WHERE org_id=$1 AND id=$2 AND user_id=$3", [ctx.orgId, taskId, ctx.userId]);
        if (!old.rowCount) return fail(404, "task_not_found");
      }
    });
  }
  async budgets(principal: Context) {
    const ctx = await context(this.pool, principal.orgId, principal.userId);
    return (await this.pool.query(`SELECT scope,meter,cap::text,spent::text,reserved::text,(cap-spent-reserved)::text available,frozen,period_end
      FROM commercial.budget_accounts WHERE org_id=$1 AND (scope='org' OR (scope='user' AND user_id=$2))
      AND period_start<=now() AND period_end>now() ORDER BY scope,meter`, [ctx.orgId, ctx.userId])).rows;
  }
}
