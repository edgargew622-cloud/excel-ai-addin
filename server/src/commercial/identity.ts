import type { Pool, PoolClient } from "pg";
import { fail, sha256, type Context } from "./core.js";

export interface IdentityVerifier { verify(bearer: string): Promise<string> }
/** C1 fixture credentials, not a production identity provider. Enabled only by dev config. */
export class TestIdentityVerifier implements IdentityVerifier {
  constructor(private readonly pool: Pool) {}
  async verify(bearer: string): Promise<string> {
    if (!/^[A-Za-z0-9_-]{24,128}$/.test(bearer)) return fail(401, "unauthorized");
    const r = await this.pool.query(`SELECT t.user_id FROM commercial.test_auth_tokens t JOIN commercial.users u ON u.id=t.user_id
      WHERE t.token_hash=$1 AND t.expires_at>now() AND u.active`, [sha256(bearer)]);
    return r.rows[0]?.user_id ?? fail(401, "unauthorized");
  }
}

/** Membership is always checked in DB; request orgId/userId never grant rights. */
export async function context(c: Pool | PoolClient, orgId: string, userId: string, write = false): Promise<Context> {
  const result = await c.query(`SELECT m.role FROM commercial.memberships m
    JOIN commercial.users u ON u.id=m.user_id JOIN commercial.organizations o ON o.id=m.org_id
    WHERE m.org_id=$1 AND m.user_id=$2 AND m.active AND u.active AND o.active
    ${write ? "FOR SHARE OF m,u,o" : ""}`, [orgId, userId]);
  if (!result.rowCount) return fail(404, "organization_not_found");
  return { orgId, userId, role: result.rows[0].role };
}

export async function entitlement(c: PoolClient, ctx: Context, mode: "hosted" | "byok"): Promise<void> {
  if (ctx.role === "viewer") return fail(403, "role_cannot_spend");
  const r = await c.query("SELECT * FROM commercial.subscriptions WHERE org_id=$1 AND status='active' AND valid_until>now() FOR SHARE", [ctx.orgId]);
  if (!r.rowCount || !(mode === "hosted" ? r.rows[0].allow_hosted : r.rows[0].allow_byok)) return fail(403, "entitlement_required");
}
