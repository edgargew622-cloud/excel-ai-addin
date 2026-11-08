import { randomBytes, createHash } from "node:crypto";
import { createRemoteJWKSet, jwtVerify, type JWTVerifyGetKey } from "jose";
import type { Pool } from "pg";
import express, { type Router } from "express";
import { fail, sha256, transaction, CommercialError } from "./core.js";
import type { SecretProtector, Envelope } from "./secrets.js";
import type { IdentityVerifier } from "./identity.js";
import { boundedJson } from "./deepseekPilot.js";

const random = () => randomBytes(32).toString("base64url");
const opaque = (s: unknown): s is string => typeof s === "string" && /^[A-Za-z0-9_-]{43}$/.test(s);
export interface EntraConfig { tenant: string; clientId: string; clientSecret: string; origin: string }
export function entraConfig(env: NodeJS.ProcessEnv): EntraConfig {
  const tenant = env.COMMERCIAL_ENTRA_TENANT ?? "", clientId = env.COMMERCIAL_ENTRA_CLIENT_ID ?? "";
  if (![tenant, clientId].every(s => /^[a-f0-9-]{36}$/i.test(s)) || !env.COMMERCIAL_ENTRA_CLIENT_SECRET) return fail(503, "entra_configuration_required");
  const origin = env.COMMERCIAL_PANEL_ORIGIN ?? "";
  // C2 is still a closed local pilot: only the local HTTPS development origin is enabled.
  if (origin !== "https://localhost:3100") return fail(503, "c2_local_https_origin_required");
  return { tenant, clientId, clientSecret: env.COMMERCIAL_ENTRA_CLIENT_SECRET, origin };
}

export class SessionIdentity implements IdentityVerifier {
  constructor(private pool: Pool) {}
  async verify(token: string): Promise<string> {
    if (!opaque(token)) return fail(401, "unauthorized");
    const r = await this.pool.query(`SELECT s.user_id FROM commercial.auth_sessions s JOIN commercial.users u ON u.id=s.user_id
      WHERE s.token_hash=$1 AND s.expires_at>now() AND NOT s.revoked AND u.active`, [sha256(token)]);
    return r.rows[0]?.user_id ?? fail(401, "unauthorized");
  }
}

/** BFF: OIDC tokens stay on the server. The dialog emits only a one-use, parent-bound ticket. */
export class EntraAuth {
  readonly issuer: string;
  private keys: JWTVerifyGetKey;
  constructor(readonly pool: Pool, readonly protector: SecretProtector, readonly config: EntraConfig,
    private fetcher: typeof fetch = fetch, keys?: JWTVerifyGetKey) {
    this.issuer = `https://login.microsoftonline.com/${config.tenant}/v2.0`;
    this.keys = keys ?? createRemoteJWKSet(new URL(`https://login.microsoftonline.com/${config.tenant}/discovery/v2.0/keys`), { timeoutDuration: 5000 });
  }
  async start(parentChallenge: unknown) {
    if (typeof parentChallenge !== "string" || !/^[a-f0-9]{64}$/.test(parentChallenge)) return fail(400, "invalid_parent_challenge");
    const state = random(), nonce = random(), verifier = random();
    await transaction(this.pool, async c => {
      await c.query("SELECT pg_advisory_xact_lock(703124803)");
      await c.query("DELETE FROM commercial.auth_flows WHERE expires_at<now()");
      await c.query("DELETE FROM commercial.auth_redemptions WHERE expires_at<now()");
      await c.query("DELETE FROM commercial.auth_sessions WHERE expires_at<now()");
      const count = await c.query("SELECT count(*)::int n FROM commercial.auth_flows");
      if (count.rows[0].n >= 100) return fail(429, "auth_flow_limit");
      await c.query(`INSERT INTO commercial.auth_flows(state_hash,nonce,verifier,parent_challenge,expires_at)
        VALUES($1,$2,$3,$4,now()+interval '5 minutes')`, [sha256(state), nonce, await this.protector.seal(verifier, `oidc:${sha256(state)}`), parentChallenge]);
    });
    return { dialogUrl: `${this.config.origin}/api/v1/auth/dialog?state=${state}` };
  }
  async authorize(state: unknown) {
    if (!opaque(state)) return fail(400, "invalid_auth_state");
    const row = (await this.pool.query("SELECT * FROM commercial.auth_flows WHERE state_hash=$1 AND expires_at>now() AND NOT claimed", [sha256(state)])).rows[0];
    if (!row) return fail(401, "invalid_auth_state");
    const verifier = await this.protector.open(row.verifier as Envelope, `oidc:${sha256(state)}`);
    const url = new URL(`https://login.microsoftonline.com/${this.config.tenant}/oauth2/v2.0/authorize`);
    url.search = new URLSearchParams({ client_id: this.config.clientId, response_type: "code", response_mode: "query",
      redirect_uri: `${this.config.origin}/api/v1/auth/callback`, scope: "openid profile", state, nonce: row.nonce,
      code_challenge: createHash("sha256").update(verifier).digest("base64url"), code_challenge_method: "S256" }).toString();
    return url.href;
  }
  async callback(state: unknown, code: unknown): Promise<string> {
    if (!opaque(state) || typeof code !== "string" || code.length < 1 || code.length > 8192) return fail(401, "invalid_auth_callback");
    // Commit the one-use claim before token exchange. Errors require a new sign-in, never token/code replay.
    const row = await transaction(this.pool, async c => {
      const r = await c.query(`UPDATE commercial.auth_flows SET claimed=true WHERE state_hash=$1 AND expires_at>now() AND NOT claimed RETURNING *`, [sha256(state)]);
      return r.rows[0] ?? fail(401, "invalid_auth_state");
    });
    const verifier = await this.protector.open(row.verifier, `oidc:${sha256(state)}`);
    const response = await this.fetcher(`https://login.microsoftonline.com/${this.config.tenant}/oauth2/v2.0/token`, {
      method: "POST", redirect: "error", signal: AbortSignal.timeout(8000), headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({ grant_type: "authorization_code", client_id: this.config.clientId, client_secret: this.config.clientSecret,
        code, code_verifier: verifier, redirect_uri: `${this.config.origin}/api/v1/auth/callback` })
    });
    if (!response.ok) return fail(401, "identity_exchange_failed");
    const tokens = await boundedJson(response, 32_768);
    const { payload } = await jwtVerify(tokens.id_token, this.keys, {
      issuer: this.issuer, audience: this.config.clientId, algorithms: ["RS256"], requiredClaims: ["exp", "iat", "sub", "oid", "nonce", "tid"], maxTokenAge: "10m", clockTolerance: 5
    }).catch(() => fail(401, "invalid_identity_token"));
    if (payload.nonce !== row.nonce || payload.tid !== this.config.tenant || typeof payload.sub !== "string" || typeof payload.oid !== "string" || !/^[a-f0-9-]{36}$/i.test(payload.oid) ||
      (payload.azp !== undefined && payload.azp !== this.config.clientId) || (Array.isArray(payload.aud) && payload.aud.length > 1 && payload.azp !== this.config.clientId)) return fail(401, "invalid_identity_token");
    // No automatic signup, organization creation or credit issuance: an operator explicitly invites pilot users.
    const subject = `${this.issuer}|${payload.oid}`;
    const user = (await this.pool.query("SELECT id FROM commercial.users WHERE external_subject=$1 AND active", [subject])).rows[0];
    if (!user) return fail(403, "pilot_invitation_required");
    const ticket = random();
    await this.pool.query(`INSERT INTO commercial.auth_redemptions(ticket_hash,user_id,parent_challenge,expires_at)
      VALUES($1,$2,$3,now()+interval '60 seconds')`, [sha256(ticket), user.id, row.parent_challenge]);
    return ticket;
  }
  async redeem(ticket: unknown, proof: unknown) {
    if (!opaque(ticket) || !opaque(proof)) return fail(401, "invalid_auth_ticket");
    return transaction(this.pool, async c => {
      const r = await c.query(`DELETE FROM commercial.auth_redemptions WHERE ticket_hash=$1 AND parent_challenge=$2 AND expires_at>now() RETURNING user_id`, [sha256(ticket), sha256(proof)]);
      if (!r.rowCount) return fail(401, "invalid_auth_ticket");
      const token = random();
      await c.query("INSERT INTO commercial.auth_sessions(token_hash,user_id,expires_at) VALUES($1,$2,now()+interval '1 hour')", [sha256(token), r.rows[0].user_id]);
      return { token, expiresIn: 3600 };
    });
  }
  router(): Router {
    const router = express.Router();
    router.use(express.json({ limit: "2kb" }));
    // Routes deliberately return constant errors, never an IdP response or callback query.
    const route = (fn: (r: express.Request, s: express.Response) => Promise<unknown>) => (r: express.Request, s: express.Response) => {
      void fn(r, s).catch(() => { if (!s.headersSent) s.status(401).json({ error: { code: "authentication_failed" } }); });
    };
    router.post("/start", route(async (r, s) => s.json(await this.start(r.body?.parentChallenge))));
    router.get("/dialog", route(async (r, s) => s.redirect(await this.authorize(r.query.state))));
    router.get("/callback", route(async (r, s) => {
      let message: { type: string; ticket?: string };
      try { message = { type: "amai-auth", ticket: await this.callback(r.query.state, r.query.code) }; }
      catch { message = { type: "amai-auth-failed" }; }
      const nonce = random();
      s.setHeader("Referrer-Policy", "no-referrer");
      s.setHeader("Content-Security-Policy", `default-src 'none'; script-src 'nonce-${nonce}' https://appsforoffice.microsoft.com; connect-src https://appsforoffice.microsoft.com; style-src 'none'; frame-ancestors 'self'`);
      s.type("html").send(`<!doctype html><html><head><meta charset="utf-8"><title>am.AI sign-in</title><script nonce="${nonce}" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script></head><body>Return to am.AI.<script nonce="${nonce}">Office.onReady(function(){Office.context.ui.messageParent(${JSON.stringify(JSON.stringify(message))},{targetOrigin:${JSON.stringify(this.config.origin)}});});</script></body></html>`);
    }));
    router.post("/redeem", route(async (r, s) => s.json(await this.redeem(r.body?.ticket, r.body?.proof))));
    router.post("/logout", route(async (r, s) => {
      const token = /^Bearer ([A-Za-z0-9_-]{43})$/.exec(r.headers.authorization ?? "")?.[1];
      if (!token) throw new CommercialError(401, "unauthorized");
      await this.pool.query("UPDATE commercial.auth_sessions SET revoked=true WHERE token_hash=$1", [sha256(token)]);
      s.sendStatus(204);
    }));
    return router;
  }
}
