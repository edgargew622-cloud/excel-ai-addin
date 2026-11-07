import express, { type Request, type Response, type NextFunction } from "express";
import type { Pool } from "pg";
import { KeyError } from "../keyStore.js";
import { PROVIDERS } from "../providers.js";
import { isLoopbackAddress } from "../localOnly.js";
import { CommercialError, fail, id, uuid, type Context } from "./core.js";
import { context, type IdentityVerifier } from "./identity.js";
import { Ledger } from "./ledger.js";
import type { CredentialStore } from "./secrets.js";
import { ChatService } from "./service.js";
import { FakeProvider, type ProviderAdapter } from "./providers.js";

interface Req extends Request { userId?: string; ctx?: Context; requestId?: string }
const wrap = (fn: (r: Req, s: Response, n: NextFunction) => Promise<unknown>) => (r: Req, s: Response, n: NextFunction) => { void fn(r, s, n).catch(n); };
const object = (body: unknown, keys: string[]): Record<string, unknown> => {
  if (!body || typeof body !== "object" || Array.isArray(body) || Object.keys(body).some(k => !keys.includes(k))) return fail(400, "invalid_body");
  return body as Record<string, unknown>;
};

export interface AuditMetric { requestId: string; status: number; durationMs: number; code?: string }
export function createCommercialApp(options: {
  pool: Pool; identity: IdentityVerifier; credentials: CredentialStore; allowFake: boolean;
  adapters?: ReadonlyMap<string, ProviderAdapter>; log?: (metric: AuditMetric) => void;
}) {
  const app = express();
  app.disable("x-powered-by");
  app.set("query parser", "simple");
  const ledger = new Ledger(options.pool);
  const adapters = options.adapters ?? new Map(options.allowFake ? [["fake", new FakeProvider()]] : []);
  const service = new ChatService(ledger, options.credentials, adapters);
  app.use((req: Req, res, next) => {
    req.requestId = id();
    const started = Date.now();
    res.setHeader("X-Request-Id", req.requestId);
    res.setHeader("Cache-Control", "no-store");
    res.setHeader("X-Content-Type-Options", "nosniff");
    res.once("finish", () => options.log?.({ requestId: req.requestId!, status: res.statusCode, durationMs: Date.now() - started }));
    const port = req.socket.localPort;
    if (!isLoopbackAddress(req.socket.remoteAddress) || ![`127.0.0.1:${port}`, `localhost:${port}`, `[::1]:${port}`].includes(req.headers.host ?? "")) return next(new CommercialError(403, "local_test_access_only"));
    if (req.headers.origin && ![`http://127.0.0.1:${port}`, `http://localhost:${port}`].includes(req.headers.origin)) return next(new CommercialError(403, "foreign_origin"));
    next();
  });
  app.get("/api/v1/health", (_r, s) => s.json({ ok: true, mode: "commercial-c1-local-test", paidProviders: false }));
  app.get("/api/v1/ready", wrap(async (_req, res) => {
    try {
      const r = await options.pool.query("SELECT version FROM commercial.schema_migrations WHERE version=1");
      res.status(r.rowCount === 1 ? 200 : 503).json({ ready: r.rowCount === 1 });
    }
    catch { res.status(503).json({ ready: false }); }
  }));
  const rates = new Map<string, { at: number; count: number }>();
  app.use("/api/v1", wrap(async (req, _res, next) => {
    const match = /^Bearer ([A-Za-z0-9_-]{24,128})$/.exec(req.headers.authorization ?? "");
    if (!match) return fail(401, "unauthorized");
    req.userId = await options.identity.verify(match[1]);
    const now = Date.now(), old = rates.get(req.userId);
    if (old && now - old.at < 60_000) { if (++old.count > 200) return fail(429, "request_rate_limit"); }
    else { if (rates.size >= 1000) rates.delete(rates.keys().next().value!); rates.set(req.userId, { at: now, count: 1 }); }
    next();
  }));
  app.use(express.json({ limit: "32kb", strict: true }));
  app.get("/api/v1/me", wrap(async (req, res) => {
    const r = await options.pool.query(`SELECT o.id,o.name,m.role FROM commercial.memberships m JOIN commercial.organizations o ON o.id=m.org_id
      JOIN commercial.users u ON u.id=m.user_id WHERE m.user_id=$1 AND m.active AND o.active AND u.active ORDER BY o.id LIMIT 100`, [req.userId]);
    res.json({ userId: req.userId, organizations: r.rows });
  }));
  const router = express.Router({ mergeParams: true });
  app.use("/api/v1/orgs/:orgId", router);
  router.use(wrap(async (req, _res, next) => {
    if (!uuid(req.params.orgId)) return fail(400, "invalid_organization");
    req.ctx = await context(options.pool, req.params.orgId, req.userId!);
    next();
  }));
  router.get("/entitlements", wrap(async (req, res) => {
    const r = await options.pool.query("SELECT plan,status,valid_until,allow_hosted,allow_byok,version FROM commercial.subscriptions WHERE org_id=$1", [req.ctx!.orgId]);
    res.json({ subscription: r.rows[0] ?? null });
  }));
  router.get("/usage", wrap(async (req, res) => {
    const ctx = req.ctx!;
    const rows = await options.pool.query(`SELECT id,state,provider,model,key_mode,pricing_version,bound::text,actual_provider_cost::text,usage,created_at
      FROM commercial.usage_requests WHERE org_id=$1 AND (user_id=$2 OR $3::boolean) ORDER BY created_at DESC,id LIMIT 100`, [ctx.orgId, ctx.userId, ctx.role === "owner" || ctx.role === "admin"]);
    res.json({ requests: rows.rows });
  }));
  router.get("/usage/:id", wrap(async (req, res) => {
    if (!uuid(req.params.id)) return fail(400, "invalid_request_id");
    res.json(await ledger.get(req.ctx!, req.params.id));
  }));
  router.get("/keys", wrap(async (req, res) => { res.json({ keys: await options.credentials.list(req.ctx!) }); }));
  router.post("/keys", wrap(async (req, res) => {
    const b = object(req.body, ["provider", "key", "shared"]);
    if (typeof b.provider !== "string" || ![...PROVIDERS.map(p => p.id), ...(options.allowFake ? ["fake"] : [])].includes(b.provider) || b.shared !== undefined && typeof b.shared !== "boolean") return fail(400, "invalid_credential_provider");
    res.status(201).json(await options.credentials.create(req.ctx!, b.provider, b.key, b.shared === true));
  }));
  router.post("/keys/:id/rotate", wrap(async (req, res) => {
    const b = object(req.body, ["key"]);
    if (!uuid(req.params.id)) return fail(400, "invalid_credential_id");
    res.json(await options.credentials.rotate(req.ctx!, req.params.id, b.key));
  }));
  router.delete("/keys/:id", wrap(async (req, res) => {
    if (!uuid(req.params.id)) return fail(400, "invalid_credential_id");
    await options.credentials.revoke(req.ctx!, req.params.id); res.sendStatus(204);
  }));
  router.post("/chat", wrap(async (req, res) => {
    const abort = new AbortController();
    const onAbort = () => abort.abort();
    const onClose = () => { if (!res.writableEnded) abort.abort(); };
    req.once("aborted", onAbort); res.once("close", onClose);
    try {
      const key = req.headers["idempotency-key"];
      if (typeof key !== "string") return fail(400, "idempotency_key_required");
      const result = await service.chat(req.ctx!, req.body, key, abort.signal);
      if (res.destroyed) return;
      res.type("text/event-stream");
      res.write(`event: request\ndata: ${JSON.stringify({ ...result.request, replay: result.replay })}\n\n`);
      if (result.text !== null) res.write(`event: delta\ndata: ${JSON.stringify({ text: result.text })}\n\n`);
      res.end("event: done\ndata: {}\n\n");
    } finally { req.off("aborted", onAbort); res.off("close", onClose); }
  }));
  app.use((_r, _s, next) => next(new CommercialError(404, "not_found")));
  app.use((error: unknown, req: Req, res: Response, _next: NextFunction) => {
    // Deliberately never log Error objects, SQL, Authorization or request bodies.
    const known = error instanceof CommercialError;
    const parseError = (error as { type?: string })?.type;
    const status = known ? error.status : error instanceof KeyError ? 400 : parseError === "entity.too.large" ? 413 : parseError === "entity.parse.failed" ? 400 : 500;
    const code = known ? error.code : error instanceof KeyError ? "invalid_credential" : status === 413 ? "body_limit" : status === 400 ? "invalid_json" : "internal_error";
    options.log?.({ requestId: req.requestId!, status, durationMs: 0, code });
    if (!res.headersSent && !res.destroyed) res.status(status).json({ error: { code }, requestId: req.requestId });
    else if (!res.destroyed) res.end();
  });
  return { app, service, ledger };
}
