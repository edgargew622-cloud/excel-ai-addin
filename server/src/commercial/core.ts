import { createHash, randomUUID } from "node:crypto";
import { Pool, type PoolClient } from "pg";

export type Role = "owner" | "admin" | "member" | "viewer";
export interface Context { orgId: string; userId: string; role: Role }
export class CommercialError extends Error {
  constructor(readonly status: number, readonly code: string) { super(code); }
}
export const fail = (status: number, code: string): never => { throw new CommercialError(status, code); };
export const sha256 = (text: string): string => createHash("sha256").update(text).digest("hex");
export const id = (): string => randomUUID();
export const uuid = (s: unknown): s is string => typeof s === "string" && /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(s);
export function units(value: unknown): bigint {
  if (typeof value !== "string" || !/^(0|[1-9][0-9]{0,14})$/.test(value)) return fail(400, "invalid_amount");
  return BigInt(value);
}
export function poolFor(url: string): Pool {
  validateLocalDatabase(url);
  return new Pool({ connectionString: url, max: 12, connectionTimeoutMillis: 5_000, idleTimeoutMillis: 10_000, query_timeout: 8_000 });
}
/** pg parses query parameters after the authority; reject all overrides before constructing a client. */
export function validateLocalDatabase(url: string): void {
  let parsed: URL;
  try { parsed = new URL(url); } catch { return fail(503, "database_configuration_required"); }
  if (!["postgres:", "postgresql:"].includes(parsed.protocol) || parsed.hash) return fail(503, "database_configuration_required");
  // No connection-string options are needed in the closed local mode. This also rejects
  // encoded/duplicate host, hostaddr, service, ssl and libpq compatibility overrides.
  if (parsed.search) return fail(503, "database_options_not_allowed");
  if (!["localhost", "127.0.0.1", "[::1]"].includes(parsed.hostname)) return fail(503, "c1_local_database_only");
}
export async function transaction<T>(pool: Pool, fn: (client: PoolClient) => Promise<T>): Promise<T> {
  const c = await pool.connect();
  try {
    await c.query("BEGIN");
    await c.query("SET LOCAL statement_timeout = '8s'");
    await c.query("SET LOCAL lock_timeout = '5s'");
    const result = await fn(c);
    await c.query("COMMIT");
    return result;
  } catch (e) { await c.query("ROLLBACK").catch(() => undefined); throw e; }
  finally { c.release(); }
}

export interface Config {
  enabled: boolean;
  databaseUrl: string;
  host: "127.0.0.1";
  port: number;
  testAuth: true;
  allowFake: boolean;
  masterKey: Buffer;
}
/** C1 cannot be accidentally deployed as production or bound to a public address. */
export function readConfig(env: Record<string, string | undefined>): Config {
  const enabled = env.COMMERCIAL_ENABLED === "1";
  if (!enabled) return { enabled: false, databaseUrl: "", host: "127.0.0.1", port: 3300, testAuth: true, allowFake: false, masterKey: Buffer.alloc(0) };
  if (env.NODE_ENV === "production") return fail(503, "production_requires_c2_identity_and_kms");
  if (env.COMMERCIAL_AUTH_MODE !== "test" || env.COMMERCIAL_BIND && env.COMMERCIAL_BIND !== "127.0.0.1") return fail(503, "c1_local_test_only");
  const url = env.COMMERCIAL_DATABASE_URL ?? "";
  validateLocalDatabase(url);
  if (!env.COMMERCIAL_DEV_MASTER_KEY || !/^[a-f0-9]{64}$/i.test(env.COMMERCIAL_DEV_MASTER_KEY)) return fail(503, "dev_master_key_required");
  const port = Number(env.COMMERCIAL_PORT ?? 3300);
  if (!Number.isInteger(port) || port < 1024 || port > 65535) return fail(503, "invalid_port");
  return { enabled, databaseUrl: url, host: "127.0.0.1", port, testAuth: true, allowFake: env.COMMERCIAL_FAKE_PROVIDER === "1", masterKey: Buffer.from(env.COMMERCIAL_DEV_MASTER_KEY, "hex") };
}
