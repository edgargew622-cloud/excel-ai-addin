import { createCipheriv, createDecipheriv, randomBytes } from "node:crypto";
import type { Pool, PoolClient } from "pg";
import { validateKey, keyHint } from "../keyStore.js";
import { fail, id, transaction, type Context } from "./core.js";
import { context } from "./identity.js";
import { audit } from "./db.js";

export interface Envelope { keyVersion: string; iv: string; tag: string; data: string }
export interface SecretProtector {
  seal(plain: string, associatedContext: string): Promise<Envelope>;
  open(envelope: Envelope, associatedContext: string): Promise<string>;
}
/** Dev/test implementation. A managed KMS implementation must replace this for production. */
export class DevKeyringProtector implements SecretProtector {
  constructor(private readonly keys: ReadonlyMap<string, Buffer>, private readonly current: string) {
    if (!keys.get(current) || [...keys.values()].some(k => k.length !== 32)) throw new Error("invalid_dev_keyring");
  }
  async seal(plain: string, associatedContext: string): Promise<Envelope> {
    const iv = randomBytes(12);
    const cipher = createCipheriv("aes-256-gcm", this.keys.get(this.current)!, iv);
    cipher.setAAD(Buffer.from(associatedContext));
    const data = Buffer.concat([cipher.update(plain, "utf8"), cipher.final()]);
    return { keyVersion: this.current, iv: iv.toString("base64"), tag: cipher.getAuthTag().toString("base64"), data: data.toString("base64") };
  }
  async open(e: Envelope, associatedContext: string): Promise<string> {
    try {
      const key = this.keys.get(e.keyVersion);
      if (!key) throw new Error();
      const decipher = createDecipheriv("aes-256-gcm", key, Buffer.from(e.iv, "base64"));
      decipher.setAAD(Buffer.from(associatedContext));
      decipher.setAuthTag(Buffer.from(e.tag, "base64"));
      return Buffer.concat([decipher.update(Buffer.from(e.data, "base64")), decipher.final()]).toString("utf8");
    } catch { return fail(503, "credential_unavailable"); }
  }
}

interface CredentialRow { id: string; provider: string; owner_user_id: string | null; version: number; ciphertext: Envelope; hint: string; revoked: boolean }
const metadata = (r: CredentialRow) => ({ id: r.id, provider: r.provider, ownerUserId: r.owner_user_id, version: r.version, hint: r.hint, revoked: r.revoked });
const aad = (org: string, row: Pick<CredentialRow, "id" | "provider" | "owner_user_id" | "version">) => `${org}:${row.id}:${row.provider}:${row.owner_user_id ?? "organization"}:${row.version}`;

export class CredentialStore {
  constructor(private readonly pool: Pool, private readonly protector: SecretProtector) {}

  async create(principal: Context, provider: string, raw: unknown, shared = false) {
    const key = validateKey(raw);
    return transaction(this.pool, async c => {
      const ctx = await context(c, principal.orgId, principal.userId, true);
      if (ctx.role === "viewer" || shared && ctx.role !== "admin" && ctx.role !== "owner") return fail(403, "credential_permission_denied");
      const row = { id: id(), provider, owner_user_id: shared ? null : ctx.userId, version: 1, hint: keyHint(key), revoked: false };
      const encrypted = await this.protector.seal(key, aad(ctx.orgId, row));
      await c.query("INSERT INTO commercial.credentials(org_id,id,provider,owner_user_id,version,ciphertext,hint) VALUES($1,$2,$3,$4,$5,$6,$7)", [ctx.orgId, row.id, provider, row.owner_user_id, row.version, encrypted, row.hint]);
      await audit(c, ctx.orgId, ctx.userId, "credential_created", row.id);
      return metadata(row as CredentialRow);
    });
  }
  async list(principal: Context) {
    const ctx = await context(this.pool, principal.orgId, principal.userId);
    const rows = await this.pool.query(`SELECT id,provider,owner_user_id,version,hint,revoked FROM commercial.credentials
      WHERE org_id=$1 AND (owner_user_id IS NULL OR owner_user_id=$2) ORDER BY id LIMIT 100`, [ctx.orgId, ctx.userId]);
    return rows.rows.map(metadata);
  }
  async rotate(principal: Context, credentialId: string, raw: unknown) {
    const key = validateKey(raw);
    return transaction(this.pool, async c => {
      const ctx = await context(c, principal.orgId, principal.userId, true);
      const row = await this.owned(c, ctx, credentialId, true);
      row.version++;
      const encrypted = await this.protector.seal(key, aad(ctx.orgId, row));
      await c.query("UPDATE commercial.credentials SET version=$3,ciphertext=$4,hint=$5 WHERE org_id=$1 AND id=$2", [ctx.orgId, credentialId, row.version, encrypted, keyHint(key)]);
      await audit(c, ctx.orgId, ctx.userId, "credential_rotated", credentialId);
      return metadata({ ...row, hint: keyHint(key) });
    });
  }
  async revoke(principal: Context, credentialId: string): Promise<void> {
    await transaction(this.pool, async c => {
      const ctx = await context(c, principal.orgId, principal.userId, true);
      await this.owned(c, ctx, credentialId, true);
      await c.query("UPDATE commercial.credentials SET revoked=true,ciphertext=NULL WHERE org_id=$1 AND id=$2", [ctx.orgId, credentialId]);
      await audit(c, ctx.orgId, ctx.userId, "credential_revoked", credentialId);
    });
  }
  /** Never expose this internal method as an HTTP key-read endpoint. */
  async resolve(principal: Context, credentialId: string, provider: string): Promise<{ value: string; version: number }> {
    return transaction(this.pool, async c => {
      const ctx = await context(c, principal.orgId, principal.userId, true);
      if (ctx.role === "viewer") return fail(403, "credential_permission_denied");
      const row = await this.owned(c, ctx, credentialId, false);
      if (row.provider !== provider || row.revoked || !row.ciphertext) return fail(404, "credential_not_found");
      return { value: await this.protector.open(row.ciphertext, aad(ctx.orgId, row)), version: row.version };
    });
  }
  private async owned(c: PoolClient, ctx: Context, credentialId: string, manage: boolean): Promise<CredentialRow> {
    const found = await c.query(`SELECT * FROM commercial.credentials WHERE org_id=$1 AND id=$2
      AND (owner_user_id IS NULL OR owner_user_id=$3) FOR UPDATE`, [ctx.orgId, credentialId, ctx.userId]);
    if (!found.rowCount) return fail(404, "credential_not_found");
    const r = found.rows[0];
    if (manage && (ctx.role === "viewer" || r.owner_user_id === null && ctx.role !== "admin" && ctx.role !== "owner")) return fail(403, "credential_permission_denied");
    if (r.revoked) return fail(404, "credential_not_found");
    return r;
  }
}
