import { pathToFileURL } from "node:url";
import { readConfig, poolFor } from "./core.js";
import { migrate } from "./db.js";
import { TestIdentityVerifier } from "./identity.js";
import { CredentialStore, DevKeyringProtector } from "./secrets.js";
import { createCommercialApp } from "./app.js";

export async function startCommercial(env = process.env) {
  const config = readConfig(env);
  if (!config.enabled) { console.log("Commercial backend disabled; local Excel server unchanged."); return null; }
  const pool = poolFor(config.databaseUrl);
  try {
    await migrate(pool);
    const credentials = new CredentialStore(pool, new DevKeyringProtector(new Map([["dev-v1", config.masterKey]]), "dev-v1"));
    const { app, ledger } = createCommercialApp({ pool, identity: new TestIdentityVerifier(pool), credentials, allowFake: config.allowFake, log: m => console.log(JSON.stringify(m)) });
    await ledger.recoverStale();
    const server = app.listen(config.port, config.host);
    await new Promise<void>((resolve, reject) => { server.once("listening", resolve); server.once("error", reject); });
    server.requestTimeout = 20_000;
    server.headersTimeout = 10_000;
    const recovery = setInterval(() => { void ledger.recoverStale().catch(() => console.error("commercial_recovery_failed")); }, 60_000);
    recovery.unref();
    let closed = false;
    const close = async () => { if (closed) return; closed = true; clearInterval(recovery); await new Promise<void>(resolve => server.close(() => resolve())); await pool.end(); };
    process.once("SIGTERM", () => { void close(); });
    process.once("SIGINT", () => { void close(); });
    console.log(`Commercial C1 test server: http://${config.host}:${config.port}; no real providers; no automatic credit grants.`);
    return { server, pool, close };
  } catch (e) { await pool.end(); throw e; }
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  void startCommercial().catch(() => { console.error("Commercial server failed; verify config, migrations and database availability."); process.exitCode = 1; });
}
