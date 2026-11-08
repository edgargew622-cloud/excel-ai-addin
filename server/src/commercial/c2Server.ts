import { pathToFileURL } from "node:url";
import { poolFor, readConfig } from "./core.js";
import { migrate } from "./db.js";
import { DevKeyringProtector, CredentialStore } from "./secrets.js";
import { EntraAuth, SessionIdentity, entraConfig } from "./auth.js";
import { FakeAgentProvider } from "./providers.js";
import { createCommercialApp } from "./app.js";

export async function startC2(env = process.env) {
  if (env.COMMERCIAL_C2_ENABLED !== "1") { console.log("Commercial C2 disabled."); return null; }
  // Preserve C1 local-only/prod guards. This entrypoint never accepts fixture bearer tokens.
  const config = readConfig({ ...env, COMMERCIAL_ENABLED: "1", COMMERCIAL_AUTH_MODE: "test" });
  const oidc = entraConfig(env);
  const pool = poolFor(config.databaseUrl);
  try {
    await migrate(pool);
    const protector = new DevKeyringProtector(new Map([["dev-v1", config.masterKey]]), "dev-v1");
    const auth = new EntraAuth(pool, protector, oidc);
    const { app, ledger } = createCommercialApp({ pool, identity: new SessionIdentity(pool), auth, c2: true, panelOrigin: oidc.origin,
      credentials: new CredentialStore(pool, protector), allowFake: true, adapters: new Map([["fake", new FakeAgentProvider()]]), log: metric => console.log(JSON.stringify(metric)) });
    await ledger.recoverStale();
    const server = app.listen(config.port, config.host);
    await new Promise<void>((resolve, reject) => { server.once("listening", resolve); server.once("error", reject); });
    server.requestTimeout = 20_000; server.headersTimeout = 10_000;
    const timer = setInterval(() => { void ledger.recoverStale().catch(() => console.error("commercial_recovery_failed")); }, 60_000); timer.unref();
    let closed = false;
    const close = async () => { if (closed) return; closed = true; clearInterval(timer); await new Promise<void>(r => server.close(() => r())); await pool.end(); };
    process.once("SIGTERM", () => { void close(); }); process.once("SIGINT", () => { void close(); });
    console.log("Commercial C2 local pilot started; Entra identity required; synthetic provider only.");
    return { server, pool, close };
  } catch (e) { await pool.end(); throw e; }
}
if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  void startC2().catch(() => { console.error("C2 startup failed; check local config and identity registration."); process.exitCode = 1; });
}
