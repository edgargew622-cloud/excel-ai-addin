import { pathToFileURL } from "node:url";
import { poolFor, readConfig, uuid } from "./core.js";
import { entraConfig } from "./auth.js";

/** Operator-only allowlist mapping. Never called by an HTTP login request or creates credits. */
export async function provisionPilot(env = process.env) {
  if (env.COMMERCIAL_C2_ENABLED !== "1" || env.COMMERCIAL_PILOT_PROVISION !== "1") throw new Error("explicit_pilot_provision_required");
  const cfg = readConfig({ ...env, COMMERCIAL_ENABLED: "1", COMMERCIAL_AUTH_MODE: "test" });
  const oidc = entraConfig(env);
  if (!uuid(env.COMMERCIAL_PILOT_USER_ID) || !uuid(env.COMMERCIAL_PILOT_OBJECT_ID)) throw new Error("pilot_user_ids_required");
  const pool = poolFor(cfg.databaseUrl);
  try {
    const updated = await pool.query(`UPDATE commercial.users SET external_subject=$2 WHERE id=$1 AND active
      AND external_subject LIKE 'c1-fixture:%' RETURNING id`, [env.COMMERCIAL_PILOT_USER_ID, `https://login.microsoftonline.com/${oidc.tenant}/v2.0|${env.COMMERCIAL_PILOT_OBJECT_ID}`]);
    if (updated.rowCount !== 1) throw new Error("existing_synthetic_user_required");
    console.log("Pilot identity linked to existing synthetic membership. No budgets changed.");
  } finally { await pool.end(); }
}
if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  void provisionPilot().catch(() => { console.error("Pilot provisioning failed; verify explicit local configuration."); process.exitCode = 1; });
}
