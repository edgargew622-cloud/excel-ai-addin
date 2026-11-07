import { pathToFileURL } from "node:url";
import { readConfig, poolFor } from "./core.js";
import { migrate, rollbackEmptyDatabase } from "./db.js";

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  void (async () => {
    const config = readConfig(process.env);
    if (!config.enabled) throw new Error("commercial_disabled");
    const pool = poolFor(config.databaseUrl);
    try {
      if (process.env.COMMERCIAL_ROLLBACK_EMPTY === "1") await rollbackEmptyDatabase(pool);
      else await migrate(pool);
      console.log("Commercial schema operation completed.");
    } finally { await pool.end(); }
  })().catch(() => { console.error("Commercial schema operation refused or failed; no credentials printed."); process.exitCode = 1; });
}
