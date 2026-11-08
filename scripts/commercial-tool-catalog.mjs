// Run from repository root: node --import tsx scripts/commercial-tool-catalog.mjs
// Generation only: include all host-supported Excel tools, exclude local-service and web tools.
import { writeFileSync } from "node:fs";
globalThis.Office = { context: { requirements: { isSetSupported: () => true } } };
const { toolsForApi } = await import("../src/excel/toolSchemas.ts");
const catalog = [...toolsForApi(false, false, false), ...toolsForApi(false, false, false, new Set()).filter(t => t.function.name === "enable_tools")];
writeFileSync(new URL("../server/src/commercial/toolCatalog.json", import.meta.url), JSON.stringify(catalog, null, 2) + "\n");
console.log(`Generated commercial catalog: ${catalog.length} existing tools.`);
