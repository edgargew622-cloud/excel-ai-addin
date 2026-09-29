/**
 * Манифест веб-панели (Mac, Excel в браузере) из основного manifest.xml.
 *
 * Панель лежит на GitHub Pages, поэтому все адреса https://localhost:3000/
 * заменяются адресом сайта, а Id — свой: на одном компьютере веб-панель и
 * панель Windows не должны считаться одной надстройкой. Версия берётся из
 * основного манифеста — разойтись они не могут. Токена в этом манифесте нет:
 * локального сервера, который он защищает, у веб-панели нет.
 *
 *   node scripts/web-manifest.mjs [--base https://…/] [--out dist-web/manifest.xml]
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

export const WEB_ADDIN_ID = "5409e1d3-ee5f-49a0-be50-a9e930bce2af";
export const DEFAULT_BASE = "https://edgargew622-cloud.github.io/excel-ai-addin/";

export function webManifest(source, base = DEFAULT_BASE) {
  if (!/^https:\/\/[^\s"<>]+\/$/.test(base)) throw new Error(`Адрес сайта должен быть https и кончаться на /: ${base}`);
  const local = "https://localhost:3000/";
  if (!source.includes(local)) throw new Error("В manifest.xml нет адресов https://localhost:3000/ — нечего заменять.");
  let out = source.split(local).join(base);
  out = out.replace(/<Id>[^<]+<\/Id>/, `<Id>${WEB_ADDIN_ID}</Id>`);
  out = out.replace(/<DisplayName DefaultValue="[^"]*" \/>/, '<DisplayName DefaultValue="am.AI (Mac и Excel в браузере)" />');
  out = out.replace(
    /<Description DefaultValue="[^"]*" \/>/,
    '<Description DefaultValue="AI-агент для Excel без установки программы: модели через OpenRouter по вашему ключу." />'
  );
  // Панель из интернета ходит к OpenRouter сама — домен объявляется явно.
  out = out.replace("<Hosts>", `<AppDomains>\n    <AppDomain>https://openrouter.ai</AppDomain>\n  </AppDomains>\n\n  <Hosts>`);
  if (out.includes("localhost")) throw new Error("В веб-манифесте остался localhost.");
  return out;
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
  const option = (name, fallback) => {
    const i = process.argv.indexOf(`--${name}`);
    return i > 0 ? process.argv[i + 1] : fallback;
  };
  const base = option("base", process.env.WEB_PANEL_BASE || DEFAULT_BASE);
  const out = resolve(root, option("out", "dist-web/manifest.xml"));
  writeFileSync(out, webManifest(readFileSync(resolve(root, "manifest.xml"), "utf8"), base));
  console.log(`Веб-манифест: ${out} (${base})`);
}
