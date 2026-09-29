import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { deleteWebKey, OPENROUTER_URL, saveWebKey, webChatRequest, webKey, webKeysState, webProviders, type KeyStorage } from "./webProvider";
import { toolsForApi, TOOL_SPECS } from "../../excel/toolSchemas";
// @ts-expect-error — скрипт сборки на JS, без объявлений типов
import { webManifest, WEB_ADDIN_ID } from "../../../scripts/web-manifest.mjs";

function memoryStorage(): KeyStorage & { data: Map<string, string> } {
  const data = new Map<string, string>();
  return {
    data,
    getItem: (key) => data.get(key) ?? null,
    setItem: (key, value) => { data.set(key, value); },
    removeItem: (key) => { data.delete(key); }
  };
}

test("the OpenRouter key lives in panel storage and is never shown back", () => {
  const store = memoryStorage();
  assert.equal(webKeysState(store).providers[0].ready, false);
  assert.deepEqual(webProviders(store), [], "без ключа моделей нет");

  const state = saveWebKey("openrouter", "  sk-or-v1-abcdef123456  ", store);
  assert.equal(webKey(store), "sk-or-v1-abcdef123456", "пробелы по краям убраны");
  assert.equal(state.providers[0].hint, "…3456");
  assert.equal(JSON.stringify(state).includes("abcdef"), false, "сам ключ в состояние панели не попадает");
  assert.equal(webProviders(store)[0].id, "openrouter");

  assert.throws(() => saveWebKey("deepseek", "ключ", store), /только OpenRouter/);
  assert.throws(() => saveWebKey("openrouter", "два слова", store), /не похоже/);
  deleteWebKey("openrouter", store);
  assert.equal(webKey(store), "");
});

test("no storage — the panel says so instead of pretending to save", () => {
  assert.equal(webKeysState(null).storage.available, false);
  assert.throws(() => saveWebKey("openrouter", "sk-or-v1-x", null), /недоступно/);
});

test("the chat request is the same one the server builds for OpenRouter", () => {
  const { url, init } = webChatRequest({
    model: "anthropic/claude-sonnet-5.5",
    messages: [{ role: "assistant", content: null, tool_calls: [{ id: "c1", name: "list_sheets", arguments: "{}" }] }],
    tools: [{ type: "function", function: { name: "list_sheets" } }]
  }, "sk-or-v1-key", "https://edgargew622-cloud.github.io");
  assert.equal(url, OPENROUTER_URL);
  const headers = init.headers as Record<string, string>;
  assert.equal(headers.Authorization, "Bearer sk-or-v1-key");
  assert.equal(headers["HTTP-Referer"], "https://edgargew622-cloud.github.io");
  const body = JSON.parse(String(init.body));
  assert.equal(body.stream, true);
  assert.equal(body.tool_choice, "auto");
  assert.deepEqual(body.messages[0].tool_calls[0], { id: "c1", type: "function", function: { name: "list_sheets", arguments: "{}" } });

  assert.throws(() => webChatRequest({ model: "openai/gpt-anything", messages: [], tools: [] }, "k", ""), /не разрешена/);
});

test("tools that need the local server are not given to the model in the web panel", () => {
  const local = new Set(TOOL_SPECS.filter((spec) => spec.needsLocal).map((spec) => spec.name));
  assert.deepEqual([...local].sort(), [
    "create_workbook_backup", "get_scenario", "import_file_table", "list_files", "read_file",
    "read_web_page", "remember_preference", "save_scenario", "web_search"
  ]);
  const web = toolsForApi(false, true, false).map((tool) => tool.function.name);
  const windows = toolsForApi(false, true, true).map((tool) => tool.function.name);
  // Работа с книгой — вся: веб-список — это список Windows без инструментов сервера.
  assert.deepEqual(web, windows.filter((name) => !local.has(name as any)));
  for (const name of local) assert.equal(web.includes(name), false, name);
});

test("the web manifest points to the site, has its own id and no token or localhost", () => {
  const source = readFileSync(new URL("../../../manifest.xml", import.meta.url), "utf8");
  const out: string = webManifest(source, "https://example.github.io/excel-ai-addin/");
  assert.match(out, new RegExp(`<Id>${WEB_ADDIN_ID}</Id>`));
  assert.match(out, /<SourceLocation DefaultValue="https:\/\/example\.github\.io\/excel-ai-addin\/taskpane\.html" \/>/);
  assert.match(out, /<AppDomain>https:\/\/openrouter\.ai<\/AppDomain>/);
  assert.equal(out.includes("localhost"), false);
  assert.equal(out.includes("?t="), false);
  assert.equal(out.match(/<Version>[^<]+<\/Version>/)?.[0], source.match(/<Version>[^<]+<\/Version>/)?.[0], "версия та же");
  assert.throws(() => webManifest(source, "http://example.com/"), /https/);
});

test("the Mac install command only creates the add-ins folder and downloads the manifest", async () => {
  // @ts-expect-error — скрипт сборки на JS, без объявлений типов
  const { macInstallCommand, macUninstallCommand, macPage } = await import("../../../scripts/web-manifest.mjs");
  const base = "https://example.github.io/excel-ai-addin/";
  const command: string = macInstallCommand(base);
  assert.equal(command,
    "mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef && " +
    "curl -fsSL https://example.github.io/excel-ai-addin/manifest.xml -o ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/am-ai.xml && " +
    "echo \"Готово. Перезапустите Excel.\"");
  assert.doesNotMatch(command, /\|\s*(ba|z)?sh|sudo|eval/, "ничего не запускает и не просит пароль");
  assert.match(macUninstallCommand(), /^rm -f ~\/Library\/Containers\/com\.microsoft\.Excel\/Data\/Documents\/wef\/am-ai\.xml /);

  const page: string = macPage(base);
  const shown = page.match(/<code class="block" id="install">([^<]*)<\/code>/)?.[1] ?? "";
  const decoded = shown.replace(/&quot;/g, "\"").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
  assert.equal(decoded, command, "на странице — ровно та команда, что копирует кнопка");
});
