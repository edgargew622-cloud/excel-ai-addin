import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { DEEPSEEK_URL, deleteWebKey, OPENROUTER_URL, saveWebKey, webChatRequest, webKey, webKeysState, webProviders, WEB_PROVIDERS, type KeyStorage } from "./webProvider";
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

test("each provider key lives in panel storage and is never shown back", () => {
  const store = memoryStorage();
  assert.deepEqual(webKeysState(store).providers.map((p) => [p.id, p.ready]), [["deepseek", false], ["openrouter", false]]);
  assert.deepEqual(webProviders(store), [], "без ключа моделей нет");

  const state = saveWebKey("openrouter", "  sk-or-v1-abcdef123456  ", store);
  assert.equal(webKey("openrouter", store), "sk-or-v1-abcdef123456", "пробелы по краям убраны");
  assert.equal(webKey("deepseek", store), "", "ключи поставщиков не смешиваются");
  assert.equal(state.providers.find((p) => p.id === "openrouter")?.hint, "…3456");
  assert.equal(JSON.stringify(state).includes("abcdef"), false, "сам ключ в состояние панели не попадает");
  assert.deepEqual(webProviders(store).map((p) => p.id), ["openrouter"], "достаточно одного ключа");

  saveWebKey("deepseek", "sk-deepseek-0001", store);
  assert.deepEqual(webProviders(store).map((p) => p.id), ["deepseek", "openrouter"]);

  assert.throws(() => saveWebKey("openai", "ключ", store), /только DeepSeek и OpenRouter/);
  assert.throws(() => saveWebKey("openrouter", "два слова", store), /не похоже/);
  deleteWebKey("openrouter", store);
  assert.equal(webKey("openrouter", store), "");
  assert.equal(webKey("deepseek", store), "sk-deepseek-0001");
});

test("OpenRouter in the web panel: only free models, cheap Mistral and Gemini Flash", () => {
  const openrouter = WEB_PROVIDERS.find((p) => p.id === "openrouter")!;
  const paid = openrouter.models.filter((model) => !model.endsWith(":free"));
  assert.deepEqual(paid, ["mistralai/mistral-small-2603", "google/gemini-3.8-flash"]);
  assert.ok(openrouter.models.some((model) => model.endsWith(":free")));
  assert.equal(openrouter.models.some((model) => model.includes(":batch")), false, "пакетные варианты с чатом не работают");
  assert.equal(openrouter.models.some((model) => model.startsWith("anthropic/")), false);
  assert.deepEqual(WEB_PROVIDERS.find((p) => p.id === "deepseek")!.models, ["deepseek-flash", "deepseek-v4-pro"], "DeepSeek — как на Windows");
});

test("no storage — the panel says so instead of pretending to save", () => {
  assert.equal(webKeysState(null).storage.available, false);
  assert.throws(() => saveWebKey("openrouter", "sk-or-v1-x", null), /недоступно/);
});

test("the chat request is the same one the server builds for each provider", () => {
  const messages = [{ role: "assistant", provider: "deepseek", content: null, reasoning_content: "думаю", tool_calls: [{ id: "c1", name: "list_sheets", arguments: "{}" }] }];
  const tools = [{ type: "function", function: { name: "list_sheets" } }];

  const or = webChatRequest({ provider: "openrouter", model: "google/gemini-3.8-flash", messages, tools }, "sk-or-v1-key", "https://edgargew622-cloud.github.io");
  assert.equal(or.url, OPENROUTER_URL);
  const orHeaders = or.init.headers as Record<string, string>;
  assert.equal(orHeaders.Authorization, "Bearer sk-or-v1-key");
  assert.equal(orHeaders["HTTP-Referer"], "https://edgargew622-cloud.github.io");
  const orBody = JSON.parse(String(or.init.body));
  assert.equal(orBody.stream, true);
  assert.equal(orBody.tool_choice, "auto");
  assert.deepEqual(orBody.messages[0].tool_calls[0], { id: "c1", type: "function", function: { name: "list_sheets", arguments: "{}" } });
  assert.equal(orBody.messages[0].reasoning_content, undefined, "рассуждения DeepSeek другим не отправляются");

  const ds = webChatRequest({ provider: "deepseek", model: "deepseek-flash", messages, tools }, "sk-ds", "https://edgargew622-cloud.github.io");
  assert.equal(ds.url, DEEPSEEK_URL);
  const dsHeaders = ds.init.headers as Record<string, string>;
  assert.equal(dsHeaders["HTTP-Referer"], undefined, "заголовки OpenRouter — только ему");
  const dsBody = JSON.parse(String(ds.init.body));
  assert.deepEqual(dsBody.thinking, { type: "enabled" });
  assert.equal(dsBody.reasoning_effort, "high");
  assert.equal(dsBody.messages[0].reasoning_content, "думаю", "режим размышлений DeepSeek требует вернуть их");

  assert.throws(() => webChatRequest({ provider: "openrouter", model: "anthropic/claude-sonnet-5.5", messages: [], tools: [] }, "k", ""), /не разрешена/);
  assert.throws(() => webChatRequest({ provider: "openrouter", model: "google/gemini-3.8-flash:batch", messages: [], tools: [] }, "k", ""), /не разрешена/);
  assert.throws(() => webChatRequest({ provider: "openai", model: "gpt-6-luna", messages: [], tools: [] }, "k", ""), /только DeepSeek и OpenRouter/);
});

test("tools that need the local server are not given to the model in the web panel", () => {
  const local = new Set(TOOL_SPECS.filter((spec) => spec.needsLocal).map((spec) => spec.name));
  assert.deepEqual([...local].sort(), [
    "create_workbook_backup", "get_scenario", "import_file_layout", "import_file_table", "list_files", "read_file",
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
  assert.match(out, /<AppDomain>https:\/\/api\.deepseek\.com<\/AppDomain>/);
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

test("task spending: tokens, cache and the price the provider reported", async () => {
  const { readUsage } = await import("./client");
  const { spendingNote } = await import("../Taskpane");
  const claude = readUsage({ prompt_tokens: 36355, completion_tokens: 3, cost: 0.0075405, prompt_tokens_details: { cached_tokens: 36250 } })!;
  assert.deepEqual(claude, { promptTokens: 36355, completionTokens: 3, cachedTokens: 36250, cost: 0.0075405 });
  assert.equal(readUsage({ prompt_tokens: 1000, completion_tokens: 10, prompt_cache_hit_tokens: 900 })!.cachedTokens, 900, "DeepSeek");
  assert.equal(readUsage({}), null);
  assert.equal(spendingNote({ calls: 13, prompt: 420_000, cached: 380_000, completion: 6000, cost: 0.214, costKnown: true }),
    "Расход задачи: 13 обращений к модели · 426 тыс. токенов (из кэша 380 тыс.) · $0,21.");
  assert.equal(spendingNote({ calls: 1, prompt: 5000, cached: 0, completion: 40, cost: 0, costKnown: false }),
    "Расход задачи: 1 обращение к модели · 5,0 тыс. токенов · цену считает поставщик.");
});

test("conversation total: summed from the task lines, survives a reload, never double-counted", async () => {
  const { conversationSpending, conversationTotalText } = await import("../Taskpane");
  const entries = [
    { kind: "user", text: "копия" },
    { kind: "notice", text: "Расход задачи: 2 обращения к модели · 71 тыс. токенов (из кэша 35 тыс.) · $0,10." },
    { kind: "notice", text: "Расход задачи: 7 обращений к модели · 325 тыс. токенов · $0,14. Всего за беседу: 9 обращений · $0,24." }
  ];
  const total = conversationSpending(entries);
  assert.deepEqual([total.tasks, total.calls, Math.round(total.cost * 100) / 100, total.costKnown], [2, 9, 0.24, true]);
  assert.equal(conversationTotalText(total), "9 обращений · $0,24");
  const mixed = conversationSpending([...entries, { kind: "notice", text: "Расход задачи: 1 обращение к модели · 5,0 тыс. токенов · цену считает поставщик." }]);
  assert.equal(conversationTotalText(mixed), "10 обращений · от $0,24 (часть цен считает поставщик)");
});

test("prices: exact from OpenRouter, an estimate for direct providers, DeepSeek peak hours double", async () => {
  const { stepCost, deepseekPeak } = await import("./prices");
  const usage = { promptTokens: 18907, completionTokens: 5, cachedTokens: 18885 };
  // Тестовый запрос GPT-6.1 Sol 01.10.2026: 18 885 из 18 907 из кэша.
  const sol = stepCost("openai", "gpt-6.1-sol", usage)!;
  assert.equal(sol.estimated, true);
  assert.equal(Math.round(sol.cost * 1e6), Math.round(22 * 2 + 18885 * 0.1 + 5 * 10));
  assert.deepEqual(stepCost("openrouter", "anthropic/claude-sonnet-5.5", { ...usage, cost: 0.0075 }), { cost: 0.0075, estimated: false });
  assert.equal(stepCost("xai", "grok-4.20-0309-reasoning", usage), null, "нет в прайсе — честно «цену считает поставщик»");
  const weekdayPeak = new Date(Date.UTC(2026, 8, 30, 7)); // среда, 07:00 UTC
  const weekendNight = new Date(Date.UTC(2026, 9, 4, 7)); // воскресенье
  assert.equal(deepseekPeak(weekdayPeak), true);
  assert.equal(deepseekPeak(weekendNight), false);
  const plain = { promptTokens: 1_000_000, completionTokens: 0, cachedTokens: 0 };
  assert.equal(stepCost("deepseek", "deepseek-flash", plain, weekendNight)!.cost, 0.15);
  assert.equal(stepCost("deepseek", "deepseek-flash", plain, weekdayPeak)!.cost, 0.3);
  const { spendingNote } = await import("../Taskpane");
  assert.equal(spendingNote({ calls: 3, prompt: 120_000, cached: 100_000, completion: 900, cost: 0.0123, costKnown: true, estimated: true }),
    "Расход задачи: 3 обращения к модели · 121 тыс. токенов (из кэша 100 тыс.) · ≈ $0,01.");
});
