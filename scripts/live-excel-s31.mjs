/**
 * S3.1 в настоящем Excel: сводные с несколькими уровнями (сценарий 5 S7).
 *
 * Устроена как scripts/live-excel-s7.mjs: вместо модели отвечает скрипт
 * готовым вызовом create_pivot_table, карточка подтверждается в панели.
 * Excel сам сводные не искажает, поэтому здесь проверяется другое: новая
 * сверка принимает настоящие макеты — вложенные уровни, одинаковые подписи
 * под разными родителями, пустые значения, регистр, пробел в конце, числа
 * рядом с текстом, несколько полей значений — без ложных тревог, и итог
 * операции — verified с проверкой всех уровней.
 *
 * Данные — на отдельном листе «S31», каждая сводная — на своём листе.
 * Запуск: node scripts/live-excel-s31.mjs [--save]
 * --save сохраняет книгу: по сохранённому файлу итоги сверяются ещё раз.
 */

const PORT = 9229;
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const DATA = [
  ["Город", "Статус", "Менеджер", "Сумма", "Код"],
  ["Москва", "Новая", "Иванов", 900, "1"],
  ["Москва", "Закрыта", "Сидоров", 450, 1],
  ["москва", "Новая", "Петрова", 100, "1"],
  ["Казань", "Новая", "Иванов", 700, 1],
  ["Казань", "Закрыта", "Сидоров", 800, "1"],
  ["Казань", "Новая", "Иванов", 300, 2],
  ["Омск", "Закрыта", "Петрова", 1300, 2],
  ["Омск", "", "Иванов", 50, 2],
  ["", "Новая", "Сидоров", 5, 2],
  ["Москва ", "Новая", "Иванов", 70, 1]
];

const SCENARIOS = [
  { name: "два уровня: город › статус, сумма", dest: "S31_2", rows: ["Город", "Статус"], values: [{ field: "Сумма", aggregation: "sum" }] },
  { name: "три уровня: город › статус › менеджер, сумма", dest: "S31_3", rows: ["Город", "Статус", "Менеджер"], values: [{ field: "Сумма", aggregation: "sum" }] },
  { name: "два уровня, три поля значений: сумма, среднее, количество", dest: "S31_V", rows: ["Статус", "Город"], values: [{ field: "Сумма", aggregation: "sum" }, { field: "Сумма", aggregation: "average" }, { field: "Сумма", aggregation: "count" }] },
  { name: "числа и текст: код › статус, максимум", dest: "S31_N", rows: ["Код", "Статус"], values: [{ field: "Сумма", aggregation: "max" }] }
];

const list = await (await fetch(`http://127.0.0.1:${PORT}/json/list`)).json();
const page = list.find((target) => target.type === "page" && target.url.includes("taskpane.html"));
if (!page) throw new Error("Панель надстройки не найдена: откройте её в Excel.");
const socket = new WebSocket(page.webSocketDebuggerUrl);
let nextId = 0;
const waiting = new Map();
const listeners = new Map();
socket.addEventListener("message", (event) => {
  const message = JSON.parse(event.data);
  if (message.id && waiting.has(message.id)) {
    const { resolve, reject } = waiting.get(message.id);
    waiting.delete(message.id);
    message.error ? reject(new Error(message.error.message)) : resolve(message.result);
  } else if (message.method) {
    for (const listener of listeners.get(message.method) ?? []) listener(message.params);
  }
});
await new Promise((resolve) => socket.addEventListener("open", resolve));
const send = (method, params = {}) => new Promise((resolve, reject) => {
  const id = ++nextId;
  waiting.set(id, { resolve, reject });
  socket.send(JSON.stringify({ id, method, params }));
});
const evaluate = async (expression) => {
  const { result, exceptionDetails } = await send("Runtime.evaluate", { expression, returnByValue: true, awaitPromise: true });
  if (exceptionDetails) throw new Error(exceptionDetails.exception?.description ?? exceptionDetails.text);
  return result.value;
};
const waitFor = async (expression, what, ms = 60000) => {
  const until = Date.now() + ms;
  while (Date.now() < until) {
    if (await evaluate(expression).catch(() => false)) return;
    await sleep(150);
  }
  throw new Error(`не дождались: ${what}`);
};
const excel = (body) => evaluate(`(async () => { let out; await Excel.run(async (ctx) => { out = await (async () => { ${body} })(); }); return out; })()`);

const sse = (chunks) => [...chunks.map((chunk) => `data: ${JSON.stringify(chunk)}`), "data: [DONE]", ""].join("\n\n");
const toolCall = (name, args) => sse([{ choices: [{ delta: { tool_calls: [{ index: 0, id: `call_${Date.now()}`, function: { name, arguments: JSON.stringify(args) } }] }, finish_reason: "tool_calls" }] }]);
const done = sse([{ choices: [{ delta: { content: "Готово." }, finish_reason: "stop" }] }]);
const answers = [];
let captured = null;
listeners.set("Fetch.requestPaused", [async ({ requestId, request }) => {
  // Ответ инструмента виден во втором запросе к «модели» — в истории беседы.
  if (request.postData && /create_pivot_table/.test(request.postData)) captured = request.postData;
  await send("Fetch.fulfillRequest", {
    requestId, responseCode: 200,
    responseHeaders: [{ name: "Content-Type", value: "text/event-stream" }],
    body: Buffer.from(answers.shift() ?? done).toString("base64")
  }).catch(() => undefined);
}]);
await send("Fetch.enable", { patterns: [{ urlPattern: "*/api/chat*" }] });

await evaluate(`
  window.__p = {
    button(text) { return [...document.querySelectorAll('button')].find((b) => b.textContent.trim().startsWith(text)); },
    type(text) {
      const area = document.querySelector('textarea');
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(area, text);
      area.dispatchEvent(new Event('input', { bubbles: true }));
    },
    ops() { return [...document.querySelectorAll('.op')].map((op) => ({ status: op.className.replace('op', '').trim(), text: op.innerText })); }
  };
  (() => { const box = [...document.querySelectorAll('label')].find((l) => l.textContent.includes('Только анализ'))?.querySelector('input'); if (box && box.checked) box.click(); })();
  true`);
console.log(`Панель: ${await evaluate("document.querySelector('.persistence-note')?.innerText ?? ''")}\n`);

// Данные и чистые листы назначения.
await excel(`
  for (const name of ${JSON.stringify(["S31", ...SCENARIOS.map((s) => s.dest)])}) {
    const old = ctx.workbook.worksheets.getItemOrNullObject(name); old.load('isNullObject'); await ctx.sync();
    if (!old.isNullObject) { old.delete(); await ctx.sync(); }
  }
  const s = ctx.workbook.worksheets.add('S31');
  s.getRange('A1:E${DATA.length}').values = ${JSON.stringify(DATA)};
  for (const name of ${JSON.stringify(SCENARIOS.map((s) => s.dest))}) ctx.workbook.worksheets.add(name);
  s.activate();
  await ctx.sync();`);

const results = [];
for (const scenario of SCENARIOS) {
  const opsBefore = await evaluate("__p.ops().length");
  captured = null;
  answers.length = 0;
  answers.push(toolCall("create_pivot_table", { sheet: "S31", sourceAddress: `A1:E${DATA.length}`, destSheet: scenario.dest, destAddress: "A1", rows: scenario.rows, values: scenario.values }), done);
  await evaluate(`__p.type('сводная S3.1'); true`);
  await waitFor("!!__p.button('Отправить') && !__p.button('Отправить').disabled", "кнопка «Отправить»");
  await evaluate(`__p.button('Отправить').click(); true`);
  let card = false;
  const until = Date.now() + 90000;
  await sleep(800);
  while (Date.now() < until) {
    if (await evaluate("!!__p.button('Выполнить')")) {
      card = true;
      await evaluate(`__p.button('Выполнить').click(); true`);
      await sleep(500);
      continue;
    }
    if (await evaluate("!!__p.button('Отправить')")) break;
    await sleep(150);
  }
  await sleep(500);
  const op = (await evaluate("__p.ops()")).slice(opsBefore).find((item) => item.text.startsWith("create_pivot_table"));
  const layout = await excel(`
    const p = ctx.workbook.worksheets.getItem('${scenario.dest}').pivotTables; p.load('items/name'); await ctx.sync();
    if (!p.items.length) return null;
    const r = p.items[0].layout.getRange(); r.load(['address', 'values']); await ctx.sync();
    return { address: r.address, values: r.values };`);
  const reply = captured ? JSON.parse(captured).messages?.findLast?.((m) => m.role === "tool")?.content ?? "" : "";
  const state = /"executionState":"(\w+)"/.exec(reply)?.[1] ?? "?";
  const ok = card && op?.status === "done" && state === "verified" && /всех уровнях/.test(reply);
  results.push(ok);
  console.log(`${ok ? "прошла " : "ПРОВАЛ "} ${scenario.name}`);
  console.log(`        итог: ${op?.status ?? "—"}; executionState: ${state}; ${(/"note":"([^"]+)"/.exec(reply)?.[1]) ?? op?.text.replace(/\n/g, " ").slice(0, 300) ?? ""}`);
  console.log(`        сводная ${layout?.address ?? "не построена"}:`);
  for (const row of layout?.values ?? []) console.log(`          ${JSON.stringify(row)}`);
  console.log("");
}

if (process.argv.includes("--save")) {
  await excel(`ctx.workbook.save(Excel.SaveBehavior.save); await ctx.sync(); return true;`);
  console.log("Книга сохранена.");
}
console.log(`прошло ${results.filter(Boolean).length} из ${results.length}`);
socket.close();
process.exit(results.every(Boolean) ? 0 : 1);
