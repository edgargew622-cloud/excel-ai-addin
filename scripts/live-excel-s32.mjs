/**
 * S3.2 в настоящем Excel: условное форматирование сверяется по всем
 * запрошенным свойствам (сценарий 6 S7).
 *
 * Устроена как scripts/live-excel-s31.mjs: вместо модели отвечает скрипт
 * готовым вызовом add_conditional_format, карточка подтверждается в панели.
 * Проверяется: правила со всеми свойствами получают verified; шкала из двух
 * цветов не даёт ложной тревоги; правка существующего правила между карточкой
 * и подтверждением (тот же ID, другой цвет) останавливает операцию.
 *
 * Данные — на отдельном листе «S32». Запуск: node scripts/live-excel-s32.mjs
 */

const PORT = 9229;
const SHEET = "S32";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const DATA = [
  ["ФИО", "Отдел", "Оклад"],
  ["Иванов", "Продажи", 185000],
  ["Петрова", "Финансы", 95000],
  ["Сидоров", "Развитие", 240000],
  ["Кузнецова", "финансы", 120000],
  ["Смирнов", "Производство", 110000],
  ["Попова", "Правовой", 150000]
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
  if (request.postData && /add_conditional_format/.test(request.postData)) captured = request.postData;
  await send("Fetch.fulfillRequest", {
    requestId, responseCode: 200,
    responseHeaders: [{ name: "Content-Type", value: "text/event-stream" }],
    body: Buffer.from(answers.shift() ?? done).toString("base64")
  }).catch(() => undefined);
}]);
await send("Fetch.enable", { patterns: [{ urlPattern: "*/api/chat*" }] });

await evaluate(`
  window.__c = {
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

const resetSheet = () => excel(`
  const old = ctx.workbook.worksheets.getItemOrNullObject('${SHEET}'); old.load('isNullObject'); await ctx.sync();
  if (!old.isNullObject) { old.delete(); await ctx.sync(); }
  const s = ctx.workbook.worksheets.add('${SHEET}');
  s.getRange('A1:C${DATA.length}').values = ${JSON.stringify(DATA)};
  s.activate();
  await ctx.sync();`);

/** Правила листа так, как их отдаёт Excel. */
const readRules = () => excel(`
  const c = ctx.workbook.worksheets.getItem('${SHEET}').getRange('A1:C20').conditionalFormats;
  c.load('items/id,items/type,items/priority'); await ctx.sync();
  const out = [];
  for (const item of c.items) {
    const r = item.getRange(); r.load('address');
    let detail = null;
    if (item.type === 'CellValue') { item.cellValue.load('rule'); item.cellValue.format.fill.load('color'); item.cellValue.format.font.load(['color','bold']); }
    if (item.type === 'ContainsText') { item.textComparison.load('rule'); item.textComparison.format.fill.load('color'); item.textComparison.format.font.load(['color','bold']); }
    if (item.type === 'ColorScale') item.colorScale.load('criteria');
    if (item.type === 'DataBar') item.dataBar.positiveFormat.load('fillColor');
    await ctx.sync();
    if (item.type === 'CellValue') detail = [item.cellValue.rule.operator, item.cellValue.rule.formula1, item.cellValue.format.fill.color, item.cellValue.format.font.color, item.cellValue.format.font.bold];
    if (item.type === 'ContainsText') detail = [item.textComparison.rule.text, item.textComparison.format.fill.color, item.textComparison.format.font.color, item.textComparison.format.font.bold];
    if (item.type === 'ColorScale') { const k = item.colorScale.criteria; detail = [k.minimum?.color, k.midpoint ? k.midpoint.color : null, k.maximum?.color]; }
    if (item.type === 'DataBar') detail = [item.dataBar.positiveFormat.fillColor];
    out.push(item.type + ' ' + r.address + ' ' + JSON.stringify(detail));
  }
  return out;`);

/** Задача: вызов инструмента; before — правка книги между карточкой и «Выполнить». */
async function run(args, before) {
  const opsBefore = await evaluate("__c.ops().length");
  captured = null;
  answers.length = 0;
  answers.push(toolCall("add_conditional_format", args), done);
  await evaluate(`__c.type('правило S3.2'); true`);
  await waitFor("!!__c.button('Отправить') && !__c.button('Отправить').disabled", "кнопка «Отправить»");
  await evaluate(`__c.button('Отправить').click(); true`);
  let card = false;
  const until = Date.now() + 90000;
  await sleep(800);
  while (Date.now() < until) {
    if (await evaluate("!!__c.button('Выполнить')")) {
      card = true;
      if (before) await before();
      await evaluate(`__c.button('Выполнить').click(); true`);
      await sleep(500);
      continue;
    }
    if (await evaluate("!!__c.button('Отправить')")) break;
    await sleep(150);
  }
  await sleep(500);
  const op = (await evaluate("__c.ops()")).slice(opsBefore).find((item) => item.text.startsWith("add_conditional_format"));
  const reply = captured ? JSON.parse(captured).messages?.findLast?.((m) => m.role === "tool")?.content ?? "" : "";
  const state = /"executionState":"(\w+)"/.exec(reply)?.[1] ?? "?";
  return { card, op, state, reply };
}

const results = [];
const record = (name, ok, details) => {
  results.push(ok);
  console.log(`${ok ? "прошла " : "ПРОВАЛ "} ${name}\n        ${details.replace(/\n/g, "\n        ")}\n`);
};

await resetSheet();

const cases = [
  ["больше 150 000: заливка, цвет текста, жирный", { sheet: SHEET, address: "C2:C7", rule: "greaterThan", value: 150000, fillColor: "#FFC7CE", fontColor: "#9C0006", bold: true }],
  ["между 100 000 и 150 000: только цвет текста, не жирный", { sheet: SHEET, address: "C2:C7", rule: "between", value: 100000, value2: 150000, fontColor: "#006100", bold: false }],
  ["текст содержит «финанс»: заливка, цвет текста, жирный", { sheet: SHEET, address: "B2:B7", rule: "textContains", text: "финанс", fillColor: "#C6EFCE", fontColor: "#006100", bold: true }],
  ["шкала из трёх цветов", { sheet: SHEET, address: "C2:C7", rule: "colorScale", minColor: "#F8696B", midColor: "#FFEB84", maxColor: "#63BE7B" }],
  ["шкала из двух цветов — без ложной «середины»", { sheet: SHEET, address: "C2:C7", rule: "colorScale", minColor: "#FFFFFF", maxColor: "#5A8AC6" }],
  ["гистограмма", { sheet: SHEET, address: "C2:C7", rule: "dataBar", barColor: "#638EC6" }]
];
for (const [name, args] of cases) {
  const { card, op, state, reply } = await run(args);
  record(name, card && op?.status === "done" && state === "verified", `итог: ${op?.status ?? "—"}; executionState: ${state}${state !== "verified" ? `; ${op?.text.replace(/\n/g, " ").slice(0, 300)}` : ""}${/"priority":(\d+)/.test(reply) ? `; приоритет ${/"priority":(\d+)/.exec(reply)[1]}` : ""}`);
}
console.log("Правила листа по Excel:");
for (const line of await readRules()) console.log(`  ${line}`);
console.log("");

// Правка существующего правила с тем же ID между карточкой и подтверждением.
{
  const before = (await readRules()).length;
  const { card, op, state } = await run(
    { sheet: SHEET, address: "C2:C7", rule: "lessThan", value: 100000, fillColor: "#FFEB9C" },
    () => excel(`
      const c = ctx.workbook.worksheets.getItem('${SHEET}').getRange('C2:C7').conditionalFormats; c.load('items/type'); await ctx.sync();
      const first = c.items.find((item) => item.type === 'CellValue');
      first.cellValue.format.fill.color = '#00B050';
      await ctx.sync();`)
  );
  const after = (await readRules()).length;
  record(
    "правка прежнего правила после предпросмотра (тот же ID, другая заливка)",
    card && state === "failed_before_write" && after === before,
    `итог: ${op?.status ?? "—"}; executionState: ${state}; правил было ${before}, стало ${after}\n${op?.text.replace(/\n/g, " ").slice(0, 300) ?? ""}`
  );
}

console.log(`прошло ${results.filter(Boolean).length} из ${results.length}`);
socket.close();
process.exit(results.every(Boolean) ? 0 : 1);
