/**
 * S6 в настоящем Excel: несколько диаграмм на одном листе (сценарий 9 S7).
 *
 * Три диаграммы подряд по одним данным: каждая по умолчанию встаёт в одну
 * и ту же ячейку правее данных. Прежде новая опускалась один раз, под те,
 * с которыми пересеклась вначале, и третья ложилась на вторую. Здесь
 * проверяется, что после каждой ни одна пара диаграмм не пересекается
 * и что ответ называет фактическое положение.
 *
 * Запуск: node scripts/live-excel-s6-charts.mjs
 */

const PORT = 9229;
const SHEET = "S6Д";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

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
  if (request.postData && /create_chart/.test(request.postData)) captured = request.postData;
  await send("Fetch.fulfillRequest", {
    requestId, responseCode: 200,
    responseHeaders: [{ name: "Content-Type", value: "text/event-stream" }],
    body: Buffer.from(answers.shift() ?? done).toString("base64")
  }).catch(() => undefined);
}]);
await send("Fetch.enable", { patterns: [{ urlPattern: "*/api/chat*" }] });

await evaluate(`
  window.__d = {
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

await excel(`
  const old = ctx.workbook.worksheets.getItemOrNullObject('${SHEET}'); old.load('isNullObject'); await ctx.sync();
  if (!old.isNullObject) { old.delete(); await ctx.sync(); }
  const s = ctx.workbook.worksheets.add('${SHEET}');
  s.getRange('A1:C4').values = [['Месяц','Выручка','Расходы'],['Январь',120,80],['Февраль',150,90],['Март',170,95]];
  s.activate();
  await ctx.sync();`);

const boxes = () => excel(`
  const c = ctx.workbook.worksheets.getItem('${SHEET}').charts; c.load('items/name,items/top,items/left,items/width,items/height'); await ctx.sync();
  return c.items.map((i) => ({ name: i.name, top: i.top, left: i.left, width: i.width, height: i.height }));`);
const overlap = (a, b) => a.left < b.left + b.width && b.left < a.left + a.width && a.top < b.top + b.height && b.top < a.top + a.height;

const results = [];
for (const chartType of ["ColumnClustered", "Line", "Pie", "BarClustered"]) {
  const opsBefore = await evaluate("__d.ops().length");
  captured = null;
  answers.length = 0;
  answers.push(toolCall("create_chart", { sheet: SHEET, address: "A1:C4", chartType }), done);
  await evaluate(`__d.type('диаграмма S6'); true`);
  await waitFor("!!__d.button('Отправить') && !__d.button('Отправить').disabled", "кнопка «Отправить»");
  await evaluate(`__d.button('Отправить').click(); true`);
  const until = Date.now() + 90000;
  await sleep(800);
  while (Date.now() < until) {
    if (await evaluate("!!__d.button('Выполнить')")) { await evaluate(`__d.button('Выполнить').click(); true`); await sleep(500); continue; }
    if (await evaluate("!!__d.button('Отправить')")) break;
    await sleep(150);
  }
  await sleep(500);
  const op = (await evaluate("__d.ops()")).slice(opsBefore).find((item) => item.text.startsWith("create_chart"));
  const reply = captured ? JSON.parse(captured).messages?.findLast?.((m) => m.role === "tool")?.content ?? "" : "";
  const state = /"executionState":"(\w+)"/.exec(reply)?.[1] ?? "?";
  const all = await boxes();
  const pairs = [];
  for (let i = 0; i < all.length; i++) for (let j = i + 1; j < all.length; j++) if (overlap(all[i], all[j])) pairs.push(`${all[i].name} × ${all[j].name}`);
  const position = /"position":(\{[^}]+\})/.exec(reply)?.[1];
  const ok = state === "verified" && pairs.length === 0 && Boolean(position);
  results.push(ok);
  console.log(`${ok ? "прошла " : "ПРОВАЛ "} диаграмма ${all.length}: ${chartType} — ${op?.status}, ${state}`);
  console.log(`        ${/"placementNote":"([^"]+)"/.exec(reply)?.[1] ?? "без сдвига"}; положение по ответу: ${position ?? "не названо"}`);
  console.log(`        на листе: ${all.map((b) => `${b.name} [${Math.round(b.top)}–${Math.round(b.top + b.height)}]`).join(", ")}`);
  console.log(`        пересечений: ${pairs.length ? pairs.join(", ") : "нет"}\n`);
}

console.log(`прошло ${results.filter(Boolean).length} из ${results.length}`);
socket.close();
process.exit(results.every(Boolean) ? 0 : 1);
