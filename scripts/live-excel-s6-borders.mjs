/**
 * S6 в настоящем Excel: внешняя рамка и цепочка отмен оформления.
 *
 * На этапе 6 у сетки «Сотрудники!A1:E7» в format-smoke.xlsx пропала внешняя
 * рамка со всех сторон; связь с отменой была предположением. Здесь та же
 * последовательность воспроизводится через панель на свежем листе:
 * сетка → толстая рамка → серые внутренние линии → соседнее оформление
 * (заливка шапки и сетка соседнего столбца, делящего с таблицей край) →
 * отмены по одной до конца. После каждого шага читаются все четыре края
 * каждой ячейки A1:G9; каждая отмена сверяется с состоянием до своей операции.
 *
 * Запуск: node scripts/live-excel-s6-borders.mjs
 */

const PORT = 9229;
const SHEET = "S6";
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
listeners.set("Fetch.requestPaused", [async ({ requestId }) => {
  await send("Fetch.fulfillRequest", {
    requestId, responseCode: 200,
    responseHeaders: [{ name: "Content-Type", value: "text/event-stream" }],
    body: Buffer.from(answers.shift() ?? done).toString("base64")
  }).catch(() => undefined);
}]);
await send("Fetch.enable", { patterns: [{ urlPattern: "*/api/chat*" }] });

await evaluate(`
  window.__b = {
    button(text) { return [...document.querySelectorAll('button')].find((b) => b.textContent.trim().startsWith(text)); },
    type(text) {
      const area = document.querySelector('textarea');
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(area, text);
      area.dispatchEvent(new Event('input', { bubbles: true }));
    },
    ops() { return [...document.querySelectorAll('.op')].map((op) => ({ status: op.className.replace('op', '').trim(), text: op.innerText })); },
    log() { return document.querySelector('.log')?.innerText ?? ''; }
  };
  (() => { const box = [...document.querySelectorAll('label')].find((l) => l.textContent.includes('Только анализ'))?.querySelector('input'); if (box && box.checked) box.click(); })();
  true`);
console.log(`Панель: ${await evaluate("document.querySelector('.persistence-note')?.innerText ?? ''")}\n`);

// Свежий лист со «штатом», как в format-smoke.xlsx.
await excel(`
  const old = ctx.workbook.worksheets.getItemOrNullObject('${SHEET}'); old.load('isNullObject'); await ctx.sync();
  if (!old.isNullObject) { old.delete(); await ctx.sync(); }
  const s = ctx.workbook.worksheets.add('${SHEET}');
  s.getRange('A1:E7').values = [
    ['ФИО','Должность','Отдел','Оклад','Комментарий'],
    ['Иванов Иван','Ведущий специалист','Продажи',185000,'Перевод'],
    ['Петрова Анна','Бухгалтер','Финансы',95000,'—'],
    ['Сидоров Пётр','Руководитель','Развитие',240000,'Испытательный срок'],
    ['Кузнецова Мария','Аналитик','Финансы',120000,'—'],
    ['Смирнов Олег','Инженер','Производство',110000,'Отпуск'],
    ['Попова Елена','Юрист','Правовой',150000,'—']];
  s.activate();
  await ctx.sync();`);

/** Все четыре края каждой ячейки A1:G9 — так, как их отдаёт Excel. */
const EDGES = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"];
const state = () => excel(`
  const s = ctx.workbook.worksheets.getItem('${SHEET}');
  const items = [];
  for (let r = 1; r <= 9; r++) for (const c of 'ABCDEFG') {
    const cell = s.getRange(c + r);
    for (const e of ${JSON.stringify(EDGES)}) { const b = cell.format.borders.getItem(e); b.load(['style','weight','color']); items.push([c + r, e, b]); }
  }
  await ctx.sync();
  const out = {};
  for (const [cell, e, b] of items) out[cell + ' ' + e] = b.style === 'None' ? 'None' : b.style + '|' + b.weight + '|' + b.color;
  return out;`);
const diff = (a, b) => Object.keys(a).filter((key) => a[key] !== b[key]).map((key) => `${key}: ${a[key]} → ${b[key]}`);
const summary = (s) => {
  // Внешние края таблицы A1:E7 и одна внутренняя линия — коротко для лога.
  const pick = (key) => (s[key] === "None" ? "нет" : s[key].split("|").slice(0, 2).join(" "));
  return `верх ${pick("A1 EdgeTop")}, низ ${pick("A7 EdgeBottom")}, лево ${pick("A1 EdgeLeft")}, право ${pick("E1 EdgeRight")}, внутри ${pick("A1 EdgeRight")}`;
};

async function run(args) {
  const opsBefore = await evaluate("__b.ops().length");
  answers.length = 0;
  answers.push(toolCall("format_range", args), done);
  await evaluate(`__b.type('оформление S6'); true`);
  await waitFor("!!__b.button('Отправить') && !__b.button('Отправить').disabled", "кнопка «Отправить»");
  await evaluate(`__b.button('Отправить').click(); true`);
  const until = Date.now() + 90000;
  await sleep(800);
  while (Date.now() < until) {
    if (await evaluate("!!__b.button('Выполнить')")) {
      await evaluate(`__b.button('Выполнить').click(); true`);
      await sleep(500);
      continue;
    }
    if (await evaluate("!!__b.button('Отправить')")) break;
    await sleep(150);
  }
  await sleep(500);
  return (await evaluate("__b.ops()")).slice(opsBefore).find((item) => item.text.startsWith("format_range"));
}

const steps = [
  ["сетка тонкая чёрная A1:E7", { sheet: SHEET, address: "A1:E7", borders: "all" }],
  ["толстая рамка A1:E7", { sheet: SHEET, address: "A1:E7", borders: "outline", borderWeight: "Thick" }],
  ["серые внутренние линии A1:E7", { sheet: SHEET, address: "A1:E7", borders: "inside", borderColor: "#A6A6A6" }],
  ["заливка шапки A1:E1", { sheet: SHEET, address: "A1:E1", fillColor: "#D9D9D9", bold: true }],
  ["сетка соседнего столбца F1:F7 (общий край с E)", { sheet: SHEET, address: "F1:F7", borders: "all" }],
  ["рамка под таблицей A8:E8 (общий край с 7-й строкой)", { sheet: SHEET, address: "A8:E8", borders: "outline" }]
];

const history = [await state()];
console.log(`исходно: ${summary(history[0])}`);
for (const [name, args] of steps) {
  const op = await run(args);
  const now = await state();
  history.push(now);
  console.log(`${op?.status === "done" ? "выполнено" : `ИТОГ ${op?.status}`}: ${name}\n        ${summary(now)}`);
}
console.log("");

// Отмены по одной: после k-й отмены состояние должно совпасть с состоянием
// до k-й с конца операции — по всем краям всех ячеек.
const results = [];
for (let k = steps.length; k >= 1; k--) {
  await waitFor("!!__b.button('Отменить') && !__b.button('Отменить').disabled", "кнопка «Отменить»", 20000);
  const label = await evaluate("__b.button('Отменить').textContent.trim()");
  const logBefore = (await evaluate("__b.log()")).length;
  await evaluate(`__b.button('Отменить').click(); true`);
  await sleep(2500);
  await waitFor("!!__b.button('Отправить')", "панель свободна");
  const said = (await evaluate("__b.log()")).slice(logBefore).trim().split("\n").filter(Boolean).slice(-1)[0] ?? "";
  const now = await state();
  const wrong = diff(history[k - 1], now);
  const ok = wrong.length === 0 && /^Отменено/.test(said);
  results.push(ok);
  console.log(`${ok ? "прошла " : "ПРОВАЛ "} отмена «${steps[k - 1][0]}» (${label})\n        панель: ${said.slice(0, 200)}\n        ${summary(now)}${wrong.length ? `\n        расходится с состоянием до операции (${wrong.length}): ${wrong.slice(0, 12).join("; ")}` : ""}`);
}

console.log(`\nпрошло ${results.filter(Boolean).length} из ${results.length}`);
socket.close();
process.exit(results.every(Boolean) ? 0 : 1);
