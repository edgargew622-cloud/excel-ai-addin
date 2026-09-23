/**
 * Живая проверка S2 в настоящем Excel (план стабилизации).
 *
 * Панель надстройки — это WebView2, то есть Chromium. Если Excel запущен
 * с переменной WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=9229,
 * к открытой панели можно подключиться по протоколу DevTools. Тогда:
 *
 * - вместо модели отвечает этот скрипт: первый запрос к /api/chat получает
 *   готовый вызов инструмента с точными аргументами, следующий — «Готово».
 *   Результат не зависит от того, как модель поняла просьбу;
 * - книга меняется между предпросмотром и подтверждением через тот же
 *   Office.js, из самой панели — ровно так, как её поменял бы человек;
 * - «Выполнить» нажимается в самой панели, а итог читается из ленты
 *   и из книги.
 *
 * Запуск: Excel открыт с копией pivot-smoke.xlsx, панель открыта.
 *   node scripts/live-excel-s2.mjs
 */

const PORT = 9229;
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function pageSocket() {
  const list = await (await fetch(`http://127.0.0.1:${PORT}/json/list`)).json();
  const page = list.find((target) => target.type === "page" && target.url.includes("taskpane.html"));
  if (!page) throw new Error("Панель надстройки не найдена: откройте её в Excel.");
  return page.webSocketDebuggerUrl;
}

function cdp(url) {
  const socket = new WebSocket(url);
  let next = 0;
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
  return {
    ready: new Promise((resolve) => socket.addEventListener("open", resolve)),
    send(method, params = {}) {
      const id = ++next;
      socket.send(JSON.stringify({ id, method, params }));
      return new Promise((resolve, reject) => waiting.set(id, { resolve, reject }));
    },
    on(method, listener) {
      if (!listeners.has(method)) listeners.set(method, []);
      listeners.get(method).push(listener);
    },
    close() { socket.close(); }
  };
}

const sse = (chunks) => [...chunks.map((chunk) => `data: ${JSON.stringify(chunk)}`), "data: [DONE]", ""].join("\n\n");
const toolCall = (name, args) => sse([{
  choices: [{ delta: { tool_calls: [{ index: 0, id: `call_${Date.now()}`, function: { name, arguments: JSON.stringify(args) } }] }, finish_reason: "tool_calls" }]
}]);
const done = sse([{ choices: [{ delta: { content: "Готово." }, finish_reason: "stop" }] }]);

const client = cdp(await pageSocket());
await client.ready;

// Ответы модели выдаёт очередь: каждый запрос к /api/chat берёт следующий.
const answers = [];
let chatRequests = 0;
client.on("Fetch.requestPaused", async ({ requestId }) => {
  chatRequests += 1;
  const body = answers.shift() ?? done;
  await client.send("Fetch.fulfillRequest", {
    requestId,
    responseCode: 200,
    responseHeaders: [{ name: "Content-Type", value: "text/event-stream" }],
    body: Buffer.from(body).toString("base64")
  }).catch(() => undefined);
});
await client.send("Fetch.enable", { patterns: [{ urlPattern: "*/api/chat*" }] });

const evaluate = async (expression) => {
  const { result, exceptionDetails } = await client.send("Runtime.evaluate", { expression, returnByValue: true, awaitPromise: true });
  if (exceptionDetails) throw new Error(exceptionDetails.exception?.description ?? exceptionDetails.text);
  return result.value;
};
const waitFor = async (expression, what, ms = 30000) => {
  const until = Date.now() + ms;
  while (Date.now() < until) {
    if (await evaluate(expression).catch(() => false)) return true;
    await sleep(150);
  }
  throw new Error(`не дождались: ${what}`);
};

/** Действие в настоящей книге через Office.js из панели. */
const excel = (body) => evaluate(`(async () => { let out; await Excel.run(async (ctx) => { out = await (async () => { ${body} })(); }); return out; })()`);

await evaluate(`
  window.__s2 = {
    button(text) { return [...document.querySelectorAll('button')].find((b) => b.textContent.trim().startsWith(text)); },
    type(text) {
      const area = document.querySelector('textarea');
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(area, text);
      area.dispatchEvent(new Event('input', { bubbles: true }));
    },
    lastOp() {
      const ops = [...document.querySelectorAll('.op')];
      const op = ops[ops.length - 1];
      return op ? { status: op.className.replace('op', '').trim(), text: op.innerText } : null;
    },
    opCount() { return document.querySelectorAll('.op').length; }
  };
  // Режим «Только анализ» запрещает запись: снимаем, если стоит.
  (() => {
    const box = [...document.querySelectorAll('label')].find((l) => l.textContent.includes('Только анализ'))?.querySelector('input');
    if (box && box.checked) box.click();
  })();
  true`);
await sleep(300);

async function countPivots() {
  return excel(`const p = ctx.workbook.pivotTables; p.load('items/name'); await ctx.sync(); return p.items.length;`);
}

/**
 * Один прогон: вызвать инструмент, при появлении карточки сделать правку
 * книги и нажать «Выполнить». Возвращает итог операции из ленты.
 */
async function run(args, betweenPreviewAndConfirm) {
  const opsBefore = await evaluate("__s2.opCount()");
  answers.push(toolCall("create_pivot_table", args), done);
  await evaluate(`__s2.type('проверка S2'); true`);
  await waitFor("!!__s2.button('Отправить') && !__s2.button('Отправить').disabled", "кнопка «Отправить»");
  await evaluate(`__s2.button('Отправить').click(); true`);

  // Либо карточка подтверждения, либо отказ ещё при подготовке.
  await waitFor(`!!__s2.button('Выполнить') || (__s2.opCount() > ${opsBefore} && !['running'].includes(__s2.lastOp().status))`, "карточка или отказ", 60000);
  const card = await evaluate("!!__s2.button('Выполнить')");
  if (card) {
    if (betweenPreviewAndConfirm) await betweenPreviewAndConfirm();
    await evaluate(`__s2.button('Выполнить').click(); true`);
  }
  await waitFor("!!__s2.button('Отправить')", "задача завершилась", 90000);
  await sleep(300);
  return { card, op: await evaluate("__s2.lastOp()") };
}

const results = [];
const record = (name, ok, details) => { results.push({ name, ok, details }); console.log(`${ok ? "прошла " : "ПРОВАЛ "} ${name}\n        ${details.replace(/\n/g, "\n        ")}`); };
const SUMS = (sheet, destAddress) => ({ sheet, sourceAddress: "A1:D13", ...(destAddress ? { destAddress } : {}), rows: ["Город"], values: [{ field: "Сумма" }] });

// Прогон начинается с чистой книги: сводные прошлых прогонов убираются.
await excel(`const p = ctx.workbook.pivotTables; p.load('items/name'); await ctx.sync(); for (const x of p.items) x.delete(); await ctx.sync();`);

console.log(`Панель: ${await evaluate("document.querySelector('.persistence-note')?.innerText ?? ''")}\n`);

// 1. Формула ="" на месте — занятая ячейка.
{
  await excel(`ctx.workbook.worksheets.getItem('Заказы').getRange('G3').formulas = [['=""']]; await ctx.sync();`);
  const pivots = await countPivots();
  const { card, op } = await run(SUMS("Заказы", "F1"));
  const kept = await excel(`const r = ctx.workbook.worksheets.getItem('Заказы').getRange('G3'); r.load('formulas'); await ctx.sync(); return r.formulas[0][0];`);
  record("формула =\"\" — занятое место", !card && /затёрты/.test(op?.text ?? "") && kept === '=""' && (await countPivots()) === pivots,
    `карточка: ${card ? "была" : "нет"}; G3: ${kept}; итог: ${op?.text}`);
  await excel(`ctx.workbook.worksheets.getItem('Заказы').getRange('G3').clear(); await ctx.sync();`);
}

// 2. Формула ="" появилась на месте после предпросмотра.
{
  const pivots = await countPivots();
  const { card, op } = await run(SUMS("Заказы", "F1"), async () => {
    await excel(`ctx.workbook.worksheets.getItem('Заказы').getRange('F2').formulas = [['=""']]; await ctx.sync();`);
  });
  const kept = await excel(`const r = ctx.workbook.worksheets.getItem('Заказы').getRange('F2'); r.load('formulas'); await ctx.sync(); return r.formulas[0][0];`);
  record("=\"\" после предпросмотра останавливает сводную", card && /перестало быть пустым/.test(op?.text ?? "") && kept === '=""' && (await countPivots()) === pivots,
    `карточка: ${card ? "была" : "нет"}; F2: ${kept}; сводных стало: ${await countPivots()}; итог: ${op?.text}`);
  await excel(`ctx.workbook.worksheets.getItem('Заказы').getRange('F2').clear(); await ctx.sync();`);
}

// 3. Значения источника изменились при тех же формулах.
{
  const pivots = await countPivots();
  const before = await excel(`const r = ctx.workbook.worksheets.getItem('Связанные').getRange('C2'); r.load(['formulas','values']); await ctx.sync(); return [r.formulas[0][0], r.values[0][0]];`);
  const { card, op } = await run(SUMS("Связанные"), async () => {
    await excel(`ctx.workbook.worksheets.getItem('Курс').getRange('B1').values = [[2]]; await ctx.sync();`);
  });
  const after = await excel(`const r = ctx.workbook.worksheets.getItem('Связанные').getRange('C2'); r.load(['formulas','values']); await ctx.sync(); return [r.formulas[0][0], r.values[0][0]];`);
  record("изменились значения источника при тех же формулах", card && /изменились после предпросмотра/.test(op?.text ?? "") && (await countPivots()) === pivots,
    `C2 до: ${before.join(" → ")}; после: ${after.join(" → ")}; итог: ${op?.text}`);
  await excel(`ctx.workbook.worksheets.getItem('Курс').getRange('B1').values = [[1]]; await ctx.sync();`);
}

// 4. Защита листа включена после предпросмотра.
{
  const pivots = await countPivots();
  const { card, op } = await run(SUMS("Связанные"), async () => {
    await excel(`ctx.workbook.worksheets.getItem('Связанные').protection.protect(); await ctx.sync();`);
  });
  record("защита листа после предпросмотра", card && /защищён/.test(op?.text ?? "") && (await countPivots()) === pivots, `итог: ${op?.text}`);
  await excel(`ctx.workbook.worksheets.getItem('Связанные').protection.unprotect(); await ctx.sync();`);
}

// 5. Обычная сводная после всех отказов.
{
  const pivots = await countPivots();
  const { card, op } = await run(SUMS("Связанные"));
  const layout = await excel(`
    const list = ctx.workbook.worksheets.getItem('Связанные').pivotTables; list.load('items/name'); await ctx.sync();
    const p = list.items[list.items.length - 1]; if (!p) return null;
    p.layout.load('layoutType'); const r = p.layout.getRange(); r.load(['address','values']); await ctx.sync();
    return { layout: p.layout.layoutType, address: r.address, total: r.values[r.values.length - 1] };`);
  record("обычная сводная строится", card && /done/.test(op?.status ?? "") && (await countPivots()) === pivots + 1 && layout?.total?.[1] === 11000,
    `статус: ${op?.status}; сводная: ${JSON.stringify(layout)}`);
}

console.log(`\nзапросов к «модели»: ${chatRequests}; прошло ${results.filter((r) => r.ok).length} из ${results.length}`);
client.close();
process.exit(results.every((r) => r.ok) ? 0 : 1);
