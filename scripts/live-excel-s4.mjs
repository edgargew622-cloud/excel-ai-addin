/**
 * Живая проверка S4 в настоящем Excel (план стабилизации).
 *
 * Устроена как scripts/live-excel-s2.mjs: Excel запущен с
 * WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=9229,
 * панель открыта; вместо модели отвечает скрипт, книга меняется через
 * Office.js из самой панели, «Выполнить» нажимается в панели.
 *
 * Сценарии:
 * 1. Группа остановилась на полпути: вторая запись идёт в неугловую ячейку
 *    объединения, которую Excel молча не принимает. Следующая изменяющая
 *    команда того же шага не должна выполниться (сценарий 8 плана).
 * 2. Фильтр на защищённом листе отклоняется до карточки.
 *
 * Запуск: node scripts/live-excel-s4.mjs
 */

const PORT = 9229;
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
const on = (method, listener) => {
  if (!listeners.has(method)) listeners.set(method, []);
  listeners.get(method).push(listener);
};

const sse = (chunks) => [...chunks.map((chunk) => `data: ${JSON.stringify(chunk)}`), "data: [DONE]", ""].join("\n\n");
const toolCalls = (calls) => sse([{
  choices: [{
    delta: { tool_calls: calls.map(([name, args], index) => ({ index, id: `call_${Date.now()}_${index}`, function: { name, arguments: JSON.stringify(args) } })) },
    finish_reason: "tool_calls"
  }]
}]);
const done = sse([{ choices: [{ delta: { content: "Готово." }, finish_reason: "stop" }] }]);

const answers = [];
let chatRequests = 0;
on("Fetch.requestPaused", async ({ requestId }) => {
  chatRequests += 1;
  await send("Fetch.fulfillRequest", {
    requestId, responseCode: 200,
    responseHeaders: [{ name: "Content-Type", value: "text/event-stream" }],
    body: Buffer.from(answers.shift() ?? done).toString("base64")
  }).catch(() => undefined);
});
await send("Fetch.enable", { patterns: [{ urlPattern: "*/api/chat*" }] });

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

await evaluate(`
  window.__s4 = {
    button(text) { return [...document.querySelectorAll('button')].find((b) => b.textContent.trim().startsWith(text)); },
    type(text) {
      const area = document.querySelector('textarea');
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(area, text);
      area.dispatchEvent(new Event('input', { bubbles: true }));
    },
    ops() { return [...document.querySelectorAll('.op')].map((op) => ({ status: op.className.replace('op', '').trim(), text: op.innerText })); }
  };
  (() => {
    const box = [...document.querySelectorAll('label')].find((l) => l.textContent.includes('Только анализ'))?.querySelector('input');
    if (box && box.checked) box.click();
  })();
  true`);
await sleep(300);
console.log(`Панель: ${await evaluate("document.querySelector('.persistence-note')?.innerText ?? ''")}\n`);

/** Задача: ответ модели — эти вызовы; на каждую карточку — «Выполнить». */
async function task(calls, { confirmCards = 5 } = {}) {
  const opsBefore = await evaluate("__s4.ops().length");
  const requestsBefore = chatRequests;
  // Остановленная задача не забирает заключительный ответ: без очистки
  // он достался бы следующей задаче вместо её команд.
  answers.length = 0;
  answers.push(toolCalls(calls), done);
  await evaluate(`__s4.type('проверка S4'); true`);
  await waitFor("!!__s4.button('Отправить') && !__s4.button('Отправить').disabled", "кнопка «Отправить»");
  await evaluate(`__s4.button('Отправить').click(); true`);
  // Задача началась, когда ответ «модели» забран: до этого кнопка
  // «Отправить» ещё видна, и цикл ниже закончился бы сразу.
  const started = Date.now() + 30000;
  while (chatRequests === requestsBefore && Date.now() < started) await sleep(100);
  await sleep(300);
  let cards = 0;
  const until = Date.now() + 120000;
  while (Date.now() < until) {
    if (await evaluate("!!__s4.button('Выполнить')")) {
      cards += 1;
      if (cards > confirmCards) break;
      await evaluate(`__s4.button('Выполнить').click(); true`);
      await sleep(400);
      continue;
    }
    if (await evaluate("!!__s4.button('Отправить')")) break;
    await sleep(150);
  }
  await sleep(400);
  const ops = (await evaluate("__s4.ops()")).slice(opsBefore);
  const log = await evaluate("document.querySelector('.log').innerText");
  return { cards, ops, requests: chatRequests - requestsBefore, log };
}

const results = [];
const record = (name, ok, details) => {
  results.push(ok);
  console.log(`${ok ? "прошла " : "ПРОВАЛ "} ${name}\n        ${details.replace(/\n/g, "\n        ")}`);
};
const cell = (address) => excel(`const r = ctx.workbook.worksheets.getItem('Заказы').getRange('${address}'); r.load('values'); await ctx.sync(); return r.values[0][0];`);

// Подготовка: объединение, в неугловую ячейку которого Excel молча не пишет.
await excel(`const s = ctx.workbook.worksheets.getItem('Заказы'); s.getRange('H20:I22').clear(); s.getRange('H21:I21').merge(); await ctx.sync();`);

// 1. Группа остановилась на полпути — следующая запись не выполняется.
if (!process.argv.includes("--only-filter")) {
  const { cards, ops, requests, log } = await task([
    ["set_ranges_values", { writes: [
      { sheet: "Заказы", address: "H20", values: [[1]] },
      { sheet: "Заказы", address: "I21", values: [[2]] }
    ] }],
    ["set_range_values", { sheet: "Заказы", address: "H22", values: [[3]] }]
  ]);
  const [h20, i21, h22] = [await cell("H20"), await cell("I21"), await cell("H22")];
  const group = ops.find((op) => op.text.startsWith("set_ranges_values"));
  const next = ops.find((op) => op.text.startsWith("set_range_values"));
  record(
    "группа остановилась на полпути — следующая запись не выполнена",
    h20 === 1 && h22 === "" && group?.status === "uncertain" && next?.status !== "done" && requests === 1,
    `H20=${JSON.stringify(h20)}, I21=${JSON.stringify(i21)}, H22=${JSON.stringify(h22)}; карточек: ${cards}; запросов к «модели»: ${requests}\n` +
    `группа: ${group?.status} — ${group?.text.split("\n").slice(0, 2).join(" ")}\n` +
    `следующая запись: ${next ? `${next.status} — ${next.text.split("\n")[0]}` : "не появилась"}\n` +
    `уведомление: ${log.split("\n").filter((line) => /повтор|проверьте/i.test(line)).slice(-1)[0] ?? "—"}`
  );
}

// 2. Фильтр на защищённом листе — отказ до карточки.
{
  await excel(`ctx.workbook.worksheets.getItem('Заказы').protection.protect({ allowAutoFilter: false }); await ctx.sync();`);
  const { cards, ops } = await task([["apply_filter", { sheet: "Заказы", address: "A1:D13", column: 0, criteria: "Москва" }]]);
  const op = ops.find((item) => item.text.startsWith("apply_filter"));
  record("фильтр на защищённом листе отклонён до карточки", cards === 0 && /защищён/.test(op?.text ?? ""), `карточек: ${cards}; итог: ${op?.status} — ${op?.text}`);
  await excel(`ctx.workbook.worksheets.getItem('Заказы').protection.unprotect(); await ctx.sync();`);
}

await excel(`const s = ctx.workbook.worksheets.getItem('Заказы'); s.getRange('H21:I21').unmerge(); s.getRange('H20:I22').clear(); await ctx.sync();`);
console.log(`\nпрошло ${results.filter(Boolean).length} из ${results.length}`);
socket.close();
process.exit(results.every(Boolean) ? 0 : 1);
