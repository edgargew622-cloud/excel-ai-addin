/**
 * Приёмка S7 в настоящем Excel: сценарии 2, 7 и 10 плана стабилизации.
 *
 * Устроена как scripts/live-excel-s4.mjs: Excel запущен с
 * WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=9229,
 * панель открыта; вместо модели отвечает скрипт, книга меняется через
 * Office.js из самой панели, кнопки нажимаются в панели.
 *
 * 2.  Заполнение, часть которого Excel молча не принял, при непустых
 *     старых ячейках: итог не должен быть полным verified.
 * 7.  Двойная отправка и отмена одновременно с отправкой: у изменения
 *     книги один владелец, панель после этого работает.
 * 10. Перезагрузка панели: беседа восстановлена, сведения исторические,
 *     записи не повторяются, режим анализа включён.
 *
 * Запуск: node scripts/live-excel-s7.mjs [2] [7] [10] — без номеров все три.
 * Сценарий 10 перезагружает панель.
 */

const PORT = 9229;
const SHEET = "Заказы";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const only = process.argv.slice(2);
const wanted = (n) => only.length === 0 || only.includes(String(n));

let socket;
let nextId = 0;
const waiting = new Map();
const listeners = new Map();
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
const requestBodies = [];
let chatRequests = 0;

async function connect() {
  const list = await (await fetch(`http://127.0.0.1:${PORT}/json/list`)).json();
  const page = list.find((target) => target.type === "page" && target.url.includes("taskpane.html"));
  if (!page) throw new Error("Панель надстройки не найдена: откройте её в Excel.");
  socket = new WebSocket(page.webSocketDebuggerUrl);
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
}

on("Fetch.requestPaused", async ({ requestId, request }) => {
  chatRequests += 1;
  requestBodies.push(request.postData ?? null);
  await send("Fetch.fulfillRequest", {
    requestId, responseCode: 200,
    responseHeaders: [{ name: "Content-Type", value: "text/event-stream" }],
    body: Buffer.from(answers.shift() ?? done).toString("base64")
  }).catch(() => undefined);
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
const cell = (address) => excel(`const r = ctx.workbook.worksheets.getItem('${SHEET}').getRange('${address}'); r.load('formulas'); await ctx.sync(); return r.formulas[0][0];`);
const column = (address) => excel(`const r = ctx.workbook.worksheets.getItem('${SHEET}').getRange('${address}'); r.load('formulas'); await ctx.sync(); return r.formulas.map((row) => row[0]);`);

/** Помощники в самой панели; после перезагрузки ставятся заново. */
async function installHelpers({ analysisOff }) {
  await evaluate(`
    window.__s7 = {
      button(text) { return [...document.querySelectorAll('button')].find((b) => b.textContent.trim().startsWith(text)); },
      type(text) {
        const area = document.querySelector('textarea');
        Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(area, text);
        area.dispatchEvent(new Event('input', { bubbles: true }));
      },
      analysisBox() { return [...document.querySelectorAll('label')].find((l) => l.textContent.includes('Только анализ'))?.querySelector('input'); },
      ops() { return [...document.querySelectorAll('.op')].map((op) => ({ status: op.className.replace('op', '').trim(), text: op.innerText })); },
      log() { return document.querySelector('.log')?.innerText ?? ''; }
    };
    true`);
  if (analysisOff) await evaluate(`(() => { const box = __s7.analysisBox(); if (box && box.checked) box.click(); return true; })()`);
  await sleep(300);
}

const idle = () => waitFor("!!__s7.button('Отправить') && !__s7.button('Выполнить')", "панель свободна", 120000);

/** Задача: ответ модели — эти вызовы; на каждую карточку — «Выполнить». */
async function task(calls, { text = "проверка S7" } = {}) {
  const opsBefore = await evaluate("__s7.ops().length");
  const requestsBefore = chatRequests;
  // Остановленная задача не забирает заключительный ответ: без очистки
  // он достался бы следующей задаче вместо её команд.
  answers.length = 0;
  answers.push(toolCalls(calls), done);
  await evaluate(`__s7.type(${JSON.stringify(text)}); true`);
  await waitFor("!!__s7.button('Отправить') && !__s7.button('Отправить').disabled", "кнопка «Отправить»");
  await evaluate(`__s7.button('Отправить').click(); true`);
  const started = Date.now() + 30000;
  while (chatRequests === requestsBefore && Date.now() < started) await sleep(100);
  await sleep(300);
  let cards = 0;
  const cardTexts = [];
  const until = Date.now() + 120000;
  while (Date.now() < until) {
    if (await evaluate("!!__s7.button('Выполнить')")) {
      cards += 1;
      cardTexts.push(await evaluate("document.querySelector('.confirm')?.innerText ?? ''"));
      await evaluate(`__s7.button('Выполнить').click(); true`);
      await sleep(400);
      continue;
    }
    if (await evaluate("!!__s7.button('Отправить')")) break;
    await sleep(150);
  }
  await sleep(400);
  const ops = (await evaluate("__s7.ops()")).slice(opsBefore);
  return { cards, cardTexts, ops, requests: chatRequests - requestsBefore };
}

const results = [];
const record = (name, ok, details) => {
  results.push(ok);
  console.log(`${ok ? "прошла " : "ПРОВАЛ "} ${name}\n        ${details.replace(/\n/g, "\n        ")}`);
};

await connect();
await send("Fetch.enable", { patterns: [{ urlPattern: "*/api/chat*" }] });
await installHelpers({ analysisOff: true });
console.log(`Панель: ${await evaluate("document.querySelector('.persistence-note')?.innerText ?? ''")}\n`);

/* --- 2. Частично проигнорированное заполнение ---------------------------------- */

if (wanted(2)) {
  // Старые значения в J40:J44; J42 — неугловая ячейка объединения I42:J42,
  // запись туда Excel молча не принимает. Старое значение объединения
  // живёт в его угле I42.
  await excel(`
    const s = ctx.workbook.worksheets.getItem('${SHEET}');
    s.getRange('I40:J44').unmerge(); s.getRange('I40:J44').clear();
    s.getRange('J40:J44').values = [[111], [111], [111], [111], [111]];
    s.getRange('I42').values = [['старое']];
    s.getRange('I42:J42').merge();
    await ctx.sync();`);
  const before = await column("J40:J44");
  const { cards, cardTexts, ops, requests } = await task([
    ["fill_range", { sheet: SHEET, address: "J40:J44", value: "=ROW()*10", isFormula: true }],
    ["set_range_values", { sheet: SHEET, address: "J46", values: [["после заполнения"]] }]
  ]);
  const after = await column("J40:J44");
  const next = await cell("J46");
  const fill = ops.find((op) => op.text.startsWith("fill_range"));
  const following = ops.find((op) => op.text.startsWith("set_range_values"));
  record(
    "2. частично проигнорированное заполнение не выдаётся за проверенное",
    fill && fill.status !== "done" && next === "" && requests === 1,
    `до: J40:J44 = ${JSON.stringify(before)}\n` +
    `после: J40:J44 = ${JSON.stringify(after)}; J46 = ${JSON.stringify(next)}\n` +
    `карточек: ${cards}; предупреждение об объединении в карточке: ${cardTexts.some((t) => /объедин/i.test(t)) ? "да" : "нет"}; запросов к «модели»: ${requests}\n` +
    `заполнение: ${fill?.status} — ${fill?.text.replace(/\n/g, " ").slice(0, 400)}\n` +
    `следующая запись: ${following ? `${following.status} — ${following.text.split("\n")[0]}` : "не появилась"}`
  );
  await excel(`const s = ctx.workbook.worksheets.getItem('${SHEET}'); s.getRange('I40:J46').unmerge(); s.getRange('I40:J46').clear(); await ctx.sync();`);
}

/* --- 7. Двойная отправка и отмена одновременно с отправкой ------------------------ */

if (wanted(7)) {
  await idle();
  // Двойная отправка: два нажатия в одном такте.
  {
    const before = chatRequests;
    answers.length = 0;
    answers.push(done, done);
    await evaluate(`__s7.type('двойная отправка'); true`);
    await waitFor("!!__s7.button('Отправить') && !__s7.button('Отправить').disabled", "кнопка «Отправить»");
    await evaluate(`(() => { const b = __s7.button('Отправить'); b.click(); b.click(); return true; })()`);
    await sleep(2500);
    await idle();
    const userMessages = await evaluate("[...document.querySelectorAll('.log .msg.user')].filter((e) => e.innerText.includes('двойная отправка')).length");
    record(
      "7а. двойная отправка — одна задача",
      chatRequests - before === 1,
      `запросов к «модели»: ${chatRequests - before}; сообщений пользователя в ленте: ${userMessages}`
    );
  }

  // Отмена и отправка в одном такте, в обоих порядках.
  for (const order of ["сначала отправка", "сначала отмена"]) {
    await idle();
    await task([["set_range_values", { sheet: SHEET, address: "J50", values: [[5]] }]]);
    await idle();
    const written = await cell("J50");
    await waitFor("!!__s7.button('Отменить') && !__s7.button('Отменить').disabled", "кнопка «Отменить»");
    const before = chatRequests;
    answers.length = 0;
    answers.push(done);
    await evaluate(`__s7.type('отмена и отправка'); true`);
    await waitFor("!!__s7.button('Отправить') && !__s7.button('Отправить').disabled", "кнопка «Отправить»");
    const logBefore = (await evaluate("__s7.log()")).length;
    await evaluate(`(() => {
      const sendButton = __s7.button('Отправить');
      const undoButton = __s7.button('Отменить');
      ${order === "сначала отправка" ? "sendButton.click(); undoButton.click();" : "undoButton.click(); sendButton.click();"}
      return true;
    })()`);
    await sleep(3000);
    await idle();
    const sent = chatRequests - before;
    const after = await cell("J50");
    const undone = after === "";
    const log = (await evaluate("__s7.log()")).slice(logBefore);
    const busyNotice = log.split("\n").find((line) => /занят|выполняется|дождитесь/i.test(line)) ?? "—";
    // Панель после этого работает: следующая задача проходит.
    const check = await task([["get_range_values", { sheet: SHEET, address: "J50" }]], { text: "панель работает?" });
    record(
      `7б. отмена и отправка в одном такте (${order}) — один владелец`,
      written === 5 && sent + (undone ? 1 : 0) === 1 && check.requests >= 1,
      `J50 записана: ${JSON.stringify(written)}; после двух нажатий: J50 = ${JSON.stringify(after)} (${undone ? "отмена выполнилась" : "отмены не было"}); запросов к «модели»: ${sent}\n` +
      `уведомление панели: ${busyNotice}\n` +
      `следующая задача: запросов ${check.requests}, ${check.ops.map((op) => `${op.status} — ${op.text.split("\n")[0]}`).join("; ") || "без операций"}`
    );
    await idle();
    if (!undone) {
      await waitFor("!!__s7.button('Отменить') && !__s7.button('Отменить').disabled", "кнопка «Отменить»", 10000).catch(() => undefined);
    }
    await excel(`ctx.workbook.worksheets.getItem('${SHEET}').getRange('J50').clear(); await ctx.sync();`);
  }
}

/* --- 10. Перезагрузка панели и восстановление беседы -------------------------------- */

if (wanted(10)) {
  await idle();
  await task([["set_range_values", { sheet: SHEET, address: "J60", values: [[7]] }]], { text: "запиши 7 в J60" });
  await idle();
  const written = await cell("J60");
  // Пользователь поправил ячейку руками: повтор записи после перезагрузки
  // вернул бы 7.
  await excel(`ctx.workbook.worksheets.getItem('${SHEET}').getRange('J60').values = [[8]]; await ctx.sync();`);
  const entriesBefore = await evaluate("__s7.log()");

  const requestsBeforeReload = chatRequests;
  await send("Page.reload", { ignoreCache: true });
  await sleep(6000);
  await installHelpers({ analysisOff: false });
  await sleep(4000);
  const note = await evaluate("document.querySelector('.persistence-note')?.innerText ?? ''");
  const analysisOn = await evaluate("!!__s7.analysisBox()?.checked");
  const entriesAfter = await evaluate("__s7.log()");
  const requestsAfterReload = chatRequests - requestsBeforeReload;
  const afterReload = await cell("J60");

  // Модель после перезагрузки снова просит записать: в режиме анализа
  // запись не должна дойти до книги, а модель — получить пометку о восстановлении.
  const bodiesBefore = requestBodies.length;
  const retry = await task([["set_range_values", { sheet: SHEET, address: "J60", values: [[7]] }]], { text: "продолжай" });
  const afterRetry = await cell("J60");
  const body = requestBodies[bodiesBefore] ?? "";
  const historical = /восстановлена после закрытия панели/.test(body);
  const writeOp = retry.ops.find((op) => op.text.startsWith("set_range_values"));

  record(
    "10. перезагрузка: беседа восстановлена, записи не повторяются, режим анализа",
    written === 7 && afterReload === 8 && afterRetry === 8 && requestsAfterReload === 0 && analysisOn && /восстановлена/.test(note) &&
      entriesAfter.includes("запиши 7 в J60") && retry.cards === 0 && (historical || body === ""),
    `до перезагрузки J60 = ${JSON.stringify(written)}, пользователь поставил 8\n` +
    `после перезагрузки: «${note}»\n` +
    `режим «Только анализ»: ${analysisOn ? "включён" : "ВЫКЛЮЧЕН"}; прежняя лента видна: ${entriesAfter.includes("запиши 7 в J60") ? "да" : "нет"} (${entriesBefore.length} → ${entriesAfter.length} символов)\n` +
    `запросов к «модели» сами по себе после перезагрузки: ${requestsAfterReload}; J60 = ${JSON.stringify(afterReload)}\n` +
    `модель снова просит записать: карточек ${retry.cards}, итог ${writeOp ? `${writeOp.status} — ${writeOp.text.replace(/\n/g, " ").slice(0, 200)}` : "операция не появилась"}; J60 = ${JSON.stringify(afterRetry)}\n` +
    `пометка «восстановлена… исторические» в запросе к модели: ${body === "" ? "тело запроса не получено" : historical ? "есть" : "НЕТ"}`
  );
  await excel(`ctx.workbook.worksheets.getItem('${SHEET}').getRange('J60').clear(); await ctx.sync();`);
}

console.log(`\nпрошло ${results.filter(Boolean).length} из ${results.length}`);
socket.close();
process.exit(results.every(Boolean) ? 0 : 1);
