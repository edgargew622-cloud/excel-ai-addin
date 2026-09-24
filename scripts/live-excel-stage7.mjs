/**
 * Живые проверки этапа 7 в настоящем Excel.
 *
 * Устроено как scripts/live-excel-s*.mjs: Excel запущен с
 * WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=9229, панель
 * открыта; вместо модели отвечает скрипт готовым вызовом, карточка
 * подтверждается в панели, итог читается из ответа инструмента и из книги.
 *
 * Запуск: node scripts/live-excel-stage7.mjs [номера сценариев…]
 */

const PORT = 9229;
const SHEET = "Э7";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const only = process.argv.slice(2);
const wanted = (id) => only.length === 0 || only.includes(id);

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
let lastBody = null;
listeners.set("Fetch.requestPaused", [async ({ requestId, request }) => {
  if (request.postData) lastBody = request.postData;
  await send("Fetch.fulfillRequest", {
    requestId, responseCode: 200,
    responseHeaders: [{ name: "Content-Type", value: "text/event-stream" }],
    body: Buffer.from(answers.shift() ?? done).toString("base64")
  }).catch(() => undefined);
}]);
await send("Fetch.enable", { patterns: [{ urlPattern: "*/api/chat*" }] });

await evaluate(`
  window.__e = {
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

/** Свежий лист с заказами: A — город, B — статус, C — сумма, D — количество. */
const resetSheet = () => excel(`
  const old = ctx.workbook.worksheets.getItemOrNullObject('${SHEET}'); old.load('isNullObject'); await ctx.sync();
  if (!old.isNullObject) { old.delete(); await ctx.sync(); }
  const s = ctx.workbook.worksheets.add('${SHEET}');
  s.getRange('A1:D7').values = [['Город','Статус','Сумма','Количество'],['Москва','Новая',900,3],['Москва','Закрыта',450,1],['Казань','Новая',700,2],['Казань','Закрыта',800,4],['Омск','Новая',1500,5],['Омск','Закрыта',1300,2]];
  s.activate();
  await ctx.sync();`);
const read = (address, props = ["formulas", "values"]) => excel(`
  const r = ctx.workbook.worksheets.getItem('${SHEET}').getRange('${address}'); r.load(${JSON.stringify(props)}); await ctx.sync();
  return Object.fromEntries(${JSON.stringify(props)}.map((p) => [p, r[p]]));`);

/** Вызов инструмента через панель: карточка, «Выполнить», ответ инструмента. */
async function run(name, args) {
  const opsBefore = await evaluate("__e.ops().length");
  lastBody = null;
  answers.length = 0;
  answers.push(toolCall(name, args), done);
  await evaluate(`__e.type('этап 7'); true`);
  await waitFor("!!__e.button('Отправить') && !__e.button('Отправить').disabled", "кнопка «Отправить»");
  await evaluate(`__e.button('Отправить').click(); true`);
  let cards = 0;
  const until = Date.now() + 90000;
  await sleep(800);
  while (Date.now() < until) {
    if (await evaluate("!!__e.button('Выполнить')")) { cards += 1; await evaluate(`__e.button('Выполнить').click(); true`); await sleep(500); continue; }
    if (await evaluate("!!__e.button('Отправить')")) break;
    await sleep(150);
  }
  await sleep(500);
  const op = (await evaluate("__e.ops()")).slice(opsBefore).find((item) => item.text.startsWith(name));
  const reply = lastBody ? JSON.parse(lastBody).messages?.findLast?.((m) => m.role === "tool")?.content ?? "" : "";
  let result = {};
  try { result = JSON.parse(reply); } catch { /* ответ не JSON */ }
  return { cards, op, reply, result, state: result.executionState ?? result.result?.executionState ?? /"executionState":"(\w+)"/.exec(reply)?.[1] ?? "?" };
}

const results = [];
const record = (name, ok, details) => {
  results.push(ok);
  console.log(`${ok ? "прошла " : "ПРОВАЛ "} ${name}\n        ${details.replace(/\n/g, "\n        ")}\n`);
};

/* --- 7.1.2: шаблон первой строки -------------------------------------------------- */

if (wanted("7.1.2")) {
  await resetSheet();
  const { cards, state, result } = await run("fill_range", { sheet: SHEET, address: "E2:G7", template: ["=C2*D2", "=E2*0.2", "=IF(B2=\"Новая\",\"в работу\",\"архив\")"] });
  const after = await read("E2:G7");
  const want = (r) => [`=C${r}*D${r}`, `=E${r}*0.2`, `=IF(B${r}="Новая","в работу","архив")`];
  const formulasOk = after.formulas.every((row, i) => JSON.stringify(row) === JSON.stringify(want(i + 2)));
  const valuesOk = after.values[0][0] === 2700 && after.values[0][1] === 540 && after.values[1][2] === "архив";
  record("7.1.2 шаблон из трёх формул протянут вниз по своим столбцам",
    cards === 1 && state === "verified" && formulasOk && valuesOk,
    `карточек: ${cards}; executionState: ${state}; проверено ячеек: ${result.checkedCells ?? result.result?.checkedCells ?? "?"}\n` +
    `E2:G2 = ${JSON.stringify(after.formulas[0])} → ${JSON.stringify(after.values[0])}\n` +
    `E7:G7 = ${JSON.stringify(after.formulas[5])} → ${JSON.stringify(after.values[5])}`);
}

/* --- 7.1.5: функции этого Excel ------------------------------------------------ */

if (wanted("7.1.5")) {
  await resetSheet();
  await excel(`ctx.workbook.worksheets.getItem('${SHEET}').getRange('F2').values = [['было']]; await ctx.sync();`);
  // TEXTSPLIT в Excel 2021 нет — по замеру.
  const first = await run("fill_range", { sheet: SHEET, address: "F2:F3", value: '=TEXTSPLIT(A2,"о")', isFormula: true });
  const target = await read("F2:F3");
  const probe = await read("F1:J1");
  record("7.1.5 недоступная функция: отказ до записи, цель и временная ячейка чисты",
    first.cards === 1 && first.state === "failed_before_write" && /нет функций TEXTSPLIT/.test(first.reply) &&
      target.formulas[0][0] === "было" && probe.formulas[0].every((cell) => cell === "" || cell === "Статус" || cell === "Количество"),
    `карточек: ${first.cards}; executionState: ${first.state}\n${first.op?.text.replace(/\s+/g, " ").slice(0, 260)}\n` +
    `F2:F3 после: ${JSON.stringify(target.formulas)}; строка 1 справа от данных: ${JSON.stringify(probe.formulas[0])}`);

  const second = await run("fill_range", { sheet: SHEET, address: "F2:F3", value: '=TEXTSPLIT(A2,"а")', isFormula: true });
  record("7.1.5 та же функция второй раз — отказ сразу, без карточки",
    second.cards === 0 && /нет функций TEXTSPLIT/.test(second.op?.text ?? ""),
    `карточек: ${second.cards}; ${second.op?.text.replace(/\s+/g, " ").slice(0, 200)}`);

  const good = await run("fill_range", { sheet: SHEET, address: "E2:E7", value: '=XLOOKUP(A2,A2:A7,C2:C7,0)', isFormula: true });
  const values = await read("E2:E7", ["values"]);
  const checked = good.result.functionsChecked ?? good.result.result?.functionsChecked;
  record("7.1.5 доступная функция проверена и записана",
    good.state === "verified" && values.values.flat().every((value) => typeof value === "number") && checked?.available?.includes("XLOOKUP"),
    `executionState: ${good.state}; functionsChecked: ${JSON.stringify(checked)}; E2:E7 = ${JSON.stringify(values.values.flat())}`);

  const russian = await run("set_range_values", { sheet: SHEET, address: "G2", values: [["=СУММ(C2:C7)"]], isFormula: true });
  record("7.1.5 русское имя функции поймано так же",
    /нет функций СУММ/.test(russian.reply || russian.op?.text || ""),
    `карточек: ${russian.cards}; executionState: ${russian.state}; ${(russian.op?.text ?? "").replace(/\s+/g, " ").slice(0, 200)}`);
}

/* --- 7.1.4: ошибки после записи на русском Excel ------------------------------- */

if (wanted("7.1.4")) {
  await resetSheet();
  await excel(`ctx.workbook.worksheets.getItem('${SHEET}').getRange('D3').values = [[0]]; await ctx.sync();`);
  const { state, result } = await run("fill_range", { sheet: SHEET, address: "E2:E7", value: "=C2/D2", isFormula: true });
  const newErrors = result.newErrors ?? result.result?.newErrors;
  const note = result.errorNote ?? result.result?.errorNote ?? "";
  record("7.1.4 #ДЕЛ/0! после записи назван",
    state === "verified" && JSON.stringify(newErrors) === JSON.stringify(["E3 #ДЕЛ/0!"]) && /деление на ноль/.test(note),
    `executionState: ${state}; newErrors: ${JSON.stringify(newErrors)}
${note}`);
}

console.log(`прошло ${results.filter(Boolean).length} из ${results.length}`);
socket.close();
process.exit(results.every(Boolean) ? 0 : 1);
