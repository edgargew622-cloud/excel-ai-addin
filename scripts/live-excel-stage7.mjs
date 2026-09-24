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
  let cardText = "";
  const until = Date.now() + 90000;
  await sleep(800);
  while (Date.now() < until) {
    if (await evaluate("!!__e.button('Выполнить')")) { cards += 1; cardText += await evaluate("document.querySelector('.preview')?.innerText ?? ''"); await evaluate(`__e.button('Выполнить').click(); true`); await sleep(500); continue; }
    if (await evaluate("!!__e.button('Отправить')")) break;
    await sleep(150);
  }
  await sleep(500);
  const op = (await evaluate("__e.ops()")).slice(opsBefore).find((item) => item.text.startsWith(name));
  const reply = lastBody ? JSON.parse(lastBody).messages?.findLast?.((m) => m.role === "tool")?.content ?? "" : "";
  let result = {};
  try { result = JSON.parse(reply); } catch { /* ответ не JSON */ }
  return { cards, cardText, op, reply, result, state: result.executionState ?? result.result?.executionState ?? /"executionState":"(\w+)"/.exec(reply)?.[1] ?? "?" };
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

/* --- 7.2.1: профиль данных --------------------------------------------------------- */

/** Лист с типичным «грязным» выгрузочным набором: пробелы, числа и даты текстом, дубликат. */
const MESSY = "Э7Ч";
const resetMessy = () => excel(`
  const old = ctx.workbook.worksheets.getItemOrNullObject('${MESSY}'); old.load('isNullObject'); await ctx.sync();
  if (!old.isNullObject) { old.delete(); await ctx.sync(); }
  const s = ctx.workbook.worksheets.add('${MESSY}');
  s.getRange('A1:D8').values = [
    ['Город', 'Сумма', 'Дата', 'Код'],
    ["' Москва", "'1 200", "'25.02.2026", "'007"],
    ['Москва', 900, "'01.02.2026", "'12"],
    ["'Казань  Север", "'1,500", 46054, "'12"],
    ["'Омск" + String.fromCharCode(160), "'1.5", '', ''],
    ['', '', '', ''],
    ['Москва', 900, "'01.02.2026", "'12"],
    ['Омск', "'2 300,50", "'2026-03-05", "'45"]];
  s.getRange('C4').numberFormat = [['dd.mm.yyyy']];
  s.activate();
  await ctx.sync();`);

if (wanted("7.2.1")) {
  await resetMessy();
  const { cards, result } = await run("profile_range", { sheet: MESSY });
  const profile = result.result ?? result;
  const column = (letter) => profile.columns?.find((item) => item.column === letter) ?? {};
  const count = (letter, key) => column(letter).findings?.[key]?.count ?? 0;
  const checks = {
    "пробелы по краям в A": count("A", "edgeSpaces") === 2,
    "двойной пробел в A": count("A", "innerSpaces") === 1,
    "неразрывный пробел в A": count("A", "nonBreakingSpaces") === 1,
    "числа текстом в B": count("B", "numbersAsText") === 2,
    "неоднозначное число 1,500": count("B", "ambiguousNumbersAsText") === 1,
    "чужой разделитель 1.5": count("B", "foreignNumbersAsText") === 1,
    "дата-число в C": column("C").dates === 1,
    "даты текстом": count("C", "datesAsText") === 2,
    "неоднозначные даты": count("C", "ambiguousDatesAsText") === 2,
    "код с нулями": count("D", "codesWithLeadingZeros") === 1,
    "пустая строка": profile.emptyRows === 1,
    "дубликат строки": profile.duplicateRows?.count === 1,
    "культура книги": profile.culture?.decimal === "," && profile.culture?.dateOrder === "DMY"
  };
  const failed = Object.entries(checks).filter(([, ok]) => !ok).map(([name]) => name);
  record("7.2.1 профиль находит всё, что мешает считать, и берёт разделители у книги",
    cards === 0 && failed.length === 0,
    `карточек: ${cards}; не совпало: ${failed.join(", ") || "—"}
` +
    `проверено: ${profile.checkedAddress}, incomplete: ${profile.incomplete}; культура: ${JSON.stringify(profile.culture)}
` +
    `дубликаты: ${JSON.stringify(profile.duplicateRows)}
` +
    profile.columns?.map((item) => `${item.column} «${item.header}»: ${JSON.stringify(item.findings ?? {})}`).join(String.fromCharCode(10)));
}

/* --- 7.2.2–7.2.3: пробелы, числа и даты из текста --------------------------------- */

const readMessy = (address, props = ["values", "valueTypes", "text", "numberFormat"]) => excel(`
  const r = ctx.workbook.worksheets.getItem('${MESSY}').getRange('${address}'); r.load(${JSON.stringify(props)}); await ctx.sync();
  return Object.fromEntries(${JSON.stringify(props)}.map((p) => [p, r[p]]));`);
const flat = (matrix) => matrix.map((row) => row[0]);

if (wanted("7.2.2")) {
  await resetMessy();
  const trim = await run("trim_text", { sheet: MESSY, address: "A2:A8" });
  const a = await readMessy("A2:A8");
  const d = await run("trim_text", { sheet: MESSY, address: "D2:D8" });
  record("7.2.2 лишние пробелы убраны, текст остался текстом",
    trim.cards === 1 && trim.state === "verified" &&
      JSON.stringify(flat(a.values)) === JSON.stringify(["Москва", "Москва", "Казань Север", "Омск", "", "Москва", "Омск"]) &&
      flat(a.valueTypes).every((type) => type === "String" || type === "Empty") &&
      /менять нечего/.test(d.op?.text ?? ""),
    `карточек: ${trim.cards}; executionState: ${trim.state}; изменено: ${trim.result.result?.changedCells ?? trim.result.changedCells}
` +
    `A2:A8 = ${JSON.stringify(flat(a.values))}; типы: ${JSON.stringify(flat(a.valueTypes))}
` +
    `коды D2:D8 без пробелов: ${(d.op?.text ?? "").replace(/\s+/g, " ").slice(0, 120)}`);
}

if (wanted("7.2.3")) {
  await resetMessy();
  const numbers = await run("convert_values", { sheet: MESSY, address: "B2:B8", to: "number" });
  const b = await readMessy("B2:B8");
  const skipped = JSON.stringify(numbers.result.result?.skipped ?? numbers.result.skipped ?? {});
  record("7.2.3 числа из текста — только однозначные, по разделителям книги",
    numbers.state === "verified" && b.values[0][0] === 1200 && b.values[6][0] === 2300.5 && b.values[2][0] === "1,500" && b.values[3][0] === "1.5" &&
      /дробь или тысячи/.test(skipped) && /не как в книге/.test(skipped),
    `B2:B8 = ${JSON.stringify(flat(b.values))}; типы: ${JSON.stringify(flat(b.valueTypes))}
пропущено: ${skipped}`);

  const explicit = await run("convert_values", { sheet: MESSY, address: "B4:B5", to: "number", decimalSeparator: "." });
  const b2 = await readMessy("B4:B5");
  record("7.2.3 названный пользователем разделитель: 1,500 → 1500, 1.5 → 1,5",
    explicit.state === "verified" && b2.values[0][0] === 1500 && b2.values[1][0] === 1.5,
    `B4:B5 = ${JSON.stringify(flat(b2.values))}`);

  const dates = await run("convert_values", { sheet: MESSY, address: "C2:C8", to: "date" });
  const c = await readMessy("C2:C8");
  record("7.2.3 даты из текста — однозначные получают дату и формат книги, неоднозначные ждут",
    dates.state === "verified" && c.values[0][0] === 46078 && c.text[0][0] === "25.02.2026" && c.numberFormat[0][0] === "dd.mm.yyyy" &&
      c.values[1][0] === "01.02.2026" && c.text[6][0] === "05.03.2026",
    `C2:C8 значения ${JSON.stringify(flat(c.values))}
вид ${JSON.stringify(flat(c.text))}; формат C2: ${c.numberFormat[0][0]}`);

  const ordered = await run("convert_values", { sheet: MESSY, address: "C3", to: "date", dateOrder: "DMY" });
  const c3 = await readMessy("C3");
  record("7.2.3 названный порядок DMY: 01.02.2026 → 1 февраля",
    ordered.state === "verified" && c3.values[0][0] === 46054 && c3.text[0][0] === "01.02.2026",
    `C3 = ${c3.values[0][0]} («${c3.text[0][0]}»)`);

  // Отмена последнего — даты C3: вернуться должны текст и формат.
  await waitFor("!!__e.button('Отменить') && !__e.button('Отменить').disabled", "кнопка «Отменить»", 20000);
  await evaluate(`__e.button('Отменить').click(); true`);
  await sleep(2500);
  const undone = await readMessy("C3");
  record("7.2.3 отмена вернула текст и прежний формат",
    undone.values[0][0] === "01.02.2026" && undone.valueTypes[0][0] === "String" && undone.numberFormat[0][0] === "General",
    `C3 после отмены: ${JSON.stringify(undone.values[0][0])} (${undone.valueTypes[0][0]}), формат ${undone.numberFormat[0][0]}`);
}

/* --- 7.2.4: дубликаты ------------------------------------------------------------------ */

if (wanted("7.2.4")) {
  const DUP = "Э7Д";
  const resetDup = (withNeighbour) => excel(`
    const old = ctx.workbook.worksheets.getItemOrNullObject('${DUP}'); old.load('isNullObject'); await ctx.sync();
    if (!old.isNullObject) { old.delete(); await ctx.sync(); }
    const s = ctx.workbook.worksheets.add('${DUP}');
    s.getRange('A1:C12').values = [['Город','Сумма','Номер'],['Москва',100,1],['москва',100,2],['Москва ',100,3],["'1",5,4],[1,5,5],['Омск',200,6],['Омск',200,7],['','',8],['','',9],['Казань',300,10],['Омск',200,11]];
    s.getRange('F2:F3').formulas = [['=C8'],['=SUM(C2:C12)']];
    ${withNeighbour ? "s.getRange('D5').values = [['заметка']];" : ""}
    s.activate();
    await ctx.sync();`);
  await resetDup(false);
  const res = await run("remove_duplicates", { sheet: DUP, address: "A1:C12", columns: ["Город", "Сумма"] });
  const after = await excel(`const r = ctx.workbook.worksheets.getItem('${DUP}').getRange('A2:C12'); r.load('values'); await ctx.sync(); return r.values;`);
  const body = res.result.result ?? res.result;
  const risks = JSON.stringify(body.affectedFormulas ?? []);
  record("7.2.4 дубликаты удалены так, как рассчитала панель; задетые формулы названы",
    res.cards === 1 && res.state === "verified" && body.removedRows === 4 &&
      JSON.stringify(after.map((row) => row[2])) === JSON.stringify([1, 3, 4, 5, 6, 8, 10, "", "", "", ""]) &&
      /F2/.test(risks) && !/F3/.test(risks),
    `карточек: ${res.cards}; executionState: ${res.state}; удалено: ${body.removedRows}, осталось: ${body.remainingRows}` + String.fromCharCode(10) +
    `номера после: ${JSON.stringify(after.map((row) => row[2]))}` + String.fromCharCode(10) +
    `задетые формулы: ${risks}`);

  await resetDup(true);
  const refused = await run("remove_duplicates", { sheet: DUP, address: "A1:C12", columns: ["Город", "Сумма"] });
  record("7.2.4 данные вплотную к таблице — отказ до карточки",
    refused.cards === 0 && /разъехались/.test(refused.op?.text ?? ""),
    `карточек: ${refused.cards}; ${(refused.op?.text ?? "").replace(/\s+/g, " ").slice(0, 220)}`);
}

/* --- 7.2.5: результат на отдельный лист ---------------------------------------------------- */

if (wanted("7.2.5")) {
  const SRC = "Э7Д";
  const OUT = "Э7Итог";
  await excel(`
    for (const name of ['${SRC}', '${OUT}']) {
      const old = ctx.workbook.worksheets.getItemOrNullObject(name); old.load('isNullObject'); await ctx.sync();
      if (!old.isNullObject) { old.delete(); await ctx.sync(); }
    }
    const s = ctx.workbook.worksheets.add('${SRC}');
    s.getRange('A1:C12').values = [['Город','Сумма','Номер'],['Москва',100,1],['москва',100,2],['Москва ',100,3],["'1",5,4],[1,5,5],['Омск',200,6],['Омск',200,7],['','',8],['','',9],['Казань',300,10],['Омск',200,11]];
    s.getRange('D5').values = [['заметка рядом']];
    ctx.workbook.worksheets.add('${OUT}');
    s.activate();
    await ctx.sync();`);
  const before = await excel(`const r = ctx.workbook.worksheets.getItem('${SRC}').getRange('A1:D12'); r.load('values'); await ctx.sync(); return JSON.stringify(r.values);`);
  const res = await run("remove_duplicates", { sheet: SRC, address: "A1:C12", columns: ["Город", "Сумма"], destSheet: OUT });
  const after = await excel(`const r = ctx.workbook.worksheets.getItem('${SRC}').getRange('A1:D12'); r.load('values'); await ctx.sync(); return JSON.stringify(r.values);`);
  const copy = await excel(`const r = ctx.workbook.worksheets.getItem('${OUT}').getRange('A1:C9'); r.load(['values','valueTypes']); await ctx.sync(); return { values: r.values, types: r.valueTypes };`);
  record("7.2.5 уникальные строки — на отдельный лист, источник не тронут",
    res.state === "verified" && before === after && copy.values.length === 9 &&
      JSON.stringify(copy.values.map((row) => row[2])) === JSON.stringify(["Номер", 1, 3, 4, 5, 6, 8, 10, ""]) && copy.types[3][0] === "String",
    `executionState: ${res.state}; источник не изменился: ${before === after}` + String.fromCharCode(10) +
    `копия, столбец «Номер»: ${JSON.stringify(copy.values.map((row) => row[2]))}; «1» осталась текстом: ${copy.types[3][0] === "String"}`);

  await waitFor("!!__e.button('Отменить') && !__e.button('Отменить').disabled", "кнопка «Отменить»", 20000);
  await evaluate(`__e.button('Отменить').click(); true`);
  await sleep(2500);
  const cleared = await excel(`const u = ctx.workbook.worksheets.getItem('${OUT}').getUsedRangeOrNullObject(true); u.load('isNullObject'); await ctx.sync(); return u.isNullObject;`);
  record("7.2.5 отмена убрала копию", cleared === true, `лист «${OUT}» пуст после отмены: ${cleared}`);
}

/* --- 7.3.1: переименование и удаление листа ------------------------------------------ */

if (wanted("7.3.1")) {
  const NL = String.fromCharCode(10);
  const setup = () => excel(`
    for (const name of ['Э7Лист', 'Э7Лист новый', 'Э7Ссылки']) {
      const old = ctx.workbook.worksheets.getItemOrNullObject(name); old.load('isNullObject'); await ctx.sync();
      if (!old.isNullObject) { old.delete(); await ctx.sync(); }
    }
    const nm = ctx.workbook.names.getItemOrNullObject('Э7Ставка'); nm.load('isNullObject'); await ctx.sync(); if (!nm.isNullObject) { nm.delete(); await ctx.sync(); }
    const a = ctx.workbook.worksheets.add('Э7Лист');
    a.getRange('A1:A2').values = [[10], [20]];
    const b = ctx.workbook.worksheets.add('Э7Ссылки');
    b.getRange('A1:A3').formulas = [['=Э7Лист!A1*2'], ['=INDIRECT("Э7Лист!A1")'], ['=Э7Ставка+1']];
    b.getRange('A4').values = [['см. лист Э7Лист']];
    ctx.workbook.names.add('Э7Ставка', a.getRange('A2'));
    a.activate();
    await ctx.sync();`);
  const refs = () => excel(`const r = ctx.workbook.worksheets.getItem('Э7Ссылки').getRange('A1:A4'); r.load(['formulas','values']); await ctx.sync(); return r.formulas.map((row, i) => row[0] + ' → ' + r.values[i][0]);`);

  await setup();
  const rename = await run("rename_sheet", { sheet: "Э7Лист", newName: "Э7Лист новый" });
  const afterRename = await refs();
  const body = rename.result.result ?? rename.result;
  record("7.3.1 переименование: ссылки переписаны, значения те же, INDIRECT назван",
    rename.cards === 1 && rename.state === "verified" && afterRename[0] === "='Э7Лист новый'!A1*2 → 20" && afterRename[2] === "=Э7Ставка+1 → 21" &&
      (body.brokenLiteralFormulas ?? []).some((item) => item.cell === "A2"),
    `executionState: ${rename.state}` + NL + afterRename.join(NL) + NL + `названо как сломанное: ${JSON.stringify(body.brokenLiteralFormulas ?? [])}`);

  await waitFor("!!__e.button('Отменить') && !__e.button('Отменить').disabled", "кнопка «Отменить»", 20000);
  await evaluate(`__e.button('Отменить').click(); true`);
  await sleep(2500);
  const back = await excel(`const s = ctx.workbook.worksheets.getItemOrNullObject('Э7Лист'); s.load('isNullObject'); await ctx.sync(); return !s.isNullObject;`);
  const afterUndo = await refs();
  record("7.3.1 отмена вернула прежнее имя", back && afterUndo[0] === "=Э7Лист!A1*2 → 20", afterUndo.join(NL));

  const del = await run("delete_sheet", { sheet: "Э7Лист" });
  const afterDelete = await refs();
  const dbody = del.result.result ?? del.result;
  record("7.3.1 удаление: лист исчез, сломанные формулы предсказаны и названы",
    del.cards === 1 && del.state === "verified" && afterDelete.slice(0, 3).every((line) => /#ССЫЛКА!|#REF!/.test(line)) &&
      (dbody.brokenFormulas ?? []).length === 3 && JSON.stringify(dbody.brokenNames) === JSON.stringify(["Э7Ставка"]),
    `executionState: ${del.state}; ошибок ссылок до/после: ${dbody.refErrorsBefore} → ${dbody.refErrorsAfter}` + NL + afterDelete.join(NL) + NL +
    `названо: ${JSON.stringify((dbody.brokenFormulas ?? []).map((item) => item.cell))}, имена: ${JSON.stringify(dbody.brokenNames)}`);
  await excel(`const b = ctx.workbook.worksheets.getItemOrNullObject('Э7Ссылки'); b.load('isNullObject'); await ctx.sync(); if (!b.isNullObject) b.delete(); const n = ctx.workbook.names.getItemOrNullObject('Э7Ставка'); n.load('isNullObject'); await ctx.sync(); if (!n.isNullObject) n.delete(); await ctx.sync();`);
}

/* --- 7.3.2: столбцы ------------------------------------------------------------------------ */

if (wanted("7.3.2")) {
  const NL = String.fromCharCode(10);
  const COLS = "Э7С";
  await excel(`
    const old = ctx.workbook.worksheets.getItemOrNullObject('${COLS}'); old.load('isNullObject'); await ctx.sync(); if (!old.isNullObject) { old.delete(); await ctx.sync(); }
    const s = ctx.workbook.worksheets.add('${COLS}');
    s.getRange('A1:E3').values = [['a','b','c','d','e'],[1,2,3,4,5],[1,2,3,4,5]];
    s.getRange('H1:H6').formulas = [['=SUM(C:C)'], ['=SUM(B:D)'], ['=SUM(B2:C2)'], ['=C2*10'], ['=SUM(A2:E2)'], ['=SUM(2:2)']];
    s.activate();
    await ctx.sync();`);
  const formulasNow = () => excel(`const r = ctx.workbook.worksheets.getItem('${COLS}').getRange('A1:J6'); r.load(['formulas','values']); await ctx.sync();
    const out = []; r.formulas.forEach((row, i) => row.forEach((f, j) => { if (typeof f === 'string' && f.startsWith('=')) out.push(String.fromCharCode(65 + j) + (i + 1) + ' ' + f + ' → ' + r.values[i][j]); })); return out;`);
  const del = await run("delete_columns", { sheet: COLS, startColumn: "C", count: 1 });
  const body = del.result.result ?? del.result;
  const kinds = (body.affectedFormulas ?? []).map((risk) => `${risk.address}:${risk.kind}`).sort();
  const after = await formulasNow();
  record("7.3.2 удаление столбца C: сломанные и суженные формулы предсказаны, ошибки сверены",
    del.cards === 1 && del.state === "verified" &&
      JSON.stringify(kinds) === JSON.stringify(["H1:broken", "H2:shrunk", "H3:shrunk", "H4:broken", "H5:shrunk"]) && body.refErrorsAfter - body.refErrorsBefore === 2,
    `executionState: ${del.state}; ошибок ссылок ${body.refErrorsBefore} → ${body.refErrorsAfter}` + NL + `предсказано: ${JSON.stringify(kinds)}` + NL + after.join(NL));

  const ins = await run("insert_columns", { sheet: COLS, startColumn: "E", count: 1 });
  const ibody = ins.result.result ?? ins.result;
  const missed = (ibody.affectedFormulas ?? []).map((risk) => `${risk.address} ${risk.formula}:${risk.kind}`);
  record("7.3.2 вставка вплотную за данными: итог, который не охватит новый столбец, назван",
    ins.cards === 1 && ins.state === "verified" && missed.some((item) => /SUM\(A2:D2\):missed/.test(item)),
    `executionState: ${ins.state}; названо: ${JSON.stringify(missed)}`);

  await excel(`const s = ctx.workbook.worksheets.getItem('${COLS}'); s.getRange('A10:C12').values = [['x','y','z'],[1,2,3],[4,5,6]]; s.tables.add('${COLS}!A10:C12', true); await ctx.sync();`);
  const refused = await run("delete_columns", { sheet: COLS, startColumn: "B", count: 1 });
  record("7.3.2 столбец через таблицу Excel — отказ до карточки",
    refused.cards === 0 && /таблицу Excel/.test(refused.op?.text ?? ""),
    (refused.op?.text ?? "").replace(/\s+/g, " ").slice(0, 200));
}

/* --- 7.3.3: группировка --------------------------------------------------------------------- */

if (wanted("7.3.3")) {
  const G = "Э7Г";
  await excel(`
    const old = ctx.workbook.worksheets.getItemOrNullObject('${G}'); old.load('isNullObject'); await ctx.sync(); if (!old.isNullObject) { old.delete(); await ctx.sync(); }
    const s = ctx.workbook.worksheets.add('${G}');
    s.getRange('A1:E8').values = [[1,2,3,4,5],[1,2,3,4,5],[1,2,3,4,5],[1,2,3,4,5],[1,2,3,4,5],[1,2,3,4,5],[1,2,3,4,5],[1,2,3,4,5]];
    s.getRange('4:4').rowHidden = true;
    s.activate();
    await ctx.sync();`);
  const rowsHidden = () => excel(`const s = ctx.workbook.worksheets.getItem('${G}'); const out = []; for (const r of [2,3,4,5,6]) { const x = s.getRange(r + ':' + r); x.load('rowHidden'); out.push(x); } await ctx.sync(); return out.map((x) => x.rowHidden);`);
  const rows = await run("group_rows_columns", { sheet: G, address: "3:5" });
  const afterRows = await rowsHidden();
  record("7.3.3 группа строк доказана свёрткой, видимость прежняя (строка 4 так и скрыта)",
    rows.cards === 1 && rows.state === "verified" && JSON.stringify(afterRows) === JSON.stringify([false, false, true, false, false]),
    `executionState: ${rows.state}; скрыты строки 2–6: ${JSON.stringify(afterRows)}`);

  const cols = await run("group_rows_columns", { sheet: G, address: "C:D", collapse: true });
  const colHidden = () => excel(`const s = ctx.workbook.worksheets.getItem('${G}'); const out = []; for (const c of ['B','C','D','E']) { const x = s.getRange(c + ':' + c); x.load('columnHidden'); out.push(x); } await ctx.sync(); return out.map((x) => x.columnHidden);`);
  const folded = await colHidden();
  record("7.3.3 группа столбцов свёрнута по просьбе", cols.state === "verified" && JSON.stringify(folded) === JSON.stringify([false, true, true, false]),
    `executionState: ${cols.state}; скрыты B–E: ${JSON.stringify(folded)}`);

  await waitFor("!!__e.button('Отменить') && !__e.button('Отменить').disabled", "кнопка «Отменить»", 20000);
  await evaluate(`__e.button('Отменить').click(); true`);
  await sleep(2500);
  const unfolded = await colHidden();
  record("7.3.3 отмена сняла группу столбцов и вернула видимость", JSON.stringify(unfolded) === JSON.stringify([false, false, false, false]),
    `скрыты B–E после отмены: ${JSON.stringify(unfolded)}`);
}

/* --- 7.3.4: сводная на новом листе ------------------------------------------------------ */

if (wanted("7.3.4")) {
  const OUT = "Э7 Сводная";
  await resetSheet();
  await excel(`const old = ctx.workbook.worksheets.getItemOrNullObject('${OUT}'); old.load('isNullObject'); await ctx.sync(); if (!old.isNullObject) { old.delete(); await ctx.sync(); }`);
  const res = await run("create_pivot_table", { sheet: SHEET, sourceAddress: "A1:D7", rows: ["Город"], values: [{ field: "Сумма" }], newSheet: OUT });
  const pivot = await excel(`const s = ctx.workbook.worksheets.getItemOrNullObject('${OUT}'); s.load('isNullObject'); await ctx.sync(); if (s.isNullObject) return null;
    const p = s.pivotTables; p.load('items/name'); await ctx.sync(); if (!p.items.length) return 'пусто';
    const r = p.items[0].layout.getRange(); r.load(['address','values']); await ctx.sync(); return { address: r.address, values: r.values };`);
  record("7.3.4 сводная на новом листе одной операцией",
    res.cards === 1 && res.state === "verified" && pivot && /A1:B5$/.test(pivot.address) && JSON.stringify(pivot.values.at(-1)) === JSON.stringify(["Общий итог", 5650]),
    `executionState: ${res.state}; ${JSON.stringify(pivot)}`);

  await waitFor("!!__e.button('Отменить') && !__e.button('Отменить').disabled", "кнопка «Отменить»", 20000);
  await evaluate(`__e.button('Отменить').click(); true`);
  await sleep(2500);
  const gone = await excel(`const s = ctx.workbook.worksheets.getItemOrNullObject('${OUT}'); s.load('isNullObject'); await ctx.sync(); return s.isNullObject;`);
  record("7.3.4 отмена убрала сводную и пустой лист", gone === true, `лист «${OUT}» удалён: ${gone}`);
}

/* --- 7.4.2: правила проверки ввода ------------------------------------------------------- */

if (wanted("7.4.2")) {
  const V = "Э7П";
  await excel(`
    const old = ctx.workbook.worksheets.getItemOrNullObject('${V}'); old.load('isNullObject'); await ctx.sync(); if (!old.isNullObject) { old.delete(); await ctx.sync(); }
    const s = ctx.workbook.worksheets.add('${V}');
    s.getRange('A1:C6').values = [['Статус','Сумма','Дата'],['Новая',100,46054],['Закрыта',-5,46100],['Отменена',300,40000],['в работе',0,''],['','abc',46200]];
    s.activate();
    await ctx.sync();`);
  const body = (res) => res.result.result ?? res.result;
  const list = await run("set_data_validation", { sheet: V, address: "A2:A6", rule: "list", items: ["Новая", "В работе", "Закрыта"] });
  record("7.4.2 список: правило стоит, нарушители названы самим Excel",
    list.cards === 1 && list.state === "verified" && JSON.stringify(body(list).invalidExamples) === JSON.stringify(["A4", "A5"]),
    `executionState: ${list.state}; нарушители: ${JSON.stringify(body(list).invalidExamples)}`);
  const positive = await run("set_data_validation", { sheet: V, address: "B2:B6", rule: "decimal", operator: "greaterThan", value: 0 });
  record("7.4.2 число > 0: нарушители −5, 0, «abc»",
    positive.state === "verified" && JSON.stringify(body(positive).invalidExamples) === JSON.stringify(["B3", "B5", "B6"]),
    `нарушители: ${JSON.stringify(body(positive).invalidExamples)}`);
  const dates = await run("set_data_validation", { sheet: V, address: "C2:C6", rule: "date", operator: "between", value: "2026-01-01", value2: "2026-12-31" });
  record("7.4.2 дата в 2026 году: сверена по смыслу, нарушитель C4",
    dates.state === "verified" && JSON.stringify(body(dates).invalidExamples) === JSON.stringify(["C4"]),
    `executionState: ${dates.state}; нарушители: ${JSON.stringify(body(dates).invalidExamples)}`);
  await waitFor("!!__e.button('Отменить') && !__e.button('Отменить').disabled", "кнопка «Отменить»", 20000);
  await evaluate(`__e.button('Отменить').click(); true`);
  await sleep(2500);
  const after = await excel(`const r = ctx.workbook.worksheets.getItem('${V}').getRange('C2:C6'); r.dataValidation.load('type'); await ctx.sync(); return r.dataValidation.type;`);
  record("7.4.2 отмена сняла правило дат", after === "None", `тип правила после отмены: ${after}`);
}

if (wanted("7.4.1")) {
  const T = "Э7Т", O = "Э7Т2";
  await excel(`
    try { ctx.workbook.names.getItem('ИмяЭ7Т').delete(); await ctx.sync(); } catch (e) { }
    for (const n of ['${T}', '${O}']) { const old = ctx.workbook.worksheets.getItemOrNullObject(n); old.load('isNullObject'); await ctx.sync(); if (!old.isNullObject) { old.delete(); await ctx.sync(); } }
    const s = ctx.workbook.worksheets.add('${T}'); const o = ctx.workbook.worksheets.add('${O}');
    s.getRange('A1:B5').values = [['Город','Сумма'],['Москва',100],['Казань',200],['Омск',300],['Тула',400]];
    const t = s.tables.add('${T}!A1:B5', true); t.name = 'ТЭ7'; t.style = 'TableStyleMedium2'; t.showTotals = true;
    await ctx.sync();
    t.columns.getItem('Город').filter.applyValuesFilter(['Москва','Омск']);
    o.getRange('A1:A2').formulas = [['=SUM(ТЭ7[Сумма])'], ['=INDIRECT("ТЭ7[Сумма]")']];
    ctx.workbook.names.add('ИмяЭ7Т', '=ТЭ7[Сумма]');
    s.activate();
    await ctx.sync();`);
  const body = (res) => res.result.result ?? res.result;
  const conv = await run("convert_table_to_range", { sheet: T, address: "B3" });
  const b = body(conv);
  record("7.4.1 таблица стала диапазоном, сверка прошла",
    conv.cards === 1 && conv.state === "verified",
    `executionState: ${conv.state}; ${JSON.stringify(b).slice(0, 300)}`);
  record("7.4.1 названы сломанная INDIRECT и пересчёт итогов после снятия фильтра",
    JSON.stringify(b.brokenFormulas) === JSON.stringify([`${O}!A2`]) && /B6: 400 → 1000/.test(b.filterNote ?? ""),
    `brokenFormulas: ${JSON.stringify(b.brokenFormulas)}; filterNote: ${b.filterNote}`);
  const after = await excel(`
    const s = ctx.workbook.worksheets.getItem('${T}'); const tl = s.tables; tl.load('items'); const h = s.getRange('A1'); h.format.fill.load('color');
    const f = ctx.workbook.worksheets.getItem('${O}').getRange('A1'); f.load(['formulas','values']); const n = ctx.workbook.names.getItem('ИмяЭ7Т'); n.load('formula');
    await ctx.sync();
    return { tables: tl.items.length, header: h.format.fill.color, other: f.formulas[0][0] + ' → ' + f.values[0][0], name: n.formula };`);
  record("7.4.1 в книге: таблиц нет, стиль остался на шапке, ссылки переписаны",
    after.tables === 0 && after.header === "#4F81BD" && after.other === `=SUM(${T}!$B$2:$B$5) → 1000` && after.name === `=${T}!$B$2:$B$5`,
    JSON.stringify(after));
  const card = conv.cardText ?? "";
  record("7.4.1 карточка предупреждала о стиле и фильтре", /останется на ячейках/.test(card) && /фильтр/.test(card), card.slice(0, 300));
  await excel(`try { ctx.workbook.names.getItem('ИмяЭ7Т').delete(); await ctx.sync(); } catch (e) { }`);
}

const cfRules = (sheet, address) => excel(`
  const col = ctx.workbook.worksheets.getItem('${sheet}').getRange('${address}').conditionalFormats; col.load('items/type,items/priority'); await ctx.sync();
  const out = col.items.map((i) => { const r = i.getRange(); r.load('address'); let f = null; if (i.type === 'CellValue') { f = i.cellValue.format.fill; f.load('color'); } else if (i.type === 'Custom') { f = i.custom.format.fill; f.load('color'); } return { i, r, f }; });
  await ctx.sync();
  return out.map(({ i, r, f }) => i.type + ' ' + r.address.replace(/^.*!/, '') + ' ' + (f ? f.color : ''));`);

if (wanted("7.4.3-undo")) {
  // Дефект: ID правила — номер в списке. Правило, добавленное после операции
  // агента наверх, сдвигает номера, и отмена удаляла чужое правило.
  const U = "Э7У";
  await excel(`
    const old = ctx.workbook.worksheets.getItemOrNullObject('${U}'); old.load('isNullObject'); await ctx.sync(); if (!old.isNullObject) { old.delete(); await ctx.sync(); }
    const s = ctx.workbook.worksheets.add('${U}');
    s.getRange('A1:C5').values = [['Город','Статус','Сумма'],['Москва','Новая',900],['Казань','Закрыта',450],['Омск','Новая',1500],['Тула','Закрыта',300]];
    const r1 = s.getRange('C2:C5').conditionalFormats.add('CellValue'); r1.cellValue.format.fill.color = '#FFFF00'; r1.cellValue.rule = { formula1: '0', operator: 'GreaterThan' };
    s.activate();
    await ctx.sync();`);
  const added = await run("add_conditional_format", { sheet: U, address: "C2:C5", rule: "greaterThan", value: 1000, fillColor: "#FF0000" });
  const afterAdd = await cfRules(U, "C2:C5");
  // «Пользователь» добавляет своё правило и ставит его наверх, как делает Excel.
  await excel(`
    const c = ctx.workbook.worksheets.getItem('${U}').getRange('C2:C5').conditionalFormats.add('Custom'); c.custom.rule.formula = '=$B2="Новая"'; c.custom.format.fill.color = '#00B050';
    await ctx.sync(); c.priority = 0; await ctx.sync();`);
  await waitFor("!!__e.button('Отменить') && !__e.button('Отменить').disabled", "кнопка «Отменить»", 20000);
  await evaluate(`__e.button('Отменить').click(); true`);
  await sleep(2500);
  const left = await cfRules(U, "C2:C5");
  const ok = left.length === 2 && left.some((rule) => rule.includes("#00B050")) && left.some((rule) => rule.includes("#FFFF00"));
  record("7.4.3 отмена удалила именно правило агента, а не сдвинутое чужое",
    added.state === "verified" && ok,
    `после добавления: ${JSON.stringify(afterAdd)}
после отмены: ${JSON.stringify(left)}`);
}

if (wanted("7.4.3")) {
  const U = "Э7У";
  const reset = () => excel(`
    const old = ctx.workbook.worksheets.getItemOrNullObject('${U}'); old.load('isNullObject'); await ctx.sync(); if (!old.isNullObject) { old.delete(); await ctx.sync(); }
    const s = ctx.workbook.worksheets.add('${U}');
    s.getRange('A1:C5').values = [['Город','Статус','Сумма'],['Москва','Новая',900],['Казань','Закрыта',450],['Омск','Новая',1500],['Тула','Закрыта',300]];
    const r1 = s.getRange('C2:C5').conditionalFormats.add('CellValue'); r1.cellValue.format.fill.color = '#FFFF00'; r1.cellValue.rule = { formula1: '0', operator: 'GreaterThan' };
    s.activate();
    await ctx.sync();`);
  const body = (res) => res.result.result ?? res.result;
  await reset();
  const row = await run("add_conditional_format", { sheet: U, address: "A2:C5", rule: "formula", formula: '=$B2="Новая"', fillColor: "#FFC7CE" });
  const afterRow = await cfRules(U, "A2:C5");
  record("7.4.3 правило по формуле: сверено, встало выше прежнего",
    row.cards === 1 && row.state === "verified" && body(row).priority === 1 && afterRow[0] === "Custom A2:C5 #FFC7CE",
    `executionState: ${row.state}; priority: ${body(row).priority}; правила: ${JSON.stringify(afterRow)}`);

  const unknown = await run("add_conditional_format", { sheet: U, address: "C2:C5", rule: "formula", formula: "=NOSUCHFN($C2)>500", fillColor: "#00FF00" });
  const afterUnknown = await cfRules(U, "A2:C5");
  record("7.4.3 неизвестная функция в условии: отказ до записи, правил не прибавилось",
    unknown.state === "failed_before_write" && /NOSUCHFN/.test(unknown.reply) && afterUnknown.length === 2,
    `executionState: ${unknown.state}; правила: ${JSON.stringify(afterUnknown)}; ${unknown.reply.slice(0, 200)}`);

  const broken = await run("add_conditional_format", { sheet: U, address: "C2:C5", rule: "formula", formula: "=$C2>>5", fillColor: "#00FF00" });
  const afterBroken = await cfRules(U, "A2:C5");
  record("7.4.3 формула с ошибкой: Excel отказал, пустое правило убрано",
    broken.state === "failed_before_write" && afterBroken.length === 2,
    `executionState: ${broken.state}; правила: ${JSON.stringify(afterBroken)}; ${broken.reply.slice(0, 220)}`);

  const listed = await run("get_conditional_formats", { sheet: U, address: "C2:C5" });
  const rules = body(listed).rules ?? [];
  record("7.4.3 список правил по приоритету",
    rules.length === 2 && /формула/.test(rules[0].rule) && /значение больше 0/.test(rules[1].rule),
    JSON.stringify(rules));

  const moved = await run("move_conditional_format", { sheet: U, address: "C2:C5", position: 2, to: "first" });
  const afterMove = await cfRules(U, "A2:C5");
  record("7.4.3 перенос правила наверх: сверен порядок",
    moved.cards === 1 && moved.state === "verified" && afterMove[0] === "CellValue C2:C5 #FFFF00",
    `executionState: ${moved.state}; правила: ${JSON.stringify(afterMove)}`);
  await waitFor("!!__e.button('Отменить') && !__e.button('Отменить').disabled", "кнопка «Отменить»", 20000);
  await evaluate(`__e.button('Отменить').click(); true`);
  await sleep(2500);
  const afterUndo = await cfRules(U, "A2:C5");
  record("7.4.3 отмена переноса вернула прежний порядок",
    afterUndo[0] === "Custom A2:C5 #FFC7CE" && afterUndo[1] === "CellValue C2:C5 #FFFF00",
    JSON.stringify(afterUndo));
}

if (wanted("7.4.4")) {
  const F = "Э7Ф", D = "Э7Допущения";
  await excel(`
    for (const n of ['${F}', '${D}']) { const old = ctx.workbook.worksheets.getItemOrNullObject(n); old.load('isNullObject'); await ctx.sync(); if (!old.isNullObject) { old.delete(); await ctx.sync(); } }
    const d = ctx.workbook.worksheets.add('${D}'); d.getRange('A1:B1').values = [['Рост', 0.1]];
    const s = ctx.workbook.worksheets.add('${F}');
    s.getRange('A1:C5').formulas = [
      ['Показатель', '2025', '2026'],
      ['Выручка', 1000, '=B2*(1+${D}!B1)'],
      ['Затраты', 600, '=B3*1.05'],
      ['Прибыль', '=B2-B3', '=C2-C3'],
      ['Проверка', '=B4-(B2-B3)', '=C4-(C2-C3)']];
    s.getRange('B3').format.font.color = '#FF0000';
    const r = s.getRange('B5:C5').conditionalFormats.add('CellValue'); r.cellValue.format.font.color = '#FF00FF'; r.cellValue.rule = { formula1: '0', operator: 'NotEqualTo' };
    s.activate();
    await ctx.sync();`);
  const body = (res) => res.result.result ?? res.result;
  const res = await run("apply_color_convention", { sheet: F, address: "A1:C5", checks: "B5:C5" });
  const b = body(res);
  const colors = await excel(`
    const s = ctx.workbook.worksheets.getItem('${F}');
    const cells = ['B1','A2','B2','B3','C2','C3','B4','B5'].map((a) => { const c = s.getRange(a); c.format.font.load('color'); return [a, c]; });
    await ctx.sync();
    return Object.fromEntries(cells.map(([a, c]) => [a, c.format.font.color]));`);
  record("7.4.4 роли ячеек: входы, формулы, ссылка на другой лист, контроль",
    res.cards === 1 && res.state === "verified" &&
      colors.B2 === "#0000FF" && colors.B3 === "#0000FF" && colors.C2 === "#008000" && colors.C3 === "#000000" && colors.B4 === "#000000" && colors.B5 === "#C00000" && colors.A2 === "#000000" && colors.B1 === "#000000" && JSON.stringify(b.yearLabels) === JSON.stringify(["B1", "C1"]),
    `executionState: ${res.state}; counts: ${JSON.stringify(b.counts)}; цвета: ${JSON.stringify(colors)}`);
  record("7.4.4 названы заменённый ручной цвет и правило, перекрывающее цвет",
    JSON.stringify(b.overwritten) === JSON.stringify(["B3"]) && /значение не равно 0/.test(b.conditionalNote ?? "") && /останется|перекрывает|будет виден/.test(b.conditionalNote ?? "") && /по умолчанию/.test(b.paletteNote ?? ""),
    `overwritten: ${JSON.stringify(b.overwritten)}; conditionalNote: ${b.conditionalNote}; paletteNote: ${b.paletteNote}`);
  await waitFor("!!__e.button('Отменить') && !__e.button('Отменить').disabled", "кнопка «Отменить»", 20000);
  await evaluate(`__e.button('Отменить').click(); true`);
  await sleep(3000);
  const back = await excel(`
    const s = ctx.workbook.worksheets.getItem('${F}');
    const cells = ['B2','B3','C2','B5'].map((a) => { const c = s.getRange(a); c.format.font.load('color'); return [a, c]; });
    await ctx.sync();
    return Object.fromEntries(cells.map(([a, c]) => [a, c.format.font.color]));`);
  record("7.4.4 отмена вернула прежние цвета, в том числе ручной красный",
    back.B3 === "#FF0000" && back.B2 === "#000000" && back.C2 === "#000000" && back.B5 === "#000000",
    JSON.stringify(back));
}

if (wanted("7.5.1")) {
  const A = "Э7Аудит", G = "Э7Вход";
  await excel(`
    for (const n of ['${A}', '${G}', 'Э7Удалить']) { const old = ctx.workbook.worksheets.getItemOrNullObject(n); old.load('isNullObject'); await ctx.sync(); if (!old.isNullObject) { old.delete(); await ctx.sync(); } }
    const g = ctx.workbook.worksheets.add('${G}'); g.getRange('A1:B1').values = [['Рост', 0.1]];
    const x = ctx.workbook.worksheets.add('Э7Удалить'); x.getRange('A1').values = [[5]];
    const s = ctx.workbook.worksheets.add('${A}');
    s.getRange('A1:E10').formulas = [
      ['Показатель', 2025, 2026, 2027, 2028],
      ['Выручка', 1000, '=B2*(1+${G}!B1)', '=C2*(1+${G}!B1)', '=D2*(1+${G}!B1)'],
      ['Затраты', 600, '=B3*1.05', 700, '=D3*1.05'],
      ['Прибыль', '=B2-B3', '=C2-C3', '=D2-C3', '=E2-E3'],
      ['Маржа', '=B4/B6', '=C4/C2', '=D4/D2', '=E4/E2'],
      ['', '', '', '', ''],
      ['Удвоенная', '=B5*2', '', '', ''],
      ['Потеря', '=Э7Удалить!A1+1', '', '', ''],
      ['Проверка', '=B4-(B2-B3)', '=C4-(C2-C3)+5', '', ''],
      ['Динамика', '=INDIRECT("B2")', '', '', '']];
    await ctx.sync();
    ctx.workbook.worksheets.getItem('Э7Удалить').delete();
    s.activate();
    await ctx.sync();`);
  const before = await excel(`const r = ctx.workbook.worksheets.getItem('${A}').getRange('A1:E10'); r.load('formulas'); await ctx.sync(); return JSON.stringify(r.formulas);`);
  const res = await run("audit_workbook", { sheet: A, checks: [`${A}!B9:C9`] });
  const b = res.result.result ?? res.result;
  const cells = (list, pattern) => (b[list] ?? []).filter((item) => pattern.test(item.reason)).map((item) => item.cell);
  record("7.5.1 доказанное: исходная ошибка, потерянная ссылка, несходящаяся проверка — без следствия",
    JSON.stringify(cells("proven", /деление на ноль/)) === '["B5"]' && JSON.stringify(cells("proven", /потерянная ссылка/)) === '["B8"]' &&
      JSON.stringify(cells("proven", /контрольное равенство/)) === '["C9"]' && !(b.proven ?? []).some((item) => item.cell === "B7") && b.totals?.errorsConsequence === 1,
    `proven: ${JSON.stringify((b.proven ?? []).map((item) => item.cell + " " + item.content + " → " + item.value))}; следствий: ${b.totals?.errorsConsequence}`);
  record("7.5.1 подозрения: число вместо формулы, формула не как у соседей, пустой вход",
    JSON.stringify(cells("suspicions", /число вместо формулы/)) === '["D3"]' && JSON.stringify(cells("suspicions", /не такая, как у соседей/)) === '["D4"]' &&
      JSON.stringify(cells("suspicions", /пустую ячейку B6/)) === '["B5"]',
    JSON.stringify((b.suspicions ?? []).map((item) => item.cell + ": " + item.reason.slice(0, 60))));
  record("7.5.1 непроверенное: INDIRECT; межлистовая ссылка не подозрение",
    JSON.stringify(cells("unverified", /INDIRECT/)) === '["B10"]' && !cells("suspicions", /./).includes("C2"),
    JSON.stringify(b.unverified ?? []));
  const after = await excel(`const r = ctx.workbook.worksheets.getItem('${A}').getRange('A1:E10'); r.load('formulas'); await ctx.sync(); return JSON.stringify(r.formulas);`);
  record("7.5.1 аудит ничего не изменил и не спрашивал подтверждения", res.cards === 0 && before === after, `карточек: ${res.cards}`);
}

if (wanted("7.5.2")) {
  const T = "Э7Доли";
  await excel(`
    const old = ctx.workbook.worksheets.getItemOrNullObject('${T}'); old.load('isNullObject'); await ctx.sync(); if (!old.isNullObject) { old.delete(); await ctx.sync(); }
    const s = ctx.workbook.worksheets.add('${T}');
    s.getRange('A1:D4').values = [['Город', 2024, 2025, 2026], ['Москва', 100, 110, 121], ['Казань', 0, 50, 60], ['Омск', 300, '', 330]];
    s.activate();
    await ctx.sync();`);
  const res = await run("add_share_growth", { sheet: T, address: "A1:D4" });
  const b = res.result.result ?? res.result;
  const block = await excel(`const r = ctx.workbook.worksheets.getItem('${T}').getRange('A6:D18'); r.load(['values','formulas','numberFormat']); await ctx.sync(); return { values: r.values, formulas: r.formulas, nf: r.numberFormat[2][1] };`);
  const near = (a, e) => typeof a === "number" && Math.abs(a - e) < 1e-9;
  record("7.5.2 доли и рост: блок сверен, контроль 100 %, неопределённый рост назван",
    res.cards === 1 && res.state === "verified" && block.values[5].slice(1).every((v) => near(v, 1)) && near(block.values[2][1], 0.25) &&
      near(block.values[9][2], 0.1) && block.values[10][2] === "" && JSON.stringify(b.undefinedGrowth) === '["C16","D17"]' && block.nf === "0.0%",
    `executionState: ${res.state}; контроль: ${JSON.stringify(block.values[5])}; B8: ${block.values[2][1]}; C15 (рост Москвы): ${block.values[9][2]}; C16: «${block.values[10][2]}»; формат: ${block.nf}\nB8 = ${block.formulas[2][1]}; B11 = ${block.formulas[5][1]}`);
  await excel(`ctx.workbook.worksheets.getItem('${T}').getRange('B2').values = [[500]]; await ctx.sync();`);
  const again = await excel(`const r = ctx.workbook.worksheets.getItem('${T}').getRange('B8:B11'); r.load('values'); await ctx.sync(); return r.values.map((x) => x[0]);`);
  record("7.5.2 блок — формулы: правка исходника пересчитала доли, контроль по-прежнему 100 %",
    near(again[0], 500 / 800) && near(again[3], 1), JSON.stringify(again));
  await excel(`ctx.workbook.worksheets.getItem('${T}').getRange('B2').values = [[100]]; await ctx.sync();`);
  await waitFor("!!__e.button('Отменить') && !__e.button('Отменить').disabled", "кнопка «Отменить»", 20000);
  await evaluate(`__e.button('Отменить').click(); true`);
  await sleep(2500);
  const cleared = await excel(`const r = ctx.workbook.worksheets.getItem('${T}').getRange('A6:D18'); r.load(['formulas','numberFormat']); await ctx.sync(); return { empty: r.formulas.every((row) => row.every((v) => v === '')), nf: r.numberFormat[2][1] };`);
  record("7.5.2 отмена очистила блок и вернула формат", cleared.empty && cleared.nf === "General", JSON.stringify(cleared));
}

if (wanted("7.5.2-cmp")) {
  const C = "Э7Сравн";
  await excel(`
    const old = ctx.workbook.worksheets.getItemOrNullObject('${C}'); old.load('isNullObject'); await ctx.sync(); if (!old.isNullObject) { old.delete(); await ctx.sync(); }
    const s = ctx.workbook.worksheets.add('${C}');
    s.getRange('A1:D5').values = [['Компания','Выручка','Маржа','Долг'],['Альфа',500,0.2,100],['Бета',300,0.25,''],['Гамма',300,0.1,0],['Дельта',900,0.15,50]];
    s.activate();
    await ctx.sync();`);
  const res = await run("add_comparison", { sheet: C, address: "A1:D5" });
  const b = res.result.result ?? res.result;
  const block = await excel(`const r = ctx.workbook.worksheets.getItem('${C}').getRange('A7:D26'); r.load(['values','formulas']); await ctx.sync(); return { values: r.values, formulas: r.formulas };`);
  const v = block.values;
  record("7.5.2 сравнение: статистика, отклонение от медианы и места сверены, равные делят место",
    res.cards === 1 && res.state === "verified" && v[2][1] === 500 && v[3][1] === 400 && v[4][3] === 0 &&
      Math.abs(v[9][1] - 0.25) < 1e-12 && v[10][3] === "" && JSON.stringify([v[16][1], v[17][1], v[18][1], v[19][1]]) === "[2,3,3,1]" &&
      JSON.stringify(b.undefinedDeviation) === '["D17"]',
    `executionState: ${res.state}; среднее/медиана выручки: ${v[2][1]}/${v[3][1]}; мин. долга: ${v[4][3]}; места по выручке: ${JSON.stringify([v[16][1], v[17][1], v[18][1], v[19][1]])}; undefinedDeviation: ${JSON.stringify(b.undefinedDeviation)}\nB16 = ${block.formulas[9][1]}\nB23 = ${block.formulas[16][1]}`);
  await waitFor("!!__e.button('Отменить') && !__e.button('Отменить').disabled", "кнопка «Отменить»", 20000);
  await evaluate(`__e.button('Отменить').click(); true`);
  await sleep(2500);
  const cleared = await excel(`const r = ctx.workbook.worksheets.getItem('${C}').getRange('A7:D26'); r.load('formulas'); await ctx.sync(); return r.formulas.every((row) => row.every((x) => x === ''));`);
  record("7.5.2 отмена убрала блок сравнения", cleared === true, String(cleared));
}

if (wanted("7.5.3")) {
  const M = "Э7Модель";
  await excel(`const old = ctx.workbook.worksheets.getItemOrNullObject('${M}'); old.load('isNullObject'); await ctx.sync(); if (!old.isNullObject) { old.delete(); await ctx.sync(); }`);
  const assumptions = {
    revenue0: 1000, growth: 0.1, cogsPct: 0.6, opexPct: 0.2, daPct: 0.05, capexPct: 0.06,
    receivablesPct: 0.1, inventoryPct: 0.08, payablesPct: 0.07, taxRate: 0.2, interestRate: 0.1,
    repayment: 50, payout: 0.3, cash0: 100, ppe0: 500, receivables0: 100, inventory0: 80, payables0: 70, debt0: 300, equity0: 410
  };
  const missing = await run("build_three_statement_model", { sheet: M, currency: "руб.", units: "тыс.", source: "проверка", firstYear: 2026, years: 5, assumptions: { ...assumptions, cash0: 150 } });
  record("7.5.3 несходящийся баланс на начало — отказ до записи с разницей",
    missing.cards === 0 && /разница 50/.test(missing.reply) && (await excel(`const w = ctx.workbook.worksheets.getItemOrNullObject('${M}'); w.load('isNullObject'); await ctx.sync(); return w.isNullObject;`)),
    missing.reply.slice(0, 200));
  const res = await run("build_three_statement_model", { sheet: M, currency: "руб.", units: "тыс.", source: "проверка", firstYear: 2026, years: 5, assumptions });
  const b = res.result.result ?? res.result;
  const sheetData = await excel(`
    const u = ctx.workbook.worksheets.getItem('${M}').getUsedRange(true); u.load(['values','address']); await ctx.sync();
    return { address: u.address, values: u.values };`);
  const find = (label) => sheetData.values.find((row) => row[0] === label);
  const check = find("Активы − обязательства и капитал (должно быть 0)");
  record("7.5.3 модель построена, каждая ячейка сверена, баланс сходится во всех годах",
    res.cards === 1 && res.state === "verified" && check.slice(1).every((v) => v === 0) && Math.abs(find("Чистая прибыль")[2] - 108) < 1e-9 && find("Долг")[6] === 50,
    `executionState: ${res.state}; проверено ячеек: ${b.checkedCells}; контроль: ${JSON.stringify(check.slice(1))}; прибыль 2026: ${find("Чистая прибыль")[2]}; последний год: ${JSON.stringify(b.lastYear)}`);
  // Допущения — живые ячейки: меняем рост, баланс по-прежнему сходится.
  const growthRow = sheetData.values.findIndex((row) => row[0] === "Рост выручки в год") + 1;
  const after = await excel(`
    const s = ctx.workbook.worksheets.getItem('${M}'); s.getRange('B${growthRow}').values = [[0.3]]; await ctx.sync();
    const u = s.getUsedRange(true); u.load('values'); await ctx.sync();
    const out = { check: u.values.find((r) => String(r[0]).startsWith('Активы −')).slice(1), rev: u.values.find((r) => r[0] === 'Выручка')[2] };
    s.getRange('B${growthRow}').values = [[0.1]]; await ctx.sync();
    return out;`);
  record("7.5.3 модель — формулы: другой рост пересчитал выручку, баланс по-прежнему сходится",
    Math.abs(after.rev - 1300) < 1e-9 && after.check.every((v) => v === 0), JSON.stringify(after));
  await waitFor("!!__e.button('Отменить') && !__e.button('Отменить').disabled", "кнопка «Отменить»", 20000);
  await evaluate(`__e.button('Отменить').click(); true`);
  await sleep(2500);
  const gone = await excel(`const w = ctx.workbook.worksheets.getItemOrNullObject('${M}'); w.load('isNullObject'); await ctx.sync(); return w.isNullObject;`);
  record("7.5.3 отмена удалила лист модели", gone === true, String(gone));
}

if (wanted("7.5.4")) {
  const D = "Э7DCF";
  await excel(`const old = ctx.workbook.worksheets.getItemOrNullObject('${D}'); old.load('isNullObject'); await ctx.sync(); if (!old.isNullObject) { old.delete(); await ctx.sync(); }`);
  const assumptions = { revenue0: 1000, growth: 0.1, ebitMargin: 0.2, taxRate: 0.2, daPct: 0.04, capexPct: 0.05, nwcPct: 0.1, wacc: 0.12, terminalGrowth: 0.03, netDebt: 200, shares: 100 };
  const bad = await run("build_dcf_model", { sheet: D, currency: "руб.", units: "млн", source: "проверка", firstYear: 2026, years: 5, assumptions: { ...assumptions, terminalGrowth: 0.13 } });
  record("7.5.4 рост после прогноза не меньше WACC — отказ до записи", bad.cards === 0 && /меньше WACC/.test(bad.reply), bad.reply.slice(0, 160));
  const res = await run("build_dcf_model", { sheet: D, currency: "руб.", units: "млн", source: "проверка", firstYear: 2026, years: 5, assumptions });
  const b = res.result.result ?? res.result;
  // Независимый расчёт в сценарии: потоки 2026–2030 и Гордон.
  let rev = 1000, nwcPrev = 100, pv = 0, fcf = 0;
  for (let t = 1; t <= 5; t++) {
    rev *= 1.1; const ebit = rev * 0.2; const nwc = rev * 0.1;
    fcf = ebit * 0.8 + rev * 0.04 - rev * 0.05 - (nwc - nwcPrev); nwcPrev = nwc;
    pv += fcf / 1.12 ** t;
  }
  const ev = pv + fcf * 1.03 / 0.09 / 1.12 ** 5;
  const sheetData = await excel(`const u = ctx.workbook.worksheets.getItem('${D}').getUsedRange(true); u.load('values'); await ctx.sync(); return u.values;`);
  const row = (label) => sheetData.find((r) => r[0] === label);
  const evCell = row("Стоимость бизнеса (EV)")[1];
  const check = row("Контроль: центр таблицы − EV (должно быть 0)")[1];
  record("7.5.4 оценка построена: EV совпал с независимым расчётом, центр чувствительности = EV",
    res.cards === 1 && res.state === "verified" && Math.abs(evCell - ev) < 1e-6 && check === 0 && Math.abs(row("Стоимость одной акции")[1] - (ev - 200) / 100) < 1e-8,
    `executionState: ${res.state}; EV в Excel ${evCell}, независимо ${ev}; контроль ${check}; доля остаточной ${b.terminalValueShare}; EV/EBITDA ${b.evToEbitda}`);
  await waitFor("!!__e.button('Отменить') && !__e.button('Отменить').disabled", "кнопка «Отменить»", 20000);
  await evaluate(`__e.button('Отменить').click(); true`);
  await sleep(2500);
  const gone = await excel(`const w = ctx.workbook.worksheets.getItemOrNullObject('${D}'); w.load('isNullObject'); await ctx.sync(); return w.isNullObject;`);
  record("7.5.4 отмена удалила лист оценки", gone === true, String(gone));
}

if (wanted("7.5.5")) {
  const L = "Э7LBO";
  await excel(`const old = ctx.workbook.worksheets.getItemOrNullObject('${L}'); old.load('isNullObject'); await ctx.sync(); if (!old.isNullObject) { old.delete(); await ctx.sync(); }`);
  const assumptions = { ebitda0: 100, entryMultiple: 8, debtMultiple: 5, fees: 20, ebitdaGrowth: 0.08, daPct: 0.2, capexPct: 0.25, nwcPct: 0.3, taxRate: 0.2, interestRate: 0.09, cashSweep: 1, exitMultiple: 8 };
  const bad = await run("build_lbo_model", { sheet: L, currency: "руб.", units: "млн", source: "проверка", entryYear: 2025, years: 5, assumptions: { ...assumptions, debtMultiple: 9 } });
  record("7.5.5 долг больше цены сделки — отказ до записи", bad.cards === 0 && /покрывает всю цену/.test(bad.reply), bad.reply.slice(0, 160));
  const res = await run("build_lbo_model", { sheet: L, currency: "руб.", units: "млн", source: "проверка", entryYear: 2025, years: 5, assumptions });
  const b = res.result.result ?? res.result;
  // Независимый расчёт в сценарии.
  let ebitda = 100, debt = 500, cash = 0;
  for (let t = 1; t <= 5; t++) {
    const prev = ebitda; ebitda *= 1.08;
    const da = ebitda * 0.2, interest = debt * 0.09, ebt = ebitda - da - interest, tax = Math.max(0, ebt) * 0.2;
    const fcf = ebt - tax + da - ebitda * 0.25 - (ebitda - prev) * 0.3;
    const repay = Math.min(debt, Math.max(0, fcf)); debt -= repay; cash += fcf - repay;
  }
  const exitEquity = ebitda * 8 - debt + cash;
  const moic = exitEquity / 320;
  const sheet = await excel(`const u = ctx.workbook.worksheets.getItem('${L}').getUsedRange(true); u.load(['values','numberFormat']); await ctx.sync(); return { values: u.values, nf: u.numberFormat };`);
  const rowIndex = (label) => sheet.values.findIndex((r) => r[0] === label);
  const checks = ["Источники − использование на входе", "Деньги на выходе − (Σ потоков − Σ погашений)", "Долг на выходе − (долг на входе − Σ погашений)"].map((label) => sheet.values[rowIndex(label)][1]);
  const moicRow = rowIndex("Кратность денег (MOIC)");
  record("7.5.5 LBO построена: MOIC совпал с независимым расчётом, три контроля — ноль",
    res.cards === 1 && res.state === "verified" && Math.abs(sheet.values[moicRow][1] - moic) < 1e-9 && checks.every((v) => v === 0),
    `executionState: ${res.state}; MOIC ${sheet.values[moicRow][1]} (независимо ${moic}); IRR ${b.irr}; контроль ${JSON.stringify(checks)}; формат MOIC ${sheet.nf[moicRow][1]}`);
  await waitFor("!!__e.button('Отменить') && !__e.button('Отменить').disabled", "кнопка «Отменить»", 20000);
  await evaluate(`__e.button('Отменить').click(); true`);
  await sleep(2500);
  const gone = await excel(`const w = ctx.workbook.worksheets.getItemOrNullObject('${L}'); w.load('isNullObject'); await ctx.sync(); return w.isNullObject;`);
  record("7.5.5 отмена удалила лист LBO", gone === true, String(gone));
}

console.log(`прошло ${results.filter(Boolean).length} из ${results.length}`);
socket.close();
process.exit(results.every(Boolean) ? 0 : 1);
