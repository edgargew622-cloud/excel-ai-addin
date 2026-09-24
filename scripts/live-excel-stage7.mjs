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

console.log(`прошло ${results.filter(Boolean).length} из ${results.length}`);
socket.close();
process.exit(results.every(Boolean) ? 0 : 1);
