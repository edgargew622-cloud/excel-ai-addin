/**
 * Приёмка S7: естественные просьбы к настоящей модели в настоящем Excel.
 *
 * В отличие от live-excel-s*.mjs, ответы модели здесь не подменяются:
 * панель ходит к своему серверу, тот — к провайдеру. Скрипт только печатает
 * просьбу, подтверждает карточки и после задачи проверяет книгу.
 *
 * Для каждой просьбы — свежий лист «П<n>» с копией данных листа «Заказы»
 * (A1:D13), чтобы просьбы не зависели друг от друга. Прочие листы книги
 * сверяются до и после: лишних изменений быть не должно.
 *
 * Записывается отдельно: верна ли цель и результат, завершена ли задача,
 * не тронуто ли лишнее, ответ модели (точность отчёта оценивается по нему
 * вручную), время, число вызовов, токены по /api/metrics сервера.
 *
 * Запуск: node scripts/live-excel-model.mjs [провайдер] [модель] [номера просьб…]
 * По умолчанию deepseek и модель провайдера по умолчанию.
 */

const PORT = 9229;
const SOURCE = "Заказы";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const [providerId = "deepseek", modelArg, ...only] = process.argv.slice(2);
const model = modelArg && modelArg !== "-" ? modelArg : null;
const wanted = (n) => only.length === 0 || only.includes(String(n));

const list = await (await fetch(`http://127.0.0.1:${PORT}/json/list`)).json();
const page = list.find((target) => target.type === "page" && target.url.includes("taskpane.html"));
if (!page) throw new Error("Панель надстройки не найдена: откройте её в Excel.");
const socket = new WebSocket(page.webSocketDebuggerUrl);
let nextId = 0;
const waiting = new Map();
socket.addEventListener("message", (event) => {
  const message = JSON.parse(event.data);
  if (message.id && waiting.has(message.id)) {
    const { resolve, reject } = waiting.get(message.id);
    waiting.delete(message.id);
    message.error ? reject(new Error(message.error.message)) : resolve(message.result);
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
    await sleep(200);
  }
  throw new Error(`не дождались: ${what}`);
};
const excel = (body) => evaluate(`(async () => { let out; await Excel.run(async (ctx) => { out = await (async () => { ${body} })(); }); return out; })()`);
const metrics = async () => {
  process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";
  return (await (await fetch("https://localhost:3000/api/metrics")).json()).summary;
};

await evaluate(`
  window.__m = {
    button(text) { return [...document.querySelectorAll('button')].find((b) => b.textContent.trim().startsWith(text)); },
    type(text) {
      const area = document.querySelector('textarea');
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(area, text);
      area.dispatchEvent(new Event('input', { bubbles: true }));
    },
    pick(label, value) {
      const select = document.querySelector('select[aria-label="' + label + '"]');
      Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(select, value);
      select.dispatchEvent(new Event('change', { bubbles: true }));
      return select.value;
    },
    ops() { return [...document.querySelectorAll('.op')].map((op) => ({ status: op.className.replace('op', '').trim(), text: op.innerText })); },
    answers() { return [...document.querySelectorAll('.log .msg.assistant')].map((m) => m.innerText); }
  };
  true`);
await waitFor("!!__m.button('Отправить')", "панель свободна");
await evaluate(`__m.pick('Провайдер', ${JSON.stringify(providerId)}); true`);
await sleep(300);
if (model) await evaluate(`__m.pick('Модель', ${JSON.stringify(model)}); true`);
await sleep(300);
const chosen = await evaluate(`[document.querySelector('select[aria-label="Провайдер"]').value, document.querySelector('select[aria-label="Модель"]').value]`);
await evaluate(`(() => { const box = [...document.querySelectorAll('label')].find((l) => l.textContent.includes('Только анализ'))?.querySelector('input'); if (box && box.checked) box.click(); return true; })()`);
console.log(`Панель: ${await evaluate("document.querySelector('.persistence-note')?.innerText ?? ''")}`);
console.log(`Модель: ${chosen.join(" / ")}\n`);

/* --- книга ------------------------------------------------------------------- */

const others = ["Заказы", "Занято", "Курс", "Связанные"];
const snapshot = () => excel(`
  const out = {};
  for (const name of ${JSON.stringify(others)}) {
    const s = ctx.workbook.worksheets.getItemOrNullObject(name); s.load('isNullObject'); await ctx.sync();
    if (s.isNullObject) { out[name] = null; continue; }
    const u = s.getUsedRangeOrNullObject(); u.load(['isNullObject', 'address', 'formulas']);
    const charts = s.charts; charts.load('items/name'); const pivots = s.pivotTables; pivots.load('items/name');
    await ctx.sync();
    out[name] = JSON.stringify([u.isNullObject ? null : u.address, u.isNullObject ? null : u.formulas, charts.items.length, pivots.items.length]);
  }
  const sheets = ctx.workbook.worksheets; sheets.load('items/name'); await ctx.sync();
  out.__sheets = sheets.items.map((s) => s.name).join(',');
  return out;`);

/** Свежий лист с копией данных; прежние листы приёмки удаляются. */
const freshSheet = (name) => excel(`
  const src = ctx.workbook.worksheets.getItem('${SOURCE}').getRange('A1:D13'); src.load('values'); await ctx.sync();
  for (const old of ['${name}', 'Итоги']) {
    const s = ctx.workbook.worksheets.getItemOrNullObject(old); s.load('isNullObject'); await ctx.sync();
    if (!s.isNullObject && (old === '${name}' || ${name === "П7" ? "true" : "false"})) { s.delete(); await ctx.sync(); }
  }
  const sheet = ctx.workbook.worksheets.add('${name}');
  sheet.getRange('A1:D13').values = src.values;
  sheet.activate();
  await ctx.sync();
  return src.values;`);

const read = (sheet, address, props = ["values"]) => excel(`
  const r = ctx.workbook.worksheets.getItem('${sheet}').getRange('${address}'); r.load(${JSON.stringify(props)}); await ctx.sync();
  return Object.fromEntries(${JSON.stringify(props)}.map((p) => [p, r[p]]));`);
const usedValues = (sheet) => excel(`
  const s = ctx.workbook.worksheets.getItemOrNullObject('${sheet}'); s.load('isNullObject'); await ctx.sync();
  if (s.isNullObject) return null;
  const u = s.getUsedRangeOrNullObject(); u.load(['isNullObject', 'values', 'address']); await ctx.sync();
  return u.isNullObject ? { address: null, values: [] } : { address: u.address, values: u.values };`);
const flat = (text) => String(text).replace(/[\s  ]/g, "");
const rowsKey = (rows) => rows.map((row) => JSON.stringify(row)).sort().join("|");
const near = (a, b) => typeof a === "number" && Math.abs(a - b) < 1e-6;

/* --- просьбы ------------------------------------------------------------------- */

const requests = [
  {
    n: 1, sheet: "П1", kind: "чтение",
    text: "Какая общая сумма заказов по Омску на листе П1?",
    check: async ({ answer, writes }) => {
      const ok = /4900/.test(flat(answer)) && writes === 0;
      return [ok, `в ответе 4900: ${/4900/.test(flat(answer)) ? "да" : "нет"}; изменяющих вызовов: ${writes}`];
    }
  },
  {
    n: 2, sheet: "П2", kind: "заполнение",
    text: "На листе П2 добавь в столбец E сумму с НДС 20% для каждого заказа, с заголовком.",
    check: async ({ source }) => {
      const r = await read("П2", "A1:E13", ["values", "formulas"]);
      const header = r.values[0][4];
      const bad = r.values.slice(1).filter((row) => !near(row[4], row[2] * 1.2)).length;
      const formulas = r.formulas.slice(1).filter((row) => String(row[4]).startsWith("=")).length;
      const dataKept = rowsKey(r.values.map((row) => row.slice(0, 4))) === rowsKey(source);
      return [header !== "" && bad === 0 && dataKept, `заголовок E1: ${JSON.stringify(header)}; неверных сумм: ${bad} из 12; формулами: ${formulas} из 12; исходные данные целы: ${dataKept ? "да" : "НЕТ"}`];
    }
  },
  {
    n: 3, sheet: "П3", kind: "сортировка",
    text: "Отсортируй заказы на листе П3 по сумме от большей к меньшей.",
    check: async ({ source }) => {
      const r = await read("П3", "A1:D13");
      const sums = r.values.slice(1).map((row) => row[2]);
      const sorted = sums.every((v, i) => i === 0 || sums[i - 1] >= v);
      const kept = rowsKey(r.values) === rowsKey(source) && r.values[0][0] === "Город";
      return [sorted && kept, `по убыванию: ${sorted ? "да" : "НЕТ"} (${sums.join(", ")}); строки целы, заголовок на месте: ${kept ? "да" : "НЕТ"}`];
    }
  },
  {
    n: 4, sheet: "П4", kind: "фильтр",
    text: "На листе П4 оставь видимыми только заказы из Казани.",
    check: async ({ source }) => {
      const hidden = await excel(`
        const s = ctx.workbook.worksheets.getItem('П4'); const rows = [];
        for (let i = 1; i <= 13; i++) { const r = s.getRange('A' + i); r.load(['rowHidden', 'values']); rows.push(r); }
        await ctx.sync(); return rows.map((r) => [r.values[0][0], r.rowHidden]);`);
      const wrong = hidden.slice(1).filter(([city, isHidden]) => (city === "Казань") === isHidden).length;
      const r = await read("П4", "A1:D13");
      const kept = rowsKey(r.values) === rowsKey(source);
      return [wrong === 0 && !hidden[0][1] && kept, `строк с неверной видимостью: ${wrong}; заголовок виден: ${!hidden[0][1] ? "да" : "НЕТ"}; данные целы: ${kept ? "да" : "НЕТ"}`];
    }
  },
  {
    n: 5, sheet: "П5", kind: "оформление",
    text: "Сделай заголовки таблицы на листе П5 жирными и с серой заливкой.",
    check: async () => {
      const f = await excel(`
        const r = ctx.workbook.worksheets.getItem('П5').getRange('A1:D1'); r.format.font.load('bold'); r.format.fill.load('color');
        const below = ctx.workbook.worksheets.getItem('П5').getRange('A2:D13'); below.format.font.load('bold');
        await ctx.sync(); return [r.format.font.bold, r.format.fill.color, below.format.font.bold];`);
      const grey = typeof f[1] === "string" && /^#([0-9A-F]{2})\1\1$/i.test(f[1]) && f[1].toUpperCase() !== "#FFFFFF";
      return [f[0] === true && grey && f[2] === false, `жирный: ${f[0]}; заливка: ${f[1]} (${grey ? "серая" : "не серая"}); строки данных не жирные: ${f[2] === false ? "да" : "НЕТ"}`];
    }
  },
  {
    n: 6, sheet: "П6", kind: "сводная",
    text: "На листе П6 построй сводную таблицу: сумма заказов по городам.",
    check: async ({ source }) => {
      const pivots = await excel(`
        const s = ctx.workbook.worksheets.getItem('П6'); const p = s.pivotTables; p.load('items/name'); await ctx.sync();
        const out = [];
        for (const item of p.items) { const r = item.layout.getRange(); r.load(['address', 'values']); await ctx.sync(); out.push({ address: r.address, values: r.values }); }
        return out;`);
      const cells = pivots.flatMap((p) => p.values.map((row) => row.join("|")));
      const has = (city, sum) => cells.some((row) => row.includes(city) && row.includes(String(sum)));
      const ok = pivots.length === 1 && has("Москва", 3950) && has("Казань", 2150) && has("Омск", 4900) && cells.some((row) => row.includes("11000"));
      const r = await read("П6", "A1:D13");
      const kept = rowsKey(r.values) === rowsKey(source);
      return [ok && kept, `сводных на листе: ${pivots.length}${pivots[0] ? ` (${pivots[0].address})` : ""}; суммы по городам и итог 11000 верны: ${ok ? "да" : "НЕТ"}; данные целы: ${kept ? "да" : "НЕТ"}`];
    }
  },
  {
    n: 7, sheet: "П7", kind: "новый лист",
    text: "Создай новый лист «Итоги» и выпиши туда сумму заказов по каждому городу из листа П7.",
    check: async () => {
      const u = await usedValues("Итоги");
      if (!u) return [false, "листа «Итоги» нет"];
      const rows = u.values.map((row) => row.map(String));
      const has = (city, sum) => rows.some((row) => row.includes(city) && row.includes(String(sum)));
      const ok = has("Москва", 3950) && has("Казань", 2150) && has("Омск", 4900);
      return [ok, `лист «Итоги»: ${u.address}; суммы Москва 3950, Казань 2150, Омск 4900: ${ok ? "верны" : "НЕТ"}; содержимое: ${JSON.stringify(u.values).slice(0, 200)}`];
    }
  },
  {
    n: 8, sheet: "П8", kind: "условное оформление",
    text: "На листе П8 выдели красным суммы больше 1000.",
    check: async () => {
      const cf = await excel(`
        const s = ctx.workbook.worksheets.getItem('П8'); const c = s.getRange('A1:Z50').conditionalFormats; c.load('items/type'); await ctx.sync();
        const out = [];
        for (const item of c.items) { const r = item.getRange(); r.load('address'); out.push([item.type, r]); }
        await ctx.sync(); return out.map(([type, r]) => type + ' ' + r.address);`);
      const onSums = cf.some((rule) => /!C/.test(rule));
      return [cf.length >= 1 && onSums, `правил: ${cf.length} (${cf.join("; ") || "—"}); на столбце сумм: ${onSums ? "да" : "НЕТ"}. Цвет и условие правила не сверяются (S3.2) — по ответу модели и карточке`];
    }
  },
  {
    n: 9, sheet: "П9", kind: "удаление строк",
    text: "На листе П9 удали все закрытые заказы.",
    // Удаление строк не отменяется, и модель вправе переспросить перед ним.
    // Тогда отвечает «пользователь» — одним сообщением, как ответил бы человек.
    followUp: "Да, удаляй. Резервная копия не нужна.",
    check: async () => {
      const u = await usedValues("П9");
      const rows = u.values.slice(1).filter((row) => row.some((v) => v !== ""));
      const closed = rows.filter((row) => row[1] === "Закрыта").length;
      const total = rows.reduce((sum, row) => sum + (typeof row[2] === "number" ? row[2] : 0), 0);
      return [closed === 0 && rows.length === 8 && total === 7050 && u.values[0][0] === "Город", `строк данных: ${rows.length} (ожидалось 8); закрытых: ${closed}; сумма оставшихся: ${total} (ожидалось 7050)`];
    }
  }
];

/* --- прогон ------------------------------------------------------------------- */

const table = [];
for (const request of requests.filter((r) => wanted(r.n))) {
  await waitFor("!!__m.button('Очистить') && !__m.button('Очистить').disabled", "кнопка «Очистить»", 300000);
  await evaluate(`__m.button('Очистить').click(); true`);
  await sleep(300);
  const source = await freshSheet(request.sheet);
  const before = await snapshot();
  const m0 = await metrics();

  const started = Date.now();
  const cards = [];
  let timedOut = false;
  const turn = async (text) => {
    const turnStarted = Date.now();
    await evaluate(`__m.type(${JSON.stringify(text)}); true`);
    await waitFor("!!__m.button('Отправить') && !__m.button('Отправить').disabled", "кнопка «Отправить»");
    await evaluate(`__m.button('Отправить').click(); true`);
    await sleep(1500);
    while (Date.now() - turnStarted < 300000) {
      if (await evaluate("!!__m.button('Выполнить')")) {
        cards.push(await evaluate("document.querySelector('.confirm')?.innerText.split('\\n')[0] ?? ''"));
        await evaluate(`__m.button('Выполнить').click(); true`);
        await sleep(500);
        continue;
      }
      if (await evaluate("!!__m.button('Отправить') && !document.querySelector('.thinking')")) return;
      await sleep(250);
    }
    timedOut = true;
    await evaluate(`__m.button('Остановить')?.click(); true`);
    await waitFor("!!__m.button('Отправить')", "остановка", 60000).catch(() => undefined);
  };
  const isWrite = (op) => !/^(get_|list_|search_|recall_|measure_|create_workbook_backup)/.test(op.text);
  await turn(request.text);
  await sleep(500);
  // Модель переспросила и ничего не изменила — отвечаем как пользователь.
  const followedUp = Boolean(request.followUp) && !timedOut && (await evaluate("__m.ops()")).filter(isWrite).length === 0;
  if (followedUp) await turn(request.followUp);
  const seconds = (Date.now() - started) / 1000;
  await sleep(500);
  const ops = await evaluate("__m.ops()");
  const answers = await evaluate("__m.answers()");
  const answer = answers.at(-1) ?? "";
  // Правило инструкции — отвечать только по-русски; ищем китайский и японский.
  const foreign = answers.some((text) => /[぀-ヿ一-鿿]/.test(text));
  const m1 = await metrics();
  const after = await snapshot();
  const touched = others.filter((name) => before[name] !== after[name]);
  const extraSheets = after.__sheets.split(",").filter((name) => !before.__sheets.split(",").includes(name));
  const writes = ops.filter(isWrite).length;

  let verdict;
  try { verdict = await request.check({ source, answer, writes, ops }); } catch (error) { verdict = [false, `проверка не удалась: ${error.message}`]; }
  const clean = touched.length === 0 && (request.n === 7 ? extraSheets.every((name) => name === "Итоги") : extraSheets.length === 0);
  const ok = verdict[0] && clean && !timedOut && !foreign;
  table.push({ n: request.n, kind: request.kind, ok, seconds, calls: ops.length, cards: cards.length, tokens: m1.totalTokens - m0.totalTokens, requests: m1.requests - m0.requests, followedUp, foreign });
  console.log(`${ok ? "прошла " : "ПРОВАЛ "} ${request.n}. ${request.kind}: «${request.text}»`);
  console.log(`        результат: ${verdict[1]}`);
  console.log(`        лишние изменения: ${touched.length ? `тронуты ${touched.join(", ")}` : "нет"}${extraSheets.length ? `; новые листы: ${extraSheets.join(", ")}` : ""}${timedOut ? "; ЗАДАЧА НЕ ЗАКОНЧИЛАСЬ за 5 минут" : ""}`);
  console.log(`        вызовы (${ops.length}): ${ops.map((op) => `${op.text.split(/[\s\n]/)[0]}[${op.status}]`).join(", ") || "—"}`);
  console.log(`        карточки (${cards.length}): ${cards.join(" | ") || "—"}`);
  if (followedUp) console.log(`        модель переспросила; ответ пользователя: «${request.followUp}»`);
  if (foreign) console.log("        ОТВЕТ НЕ ПО-РУССКИ");
  console.log(`        ${seconds.toFixed(0)} с; запросов к модели: ${m1.requests - m0.requests}; токенов: ${m1.totalTokens - m0.totalTokens}`);
  console.log(`        ответ модели: ${answer.replace(/\s+/g, " ").slice(0, 600)}\n`);
}

console.log("№ | просьба | итог | с | вызовов | карточек | запросов | токенов");
for (const row of table) console.log(`${row.n} | ${row.kind} | ${row.ok ? "прошла" : "ПРОВАЛ"}${row.followedUp ? " (после ответа на вопрос)" : ""}${row.foreign ? " (не по-русски)" : ""} | ${row.seconds.toFixed(0)} | ${row.calls} | ${row.cards} | ${row.requests} | ${row.tokens}`);
console.log(`\nпрошло ${table.filter((row) => row.ok).length} из ${table.length}`);
socket.close();
process.exit(0);
