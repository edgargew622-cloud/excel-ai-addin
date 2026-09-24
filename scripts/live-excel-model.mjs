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
  },
  // Этап 7, 7.1.3: «посчитай и запиши» — в книге формулы, а не числа модели.
  {
    n: 10, sheet: "П10", kind: "итог формулой",
    text: "На листе П10 в ячейку F1 запиши общую сумму всех заказов.",
    check: async () => {
      const r = await read("П10", "F1", ["formulas", "values"]);
      const formula = String(r.formulas[0][0]);
      return [formula.startsWith("=") && r.values[0][0] === 11000, `F1: ${formula} → ${r.values[0][0]} (ожидалась формула, итог 11000)`];
    }
  },
  {
    n: 11, sheet: "П11", kind: "среднее формулой",
    text: "На листе П11 посчитай среднюю сумму заказа по Казани и запиши результат рядом с таблицей.",
    check: async () => {
      const found = await formulaCellsBeyondD("П11");
      const hit = found.find((cell) => Math.abs(Number(cell.value) - 537.5) < 1e-9);
      return [Boolean(hit), `ячейки правее D: ${JSON.stringify(found)} (ожидалась формула со значением 537,5)`];
    }
  },
  {
    n: 12, sheet: "П12", kind: "доля формулами",
    text: "На листе П12 добавь столбец «Доля» с долей каждого заказа от общей суммы.",
    check: async () => {
      const u = await usedValues("П12");
      const column = u.values[0].findIndex((header) => /доля/i.test(String(header)));
      if (column < 4) return [false, `столбца «Доля» правее данных нет; шапка: ${JSON.stringify(u.values[0])}`];
      const letter = String.fromCharCode(65 + column);
      const r = await read("П12", `${letter}2:${letter}13`, ["formulas", "values"]);
      const formulas = r.formulas.filter((row) => String(row[0]).startsWith("=")).length;
      const sums = u.values.slice(1).map((row) => row[2]);
      const bad = r.values.filter((row, i) => Math.abs(Number(row[0]) - sums[i] / 11000) > 1e-9 && Math.abs(Number(row[0]) - sums[i] / 110) > 1e-9).length;
      return [formulas === 12 && bad === 0, `столбец ${letter}: формулами ${formulas} из 12; неверных долей: ${bad}; E2 = ${r.formulas[0][0]}`];
    }
  },
  {
    n: 13, sheet: "П13", kind: "итог значением по просьбе",
    text: "На листе П13 запиши в F1 общую сумму заказов значением, без формулы.",
    check: async () => {
      const r = await read("П13", "F1", ["formulas", "values"]);
      const formula = String(r.formulas[0][0]);
      return [!formula.startsWith("=") && r.values[0][0] === 11000, `F1: ${formula} (ожидалось число 11000 без формулы)`];
    }
  }
];

// Этап 7, 7.2: очистка «грязной» выгрузки обычными словами.
requests.push({
  n: 14, sheet: "П14", kind: "очистка данных",
  allowSheets: ["П14 чисто"],
  text: "На листе П14 почисти данные: убери лишние пробелы, преврати числа и даты, записанные текстом, в настоящие и убери повторяющиеся строки.",
  setup: () => excel(`
    for (const name of ['П14', 'П14 чисто']) {
      const old = ctx.workbook.worksheets.getItemOrNullObject(name); old.load('isNullObject'); await ctx.sync();
      if (!old.isNullObject) { old.delete(); await ctx.sync(); }
    }
    const s = ctx.workbook.worksheets.add('П14');
    s.getRange('A1:D9').values = [
      ['Город', 'Сумма', 'Дата', 'Код'],
      ["' Москва", "'1 200", "'25.02.2026", "'007"],
      ['Москва', 900, "'01.02.2026", "'12"],
      ["'Казань  Север", "'1,500", "'14.03.2026", "'12"],
      ["'Омск ", "'2 300,50", "'2026-03-05", "'45"],
      ['Москва', 900, "'01.02.2026", "'12"],
      ['Тула', "'700", "'28.02.2026", "'33"],
      ['Москва', 900, "'01.02.2026", "'12"],
      ['Казань', "'450", "'03.04.2026", "'51"]];
    s.activate();
    await ctx.sync();
    return [];`),
  followUps: [
    "Даты записаны как день.месяц.год. 1,500 — это полторы тысячи, то есть 1500. Повторы убери на отдельный лист «П14 чисто», исходные данные не трогай.",
    "Да, делай так.",
    "Да."
  ],
  check: async () => {
    const source = await usedValues("П14");
    const clean = await usedValues("П14 чисто");
    const texts = source.values.slice(1).map((row) => row[0]).filter((value) => typeof value === "string" && value !== value.trim());
    const b = source.values.slice(1).map((row) => row[1]);
    const c = source.values.slice(1).map((row) => row[2]);
    const numbers = b.every((value) => typeof value === "number");
    const dates = c.every((value) => typeof value === "number");
    const duplicatesRemoved = clean ? clean.values.length - 1 : null;
    return [
      texts.length === 0 && numbers && dates && duplicatesRemoved === 6,
      `пробелы по краям осталось: ${texts.length}; суммы числами: ${numbers} (${JSON.stringify(b)}); даты числами: ${dates} (${JSON.stringify(c)}); ` +
      `лист «П14 чисто»: ${clean ? `${clean.address}, строк данных ${duplicatesRemoved} (ожидалось 6)` : "нет"}`
    ];
  }
});

// Этап 7, 7.4–7.5: аудит, цвета модели и подсветка строк обычными словами.
const modelSetup = (name, input) => excel(`
  for (const n of ['${name}', '${input}']) {
    const old = ctx.workbook.worksheets.getItemOrNullObject(n); old.load('isNullObject'); await ctx.sync();
    if (!old.isNullObject) { old.delete(); await ctx.sync(); }
  }
  const g = ctx.workbook.worksheets.add('${input}'); g.getRange('A1:B1').values = [['Рост', 0.1]];
  const s = ctx.workbook.worksheets.add('${name}');
  s.getRange('A1:E9').formulas = [
    ['Показатель', 2025, 2026, 2027, 2028],
    ['Выручка', 1000, "=B2*(1+'${input}'!B1)", "=C2*(1+'${input}'!B1)", "=D2*(1+'${input}'!B1)"],
    ['Затраты', 600, '=B3*1.05', 700, '=D3*1.05'],
    ['Прибыль', '=B2-B3', '=C2-C3', '=D2-C3', '=E2-E3'],
    ['Маржа', '=B4/B6', '=C4/C2', '=D4/D2', '=E4/E2'],
    ['', '', '', '', ''],
    ['Удвоенная', '=B5*2', '', '', ''],
    ['', '', '', '', ''],
    ['Проверка', '=B4-(B2-B3)', '=C4-(C2-C3)+5', '=D4-(D2-D3)', '=E4-(E2-E3)']];
  s.activate();
  await ctx.sync();
  return [];`);

requests.push({
  n: 15, sheet: "П15", kind: "аудит модели",
  text: "Проверь модель на листе П15: есть ли там ошибки? Строка 9 — проверка, там везде должен быть ноль. Ничего не исправляй.",
  setup: () => modelSetup("П15", "П15 вход"),
  check: async ({ answer, writes }) => {
    const named = ["B5", "D3", "D4", "C9"].filter((cell) => answer.includes(cell));
    return [
      writes === 0 && named.length === 4,
      `записей: ${writes} (ожидалось 0); названы ${named.join(", ") || "—"} из B5 (деление на пустую B6), D3 (число вместо формулы), D4 (формула не как у соседей), C9 (проверка = 5)`
    ];
  }
});

requests.push({
  n: 16, sheet: "П16", kind: "цвета финансовой модели",
  text: "Раскрась модель на листе П16 по-финансовому: входы синим, формулы чёрным, ссылки на другие листы зелёным. Строка 9 — проверочная.",
  setup: () => modelSetup("П16", "П16 вход"),
  followUps: ["Да, делай."],
  check: async () => {
    const colors = await excel(`
      const s = ctx.workbook.worksheets.getItem('П16');
      const cells = ['B1', 'B2', 'C2', 'C3', 'D3', 'B4', 'A2'].map((a) => { const c = s.getRange(a); c.format.font.load('color'); return [a, c]; });
      await ctx.sync();
      return Object.fromEntries(cells.map(([a, c]) => [a, c.format.font.color]));`);
    const blue = (color) => /^#0000FF$/i.test(color);
    return [
      blue(colors.B2) && blue(colors.D3) && colors.C2 === "#008000" && colors.C3 === "#000000" && colors.B4 === "#000000" && colors.A2 === "#000000" && colors.B1 === "#000000",
      `цвета: ${JSON.stringify(colors)} (ожидались B2, D3 синие; C2 зелёная; C3, B4 чёрные; подпись A2 и год B1 не тронуты)`
    ];
  }
});

requests.push({
  n: 17, sheet: "П17", kind: "подсветка строк по условию",
  text: "На листе П17 подсвети целиком строки, где статус «Новая», светло-красной заливкой.",
  setup: () => excel(`
    const old = ctx.workbook.worksheets.getItemOrNullObject('П17'); old.load('isNullObject'); await ctx.sync();
    if (!old.isNullObject) { old.delete(); await ctx.sync(); }
    const s = ctx.workbook.worksheets.add('П17');
    s.getRange('A1:C6').values = [['Город','Статус','Сумма'],['Москва','Новая',900],['Казань','Закрыта',450],['Омск','Новая',1500],['Тула','Закрыта',300],['Сочи','Новая',200]];
    s.activate();
    await ctx.sync();
    return [];`),
  followUps: ["Да, делай."],
  check: async () => {
    const rules = await excel(`
      const col = ctx.workbook.worksheets.getItem('П17').getRange('A1:C6').conditionalFormats; col.load('items/type'); await ctx.sync();
      const out = col.items.map((i) => { const r = i.getRange(); r.load('address'); if (i.type === 'Custom') i.custom.rule.load('formula'); return { i, r }; });
      await ctx.sync();
      return out.map(({ i, r }) => ({ type: i.type, range: r.address.replace(/^.*!/, ''), formula: i.type === 'Custom' ? i.custom.rule.formula : null }));`);
    const row = rules.find((rule) => rule.type === "Custom" && /\$B2/.test(rule.formula ?? "") && /^A2:C6$/.test(rule.range));
    return [Boolean(row) && rules.length === 1, `правила: ${JSON.stringify(rules)} (ожидалось одно правило по формуле с $B2 на A2:C6)`];
  }
});

requests.push({
  n: 18, sheet: "П18", kind: "сравнительный анализ",
  text: "Сравни компании на листе П18 по всем показателям: кто выше или ниже медианы и какие у них места.",
  setup: () => excel(`
    const old = ctx.workbook.worksheets.getItemOrNullObject('П18'); old.load('isNullObject'); await ctx.sync();
    if (!old.isNullObject) { old.delete(); await ctx.sync(); }
    const s = ctx.workbook.worksheets.add('П18');
    s.getRange('A1:D5').values = [['Компания','Выручка','Маржа','Долг'],['Альфа',500,0.2,100],['Бета',300,0.25,''],['Гамма',300,0.1,0],['Дельта',900,0.15,50]];
    s.activate();
    await ctx.sync();
    return [];`),
  followUps: ["Да, делай."],
  check: async ({ answer, ops }) => {
    const used = ops.some((op) => op.text.startsWith("add_comparison"));
    const r = await read("П18", "A7:D26", ["formulas"]);
    const formulas = r.formulas.flat().filter((f) => typeof f === "string" && f.startsWith("=")).length;
    const debt = /долг/i.test(answer) && /(меньше|наоборот|ниже.*лучше|лучше.*меньш)/i.test(answer);
    return [used && formulas > 20 && debt, `add_comparison: ${used}; формул в блоке: ${formulas}; про долг «меньше — лучше» сказано: ${debt}`];
  }
});

requests.push({
  n: 19, sheet: "П19", kind: "трёхотчётная модель",
  allowSheets: ["П19"],
  text: "Построй на новом листе П19 трёхотчётную модель компании на 5 лет, начиная с 2026 года.",
  setup: () => excel(`
    const old = ctx.workbook.worksheets.getItemOrNullObject('П19'); old.load('isNullObject'); await ctx.sync();
    if (!old.isNullObject) { old.delete(); await ctx.sync(); }
    return [];`),
  followUps: [
    "Валюта рубли, всё в тысячах. Выручка 2025 года 1000, растёт на 10 % в год. Себестоимость 60 % выручки, операционные расходы 20 %, амортизация 5 %, капвложения 6 %. " +
      "Дебиторка 10 % выручки, запасы 8 %, кредиторка 7 %. Налог 20 %, ставка по долгу 10 %, гасим 50 в год, на дивиденды 30 % прибыли. " +
      "На начало: деньги 100, основные средства 500, дебиторка 100, запасы 80, кредиторка 70, долг 300, капитал 410. Источник — мои оценки.",
    "Да, строй.",
    "Да."
  ],
  check: async ({ ops }) => {
    const firstBuild = ops.findIndex((op) => op.text.startsWith("build_three_statement_model") && op.status === "done");
    const exists = await excel(`const w = ctx.workbook.worksheets.getItemOrNullObject('П19'); w.load('isNullObject'); await ctx.sync(); return !w.isNullObject;`);
    if (!exists) return [false, `лист П19 не построен; вызовы: ${ops.map((op) => op.text.split(/\s/)[0] + "[" + op.status + "]").join(", ")}`];
    const values = await excel(`const u = ctx.workbook.worksheets.getItem('П19').getUsedRange(true); u.load('values'); await ctx.sync(); return u.values;`);
    const row = (label) => values.find((r) => r[0] === label);
    const check = values.find((r) => String(r[0]).startsWith("Активы −")).slice(1);
    const inputs = [row("Выручка последнего фактического года")[1], row("Рост выручки в год")[1], row("Капитал на начало")[1], row("Доля прибыли на дивиденды")[1]];
    return [
      firstBuild !== -1 && check.every((v) => v === 0) && JSON.stringify(inputs) === "[1000,0.1,410,0.3]" && Math.abs(row("Чистая прибыль")[2] - 108) < 1e-9,
      `контроль: ${JSON.stringify(check)}; входы: ${JSON.stringify(inputs)}; прибыль 2026: ${row("Чистая прибыль")[2]}`
    ];
  }
});

/** Непустые ячейки правее столбца D: где модель положила результат. */
async function formulaCellsBeyondD(sheet) {
  const u = await usedValues(sheet);
  const r = await read(sheet, u.address.replace(/^.*!/, ""), ["formulas", "values"]);
  const start = /([A-Z]+)(\d+)/.exec(u.address.replace(/^.*!/, ""));
  const column0 = start[1].charCodeAt(0) - 65;
  const row0 = Number(start[2]);
  const out = [];
  r.formulas.forEach((row, i) => row.forEach((formula, j) => {
    if (column0 + j >= 4 && formula !== "") out.push({ cell: `${String.fromCharCode(65 + column0 + j)}${row0 + i}`, formula, value: r.values[i][j] });
  }));
  return out;
}

/* --- прогон ------------------------------------------------------------------- */

const table = [];
for (const request of requests.filter((r) => wanted(r.n))) {
  await waitFor("!!__m.button('Очистить') && !__m.button('Очистить').disabled", "кнопка «Очистить»", 300000);
  await evaluate(`__m.button('Очистить').click(); true`);
  await sleep(300);
  const source = request.setup ? await request.setup() : await freshSheet(request.sheet);
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
  const isWrite = (op) => !/^(get_|list_|search_|recall_|measure_|profile_|audit_|create_workbook_backup)/.test(op.text);
  await turn(request.text);
  await sleep(500);
  // Модель переспросила и ничего не изменила — отвечаем как пользователь.
  let followedUp = Boolean(request.followUp) && !timedOut && (await evaluate("__m.ops()")).filter(isWrite).length === 0;
  if (followedUp) await turn(request.followUp);
  // Ответы на уточняющие вопросы: пока модель спрашивает и ответы есть.
  const replies = [];
  for (const reply of request.followUps ?? []) {
    const last = (await evaluate("__m.answers()")).at(-1) ?? "";
    // Модель спрашивает не только знаком вопроса: «Пришлите, пожалуйста, допущения».
    if (timedOut || !/\?|пришлите|укажите|напишите|подтвердите/i.test(last)) break;
    replies.push({ question: last.replace(/\s+/g, " ").slice(-400), reply });
    await turn(reply);
  }
  if (replies.length) followedUp = true;
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
  // Новые листы допустимы только там, где их прямо просят.
  const allowedSheets = request.allowSheets ?? (request.n === 7 ? ["Итоги"] : []);
  const clean = touched.length === 0 && extraSheets.every((name) => allowedSheets.includes(name));
  const ok = verdict[0] && clean && !timedOut && !foreign;
  table.push({ n: request.n, kind: request.kind, ok, seconds, calls: ops.length, cards: cards.length, tokens: m1.totalTokens - m0.totalTokens, requests: m1.requests - m0.requests, followedUp, foreign });
  console.log(`${ok ? "прошла " : "ПРОВАЛ "} ${request.n}. ${request.kind}: «${request.text}»`);
  console.log(`        результат: ${verdict[1]}`);
  console.log(`        лишние изменения: ${touched.length ? `тронуты ${touched.join(", ")}` : "нет"}${extraSheets.length ? `; новые листы: ${extraSheets.join(", ")}` : ""}${timedOut ? "; ЗАДАЧА НЕ ЗАКОНЧИЛАСЬ за 5 минут" : ""}`);
  console.log(`        вызовы (${ops.length}): ${ops.map((op) => `${op.text.split(/[\s\n]/)[0]}[${op.status}]`).join(", ") || "—"}`);
  console.log(`        карточки (${cards.length}): ${cards.join(" | ") || "—"}`);
  if (followedUp && request.followUp) console.log(`        модель переспросила; ответ пользователя: «${request.followUp}»`);
  for (const item of replies) console.log(`        вопрос модели: …${item.question}${String.fromCharCode(10)}        ответ пользователя: «${item.reply}»`);
  if (foreign) console.log("        ОТВЕТ НЕ ПО-РУССКИ");
  console.log(`        ${seconds.toFixed(0)} с; запросов к модели: ${m1.requests - m0.requests}; токенов: ${m1.totalTokens - m0.totalTokens}`);
  console.log(`        ответ модели: ${answer.replace(/\s+/g, " ").slice(0, 600)}\n`);
}

console.log("№ | просьба | итог | с | вызовов | карточек | запросов | токенов");
for (const row of table) console.log(`${row.n} | ${row.kind} | ${row.ok ? "прошла" : "ПРОВАЛ"}${row.followedUp ? " (после ответа на вопрос)" : ""}${row.foreign ? " (не по-русски)" : ""} | ${row.seconds.toFixed(0)} | ${row.calls} | ${row.cards} | ${row.requests} | ${row.tokens}`);
console.log(`\nпрошло ${table.filter((row) => row.ok).length} из ${table.length}`);
socket.close();
process.exit(0);
