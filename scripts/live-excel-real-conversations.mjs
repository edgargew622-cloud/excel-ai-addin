/**
 * Набор проверок из живых бесед (с 01.10.2026): просьбы, на которых агент
 * ошибался у пользователя, повторяются в Excel на модели по умолчанию, и итог
 * проверяется по книге, а не по словам модели. Каждая найденная в пилоте
 * ошибка добавляется сюда сценарием.
 *
 * Запуск: Excel с --remote-debugging-port=9229 (WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS),
 * открытая панель am.AI, ключ DeepSeek; node scripts/live-excel-real-conversations.mjs
 *
 * А. 01.10.2026: 11 заголовков жирным по одной ячейке → предел 8 изменений.
 *    Ждём одну операцию format_range со списком «A3,A5,…» и ровно 11 жирных строк.
 * Б. 30.09.2026: «создай таблицу 10×6» на пустом листе → вставка строк и столбцов.
 *    Ждём: без insert_rows / insert_columns.
 */
// 10.2: цвета диаграмм — просьбы обычными словами, проверка по Excel.
const list = await (await fetch("http://127.0.0.1:9229/json/list")).json();
const page = list.find((t) => t.type === "page" && t.url.includes("localhost:3000/taskpane.html"));
const ws = new WebSocket(page.webSocketDebuggerUrl);
await new Promise((r) => ws.addEventListener("open", r));
let id = 0; const waiting = new Map();
ws.addEventListener("message", (e) => { const m = JSON.parse(e.data); if (m.id && waiting.has(m.id)) { waiting.get(m.id)(m); waiting.delete(m.id); } });
const ev = (expression) => new Promise((resolve) => { const i = ++id; waiting.set(i, (m) => resolve(m.result?.result?.value ?? m.result?.exceptionDetails?.exception?.description)); ws.send(JSON.stringify({ id: i, method: "Runtime.evaluate", params: { expression, returnByValue: true, awaitPromise: true } })); });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const excel = (body) => ev(`(async () => { try { let out; await Excel.run(async (ctx) => { out = await (async () => { ${body} })(); }); return out; } catch (e) { return 'ОШИБКА: ' + e.message; } })()`);
const btn = (t) => `[...document.querySelectorAll('button')].find((b) => b.textContent.trim().startsWith(${JSON.stringify(t)}))`;
async function ask(text, { clear = false } = {}) {
  if (clear) { await ev(`(() => { const b = ${btn("Очистить")}; b && b.click(); return 1; })()`); await sleep(400); }
  await ev(`(() => { const box = [...document.querySelectorAll('label')].find((l) => l.textContent.includes('Только анализ'))?.querySelector('input'); if (box && box.checked) box.click(); return 1; })()`);
  await sleep(300);
  await ev(`(() => { const a = document.querySelector('textarea:not(.copy-area)'); Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(a, ${JSON.stringify(text)}); a.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
  await sleep(300); await ev(`${btn("Отправить")}.click(); 1`);
  const started = Date.now(); const cards = []; await sleep(1500);
  while (Date.now() - started < 240000) {
    const card = await ev(`(() => { const b = ${btn("Выполнить")}; if (!b) return null; const p = document.querySelector('.preview'); return p ? p.innerText.replace(/\s+/g, ' ').slice(0, 300) : 'карточка'; })()`);
    if (card) { cards.push(card); await ev(`${btn("Выполнить")}.click(); 1`); await sleep(900); continue; }
    if (await ev(`!!${btn("Отправить")} && !document.querySelector('.thinking')`)) break;
    await sleep(500);
  }
  await sleep(800);
  const reply = String(await ev(`[...document.querySelectorAll('.msg.assistant')].pop()?.innerText ?? ''`)).replace(/\s+/g, " ").slice(0, 260);
  const ops = await ev(`[...document.querySelectorAll('.op')].slice(-6).map(o => o.innerText.replace(/\\s+/g,' ')).join(' | ')`);
  return { cards, reply, ops, seconds: Math.round((Date.now() - started) / 1000) };
}

// Повторы реальных бесед 30.09–01.10.2026 на DeepSeek (набор проверок из живых бесед).
const HEAD = [3, 5, 8, 12, 15, 18, 32, 42, 50, 52, 58];
const lines = Array.from({ length: 58 }, (_, i) => HEAD.includes(i + 1) ? `БЛОК ${i + 1}: ЗАГОЛОВОК РАЗДЕЛА` : `Строка ${i + 1}: обычный текст документа, реквизиты и отметки`);
console.log("листы:", await excel(`
  for (const name of ['Путевой', 'Пустой']) { const old = ctx.workbook.worksheets.getItemOrNullObject(name); old.load('isNullObject'); await ctx.sync(); if (!old.isNullObject) { old.delete(); await ctx.sync(); } }
  const s = ctx.workbook.worksheets.add('Путевой'); s.getRange('A1:A58').values = ${JSON.stringify(lines.map((l) => [l]))};
  ctx.workbook.worksheets.add('Пустой'); s.activate(); await ctx.sync(); return 'ок';`));
await ev(`(() => { const set = (el, v) => { Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(el, v); el.dispatchEvent(new Event('change', { bubbles: true })); }; set(document.querySelector('select[aria-label="Провайдер"]'), 'deepseek'); return 1; })()`);
await sleep(500);

let r = await ask("На листе Путевой выдели жирным все заголовки блоков — строки, написанные ЗАГЛАВНЫМИ буквами.", { clear: true });
console.log("\nА) заголовки жирным:", JSON.stringify({ ops: r.ops, cards: r.cards.length, seconds: r.seconds }));
console.log("   карточка:", (r.cards[0] ?? "").slice(0, 200));
console.log("   Excel:", await excel(`const s = ctx.workbook.worksheets.getItem('Путевой'); const cells = Array.from({ length: 58 }, (_, i) => { const c = s.getRange('A' + (i + 1)); c.format.font.load('bold'); return c; }); await ctx.sync();
  const boldRows = cells.map((c, i) => c.format.font.bold ? i + 1 : 0).filter(Boolean); return 'жирные строки: ' + boldRows.join(',') + ' | ожидалось: ${HEAD.join(",")}';`));
console.log("   ответ:", r.reply.slice(0, 220));

await excel(`ctx.workbook.worksheets.getItem('Пустой').activate(); await ctx.sync(); return 1;`);
r = await ask("привет как дела ?, создай таблицу на 10 строк и 6 стобцов", { clear: true });
console.log("\nБ) таблица на пустом листе:", JSON.stringify({ ops: r.ops, cards: r.cards.length, seconds: r.seconds }));
console.log("   Excel:", await excel(`const s = ctx.workbook.worksheets.getItem('Пустой'); s.tables.load('items/name'); const u = s.getUsedRangeOrNullObject(); u.load('address,isNullObject'); await ctx.sync();
  const t = s.tables.items[0]; let a = ''; if (t) { const rg = t.getRange(); rg.load('address'); await ctx.sync(); a = rg.address; } return 'таблиц: ' + s.tables.items.length + (a ? ' ' + a : '') + ' | занято: ' + (u.isNullObject ? 'ничего' : u.address);`));
console.log("   ответ:", r.reply.slice(0, 220));
ws.close();
