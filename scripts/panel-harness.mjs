/**
 * Проверка поведения панели без Excel: настоящая собранная панель
 * в настоящем Chromium (тот же движок, что у WebView2 в Excel).
 *
 * Зачем. Защита от двух нажатий и от гонки задачи с отменой — это поведение
 * интерфейса, и проверять его на модульных тестах мало: они не видят, успевает
 * ли React перерисовать панель между нажатиями. А руками гонку не поймать —
 * кнопка сменяется быстрее, чем человек нажимает второй раз.
 *
 * Как. Chrome запускается без окна, панель грузится из каталога сборки
 * `releases/<выпуск>/panel`, а вместо Excel и модели отвечает этот скрипт
 * через протокол DevTools: `office.js` подменяется поддельным Excel,
 * запросы к `/api/chat` задерживаются, пока скрипт не решит ответить или
 * оборвать. Нажатия и клавиши — настоящие события браузера.
 *
 * Запуск: node scripts/panel-harness.mjs <выпуск> [<выпуск> …]
 * Зависимостей нет: WebSocket встроен в Node.
 */

import { spawn } from "node:child_process";
import { mkdtempSync, readFileSync, rmSync, existsSync } from "node:fs";
import { tmpdir } from "node:os";
import { extname, join, resolve } from "node:path";

const root = resolve(import.meta.dirname, "..");
const CHROME = [
  "C:/Program Files/Google/Chrome/Application/chrome.exe",
  "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"
].find((path) => existsSync(path));
const ORIGIN = "https://harness.test";
const PORT = 9333;

const TYPES = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".png": "image/png", ".svg": "image/svg+xml" };

/** Поддельный Excel: ровно то, что панели нужно при старте и перед задачей. */
const OFFICE_STUB = `
window.Office = {
  HostType: { Excel: "Excel" },
  onReady(cb) { const info = { host: "Excel" }; setTimeout(() => cb && cb(info)); return Promise.resolve(info); },
  context: { document: { url: "C:/harness/проверка.xlsx" }, requirements: { isSetSupported: () => false } }
};
const item = (props) => Object.assign({ load() {} }, props);
window.Excel = {
  run: async (fn) => fn({
    workbook: {
      worksheets: { getActiveWorksheet: () => item({ id: "sheet-1", name: "Лист1" }) },
      getActiveCell: () => item({ address: "Лист1!A1" }),
      getSelectedRange: () => item({ address: "Лист1!A1" })
    },
    sync: async () => undefined
  })
};`;

const ANSWER = [
  `data: ${JSON.stringify({ choices: [{ delta: { content: "Готово." }, finish_reason: "stop" }] })}`,
  "data: [DONE]",
  ""
].join("\n\n");

function sleep(ms) { return new Promise((r) => setTimeout(r, ms)); }

async function connect() {
  for (let attempt = 0; attempt < 50; attempt++) {
    try {
      const list = await (await fetch(`http://127.0.0.1:${PORT}/json/list`)).json();
      const page = list.find((target) => target.type === "page");
      if (page) return page.webSocketDebuggerUrl;
    } catch { /* Chrome ещё поднимается */ }
    await sleep(200);
  }
  throw new Error("Chrome не ответил на порту отладки.");
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
    ready: new Promise((r) => socket.addEventListener("open", r)),
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

async function runRelease(release) {
  const panelDir = join(root, "releases", release, "panel");
  if (!existsSync(join(panelDir, "taskpane.html"))) throw new Error(`Нет сборки панели ${panelDir}.`);

  const profile = mkdtempSync(join(tmpdir(), "panel-harness-"));
  const chrome = spawn(CHROME, [
    "--headless=new", `--remote-debugging-port=${PORT}`, `--user-data-dir=${profile}`,
    "--no-first-run", "--no-default-browser-check", "--disable-extensions", "about:blank"
  ], { stdio: "ignore" });

  const client = cdp(await connect());
  await client.ready;

  // Запросы к модели копятся здесь и ждут решения проверки.
  const chats = [];
  client.on("Fetch.requestPaused", async ({ requestId, request }) => {
    const url = new URL(request.url);
    const fulfil = (body, type, status = 200) => client.send("Fetch.fulfillRequest", {
      requestId, responseCode: status,
      responseHeaders: [{ name: "Content-Type", value: type }, { name: "Cache-Control", value: "no-store" }],
      body: Buffer.from(body).toString("base64")
    }).catch(() => undefined);

    if (url.hostname === "appsforoffice.microsoft.com") return fulfil(OFFICE_STUB, "text/javascript");
    if (url.origin !== ORIGIN) return client.send("Fetch.failRequest", { requestId, errorReason: "BlockedByClient" }).catch(() => undefined);
    if (url.pathname === "/api/providers") {
      return fulfil(JSON.stringify([{ id: "fake", label: "Проверка", models: ["m"], defaultModel: "m", capabilities: ["chat"] }]), "application/json");
    }
    if (url.pathname === "/api/chat") {
      chats.push({
        answer: () => fulfil(ANSWER, "text/event-stream"),
        fail: () => client.send("Fetch.failRequest", { requestId, errorReason: "ConnectionRefused" }).catch(() => undefined)
      });
      return;
    }
    const file = join(panelDir, decodeURIComponent(url.pathname));
    if (!file.startsWith(panelDir) || !existsSync(file)) return fulfil("нет", "text/plain", 404);
    return fulfil(readFileSync(file), TYPES[extname(file)] ?? "application/octet-stream");
  });

  await client.send("Fetch.enable", { patterns: [{ urlPattern: "*" }] });
  await client.send("Runtime.enable");
  await client.send("Page.navigate", { url: `${ORIGIN}/taskpane.html` });

  const evaluate = async (expression) => {
    const { result, exceptionDetails } = await client.send("Runtime.evaluate", { expression, returnByValue: true, awaitPromise: true });
    if (exceptionDetails) throw new Error(exceptionDetails.exception?.description ?? exceptionDetails.text);
    return result.value;
  };
  const waitFor = async (expression, what, ms = 5000) => {
    const until = Date.now() + ms;
    while (Date.now() < until) {
      if (await evaluate(expression).catch(() => false)) return;
      await sleep(50);
    }
    throw new Error(`не дождались: ${what}`);
  };

  // Помощники внутри страницы: ввод в поле React и поиск кнопок по тексту.
  await waitFor("!!document.querySelector('textarea')", "панель загрузилась", 15000);
  await evaluate(`
    window.__h = {
      type(text) {
        const area = document.querySelector('textarea');
        Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(area, text);
        area.dispatchEvent(new Event('input', { bubbles: true }));
      },
      button(text) { return [...document.querySelectorAll('button')].find((b) => b.textContent.trim().startsWith(text)); },
      state() {
        const send = this.button('Отправить'), stop = this.button('Остановить'), undo = this.button('Отменить') || this.button('Undo');
        return {
          areaDisabled: document.querySelector('textarea').disabled,
          send: send ? (send.disabled ? 'выключена' : 'активна') : 'нет',
          stop: stop ? 'есть' : 'нет',
          undo: undo ? (undo.disabled ? 'выключена' : 'активна') : 'нет',
          log: document.querySelector('.log').innerText
        };
      }
    };
    true`);
  // Снять режим «Только анализ» не нужно: задача без инструментов, модель отвечает текстом.
  await waitFor("__h.button('Отправить') !== undefined", "кнопка «Отправить»");

  const results = [];
  const record = (name, ok, details) => results.push({ name, ok, details });
  const count = (text) => evaluate(`__h.state().log.split(${JSON.stringify(text)}).length - 1`);
  const finishAll = async () => { while (chats.length) chats.shift().answer(); await sleep(300); };

  // 1. Два нажатия мышью в одном такте: React не успевает перерисовать.
  {
    const text = "проверка-двойной-клик";
    await evaluate(`__h.type(${JSON.stringify(text)}); true`);
    await waitFor("!__h.button('Отправить').disabled", "кнопка включилась после ввода");
    await evaluate(`const b = __h.button('Отправить'); b.click(); b.click(); true`);
    await sleep(700);
    const requests = chats.length;
    const shown = await count(text);
    record("два клика в одном такте", requests === 1 && shown === 1, `запросов к модели: ${requests}, просьба в ленте: ${shown} раз`);
    await finishAll();
    await waitFor("__h.button('Отправить') !== undefined", "задача завершилась");
  }

  // 2. Зажатый Enter: настоящие события клавиатуры с автоповтором.
  {
    const text = "проверка-зажатый-enter";
    await evaluate(`__h.type(${JSON.stringify(text)}); document.querySelector('textarea').focus(); true`);
    for (let index = 0; index < 8; index++) {
      await client.send("Input.dispatchKeyEvent", {
        type: "keyDown", key: "Enter", code: "Enter", windowsVirtualKeyCode: 13, autoRepeat: index > 0
      });
    }
    await client.send("Input.dispatchKeyEvent", { type: "keyUp", key: "Enter", code: "Enter", windowsVirtualKeyCode: 13 });
    await sleep(700);
    const requests = chats.length;
    const shown = await count(text);
    record("зажатый Enter", requests === 1 && shown === 1, `запросов к модели: ${requests}, просьба в ленте: ${shown} раз`);

    // 3. Пока задача идёт, всё, что меняет книгу, выключено.
    const busy = await evaluate("__h.state()");
    record(
      "панель заблокирована во время задачи",
      busy.areaDisabled && busy.stop === "есть" && busy.undo !== "активна",
      `поле ввода ${busy.areaDisabled ? "выключено" : "ВКЛЮЧЕНО"}, «Остановить»: ${busy.stop}, «Отменить»: ${busy.undo}`
    );
    await finishAll();
    await waitFor("__h.button('Отправить') !== undefined", "задача завершилась");
  }

  // 4. Остановка посреди задачи и сразу новая задача.
  {
    await evaluate(`__h.type('проверка-остановка'); true`);
    await waitFor("!__h.button('Отправить').disabled", "кнопка включилась");
    await evaluate(`__h.button('Отправить').click(); true`);
    await waitFor("__h.button('Остановить') !== undefined", "появилась «Остановить»");
    await evaluate(`__h.button('Остановить').click(); true`);
    await waitFor("__h.state().log.includes('Остановлено вами')", "сообщение об остановке");
    chats.length = 0; // брошенный запрос больше не ждём
    await evaluate(`__h.type('проверка-после-остановки'); true`);
    await sleep(200);
    await evaluate(`__h.button('Отправить') && __h.button('Отправить').click(); true`);
    await sleep(700);
    const after = await evaluate("__h.state()");
    record("после остановки можно работать дальше", chats.length === 1, `новых запросов к модели: ${chats.length}; поле ввода ${after.areaDisabled ? "выключено (идёт новая задача)" : "включено"}`);
    await finishAll();
    await waitFor("__h.button('Отправить') !== undefined", "задача завершилась");
  }

  // 5. Обрыв соединения с сервером: панель не должна остаться заблокированной.
  {
    await evaluate(`__h.type('проверка-обрыв-сети'); true`);
    await waitFor("!__h.button('Отправить').disabled", "кнопка включилась");
    await evaluate(`__h.button('Отправить').click(); true`);
    for (let waited = 0; !chats.length && waited < 5000; waited += 50) await sleep(50);
    chats.shift()?.fail();
    await sleep(800);
    const after = await evaluate("__h.state()");
    record(
      "после ошибки сети панель свободна",
      !after.areaDisabled && after.stop === "нет",
      `поле ввода ${after.areaDisabled ? "ВЫКЛЮЧЕНО" : "включено"}, «Остановить»: ${after.stop}`
    );
  }

  client.close();
  chrome.kill();
  await sleep(500);
  try { rmSync(profile, { recursive: true, force: true }); } catch { /* Chrome ещё держит файлы */ }
  return results;
}

const releases = process.argv.slice(2);
if (!releases.length) {
  console.error("Укажите выпуск: node scripts/panel-harness.mjs <выпуск> [<выпуск> …]");
  process.exit(2);
}
let failed = 0;
for (const release of releases) {
  console.log(`\n== ${release}`);
  try {
    for (const result of await runRelease(release)) {
      if (!result.ok) failed += 1;
      console.log(`${result.ok ? "  прошла " : "  ПРОВАЛ "} ${result.name}: ${result.details}`);
    }
  } catch (error) {
    failed += 1;
    console.log(`  ОШИБКА ПРОВЕРКИ: ${error.message}`);
  }
}
process.exit(failed ? 1 : 0);
