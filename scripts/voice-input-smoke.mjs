/**
 * Voice input UI smoke: real Chromium + built panel, fake SpeechRecognition.
 * No microphone, external speech service or paid model is accessed.
 * Run after npm run build: node scripts/voice-input-smoke.mjs
 * Screenshot: logs/voice-preview.png (synthetic draft only).
 */
import { spawn } from "node:child_process";
import { mkdtempSync, readFileSync, writeFileSync, mkdirSync, rmSync, existsSync } from "node:fs";
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

async function runVoice() {
  const panelDir = join(root, "dist");
  if (!existsSync(join(panelDir, "taskpane.html"))) throw new Error(`Нет сборки панели ${panelDir}.`);

  const profile = mkdtempSync(join(tmpdir(), "panel-harness-"));
  const chrome = spawn(CHROME, [
    "--headless=new", `--remote-debugging-port=${PORT}`, `--user-data-dir=${profile}`,
    "--no-first-run", "--no-default-browser-check", "--disable-extensions", "about:blank"
  ], { stdio: "ignore" });

  let client;
  try {
  client = cdp(await connect());
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

  await client.send('Emulation.setDeviceMetricsOverride', {width:360,height:720,deviceScaleFactor:1,mobile:false});
  await client.send('Page.enable');
  await client.send('Page.addScriptToEvaluateOnNewDocument', {source: `
    window.__voiceStarts=0;
    window.SpeechRecognition=class {
      constructor(){window.__recognition=this;}
      start(){window.__voiceStarts++;this.onstart?.();}
      stop(){this.onresult?.({resultIndex:0,results:[{isFinal:true,0:{transcript:'voice test'}}]});this.onend?.();}
      abort(){}
    };
  `});
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


  record('microphone exists separately before send', await evaluate(`!!document.querySelector('.voice-button') && document.querySelector('.voice-button').nextElementSibling?.classList.contains('send')`), 'independent button');
  await evaluate(`__h.type('My draft');true`);
  await sleep(80);
  await evaluate(`document.querySelector('.voice-button').click();true`);
  await waitFor(`document.querySelector('.voice-button').getAttribute('aria-pressed')==='true'`,'microphone active');
  record('recording disables send without calling model', await evaluate(`document.querySelector('.send').disabled`) && chats.length===0, 'no automatic submission');
  await evaluate(`__h.type('Edited draft');true`);
  await sleep(80);
  await evaluate(`document.querySelector('.voice-button').click();true`);
  await waitFor(`document.querySelector('textarea').value.includes('voice test')`,'transcript');
  record('final speech appends to latest draft',await evaluate(`document.querySelector('textarea').value==='Edited draft voice test'`),await evaluate(`document.querySelector('textarea').value`));
  record('finishing dictation does not submit',chats.length===0,'chat calls='+chats.length);
  mkdirSync(join(root,'logs'),{recursive:true});
  const image=await client.send('Page.captureScreenshot',{format:'png'});
  writeFileSync(join(root,'logs','voice-preview.png'),Buffer.from(image.data,'base64'));
  await evaluate(`document.querySelector('.voice-button').click();true`);
  await sleep(80);
  await evaluate(`window.__recognition.onerror({error:'network'});true`);
  await waitFor(`document.querySelector('.voice-button').getAttribute('aria-pressed')==='false'`,'error released mic');
  record('service error preserves draft',await evaluate(`document.querySelector('textarea').value==='Edited draft voice test' && document.querySelector('.voice-note').innerText.length>0`),'network fallback');
  await evaluate(`delete window.SpeechRecognition;delete window.webkitSpeechRecognition;document.querySelector('.voice-button').click();true`);
  await waitFor(`document.querySelector('.voice-note').innerText.includes('Win + H')`,'unsupported fallback');
  record('unsupported host does not pretend to record',await evaluate(`document.querySelector('.voice-button').getAttribute('aria-pressed')==='false'`),'system dictation fallback');
  record('composer fits narrow pane',await evaluate(`document.documentElement.scrollWidth<=360 && document.querySelector('.voice-button').getBoundingClientRect().right<document.querySelector('.send').getBoundingClientRect().left`),'360px panel');
  return results;
  } finally {
    client?.close();
    chrome.kill();
    await sleep(500);
    try { rmSync(profile, { recursive: true, force: true }); } catch { /* browser still releasing profile */ }
  }
}

let failed = 0;
try {
  for (const result of await runVoice()) {
    if (!result.ok) failed++;
    console.log(`${result.ok ? 'PASS' : 'FAIL'} ${result.name}: ${result.details}`);
  }
} catch (error) {
  failed++;
  console.error(error.message);
}
process.exit(failed ? 1 : 0);
