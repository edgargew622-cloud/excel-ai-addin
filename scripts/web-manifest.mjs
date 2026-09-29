/**
 * Манифест веб-панели (Mac, Excel в браузере) из основного manifest.xml.
 *
 * Панель лежит на GitHub Pages, поэтому все адреса https://localhost:3000/
 * заменяются адресом сайта, а Id — свой: на одном компьютере веб-панель и
 * панель Windows не должны считаться одной надстройкой. Версия берётся из
 * основного манифеста — разойтись они не могут. Токена в этом манифесте нет:
 * локального сервера, который он защищает, у веб-панели нет.
 *
 *   node scripts/web-manifest.mjs [--base https://…/] [--out dist-web/manifest.xml]
 *   Рядом с манифестом пишется mac.html — установка на Mac одной командой.
 */

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

export const WEB_ADDIN_ID = "5409e1d3-ee5f-49a0-be50-a9e930bce2af";
export const DEFAULT_BASE = "https://edgargew622-cloud.github.io/excel-ai-addin/";

export function webManifest(source, base = DEFAULT_BASE) {
  if (!/^https:\/\/[^\s"<>]+\/$/.test(base)) throw new Error(`Адрес сайта должен быть https и кончаться на /: ${base}`);
  const local = "https://localhost:3000/";
  if (!source.includes(local)) throw new Error("В manifest.xml нет адресов https://localhost:3000/ — нечего заменять.");
  let out = source.split(local).join(base);
  out = out.replace(/<Id>[^<]+<\/Id>/, `<Id>${WEB_ADDIN_ID}</Id>`);
  out = out.replace(/<DisplayName DefaultValue="[^"]*" \/>/, '<DisplayName DefaultValue="am.AI (Mac и Excel в браузере)" />');
  out = out.replace(
    /<Description DefaultValue="[^"]*" \/>/,
    '<Description DefaultValue="AI-агент для Excel без установки программы: модели DeepSeek и OpenRouter по вашему ключу." />'
  );
  // Панель из интернета ходит к DeepSeek и OpenRouter сама — домены объявляются явно.
  out = out.replace("<Hosts>", `<AppDomains>\n    <AppDomain>https://api.deepseek.com</AppDomain>\n    <AppDomain>https://openrouter.ai</AppDomain>\n  </AppDomains>\n\n  <Hosts>`);
  if (out.includes("localhost")) throw new Error("В веб-манифесте остался localhost.");
  return out;
}

/** Папка, из которой Excel для Mac берёт надстройки пользователя. */
export const MAC_WEF = "~/Library/Containers/com.microsoft.Excel/Data/Documents/wef";
export const MAC_FILE = "am-ai.xml";

/**
 * Установка на Mac одной командой: создать папку wef и скачать туда манифест.
 * Команда ничего не запускает — только кладёт XML-файл, поэтому Mac не
 * спрашивает про неизвестного разработчика; её можно прочитать целиком.
 * Ошибка скачивания (-f) останавливает цепочку до «Готово».
 */
export function macInstallCommand(base = DEFAULT_BASE) {
  return `mkdir -p ${MAC_WEF} && curl -fsSL ${base}manifest.xml -o ${MAC_WEF}/${MAC_FILE} && echo "Готово. Перезапустите Excel."`;
}

export function macUninstallCommand() {
  return `rm -f ${MAC_WEF}/${MAC_FILE} && echo "Удалено. Перезапустите Excel."`;
}

function htmlEscape(value) {
  return value.replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;" })[c]);
}

/** Страница установки для Mac на сайте панели: шаги, команда с кнопкой «Скопировать», удаление. */
export function macPage(base = DEFAULT_BASE) {
  const install = htmlEscape(macInstallCommand(base));
  const uninstall = htmlEscape(macUninstallCommand());
  return `<!doctype html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>am.AI для Mac</title>
<style>
  :root { --bg: #ffffff; --fg: #1d1d1f; --muted: #5f6368; --code: #f3f4f6; --line: #e3e5e8; --accent: #1a73e8; }
  @media (prefers-color-scheme: dark) { :root { --bg: #151618; --fg: #ececed; --muted: #a0a4a8; --code: #222428; --line: #33363b; --accent: #7ab4ff; } }
  body { margin: 0; background: var(--bg); color: var(--fg); font: 16px/1.55 -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
  main { max-width: 720px; margin: 0 auto; padding: 32px 16px 48px; }
  h1 { font-size: 28px; margin: 0 0 4px; }
  h2 { font-size: 19px; margin: 32px 0 8px; }
  p, li { color: var(--fg); }
  .muted { color: var(--muted); }
  ol { padding-left: 22px; }
  li { margin: 8px 0; }
  .cmd { display: flex; gap: 8px; align-items: stretch; margin: 10px 0; }
  code.block { flex: 1; min-width: 0; display: block; background: var(--code); border: 1px solid var(--line); border-radius: 8px; padding: 10px 12px; font: 13px/1.5 ui-monospace, Menlo, monospace; overflow-x: auto; white-space: pre; }
  button { font: inherit; font-size: 14px; border: 1px solid var(--line); background: var(--code); color: var(--fg); border-radius: 8px; padding: 0 14px; cursor: pointer; }
  button:hover { border-color: var(--accent); }
  kbd { font: 13px ui-monospace, Menlo, monospace; border: 1px solid var(--line); border-radius: 4px; padding: 1px 5px; }
  a { color: var(--accent); }
</style>
</head>
<body>
<main>
  <h1>am.AI для Excel на Mac</h1>
  <p class="muted">AI-агент в Excel без установки программы. Модели — DeepSeek или OpenRouter по вашему ключу.</p>

  <h2>Установка</h2>
  <ol>
    <li>Откройте Терминал: <kbd>⌘</kbd>&nbsp;+&nbsp;<kbd>Пробел</kbd>, наберите «Терминал», <kbd>Enter</kbd>.</li>
    <li>Скопируйте команду, вставьте её в Терминал и нажмите <kbd>Enter</kbd>:
      <div class="cmd"><code class="block" id="install">${install}</code><button data-copy="install">Скопировать</button></div>
      Должно появиться «Готово. Перезапустите Excel.»
    </li>
    <li>Закройте Excel полностью (<kbd>⌘</kbd>&nbsp;+&nbsp;<kbd>Q</kbd>) и откройте снова.</li>
    <li>В книге: «Вставка» → «Надстройки» → «Мои надстройки» → <b>am.AI (Mac и Excel в браузере)</b>. На некоторых версиях кнопка «Надстройки» — на вкладке «Главная».</li>
    <li>В панели нажмите «Ключи» и вставьте ключ — достаточно одного:
      <ul>
        <li><b>DeepSeek</b> — дешёвые модели, ключ на <a href="https://platform.deepseek.com/api_keys">platform.deepseek.com</a> (нужно пополнить баланс на пару долларов);</li>
        <li><b>OpenRouter</b> — бесплатные модели с пометкой <code>:free</code>, а также Mistral и Gemini; ключ на <a href="https://openrouter.ai/keys">openrouter.ai/keys</a>. У бесплатных бывают перебои и дневной лимит.</li>
      </ul>
    </li>
  </ol>

  <h2>Что делает команда</h2>
  <p>Создаёт папку, из которой Excel для Mac берёт надстройки, и скачивает в неё файл описания надстройки (<a href="${htmlEscape(base)}manifest.xml">manifest.xml</a>). Программа не устанавливается и ничего не запускается, пароль не нужен. Сама панель загружается с этого сайта и обновляется сама.</p>

  <h2>Удаление</h2>
  <div class="cmd"><code class="block" id="uninstall">${uninstall}</code><button data-copy="uninstall">Скопировать</button></div>
  <p>Затем перезапустите Excel. Ключи удаляются кнопкой «Удалить» в окне «Ключи» — сделайте это до удаления надстройки.</p>

  <h2>Что нужно знать</h2>
  <ul>
    <li>Нужен Excel для Mac 2019, 2021, 2024 или Microsoft 365.</li>
    <li>Ключ хранится в самой панели на этом компьютере, без Связки ключей, — не вводите его на чужом Mac.</li>
    <li>Всё, что агент прочитал из книги, уходит выбранному поставщику модели — DeepSeek или OpenRouter. Бесплатные модели поставщики могут использовать для обучения.</li>
    <li>Нет памяти, прикреплённых файлов, поиска в интернете и резервной копии книги — это есть только в версии для Windows.</li>
    <li>На настоящем Mac панель ещё не проверялась (проверена в движке WebKit, на котором работает Excel для Mac). Если что-то не так — напишите в <a href="https://github.com/edgargew622-cloud/excel-ai-addin/issues">Issues</a>.</li>
  </ul>
</main>
<script>
  for (const button of document.querySelectorAll("button[data-copy]")) {
    button.addEventListener("click", async () => {
      const text = document.getElementById(button.dataset.copy).textContent;
      try { await navigator.clipboard.writeText(text); button.textContent = "Скопировано"; }
      catch { const range = document.createRange(); range.selectNodeContents(document.getElementById(button.dataset.copy)); const sel = getSelection(); sel.removeAllRanges(); sel.addRange(range); button.textContent = "Выделено — ⌘C"; }
      setTimeout(() => { button.textContent = "Скопировать"; }, 2000);
    });
  }
</script>
</body>
</html>
`;
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
  const option = (name, fallback) => {
    const i = process.argv.indexOf(`--${name}`);
    return i > 0 ? process.argv[i + 1] : fallback;
  };
  const base = option("base", process.env.WEB_PANEL_BASE || DEFAULT_BASE);
  const out = resolve(root, option("out", "dist-web/manifest.xml"));
  writeFileSync(out, webManifest(readFileSync(resolve(root, "manifest.xml"), "utf8"), base));
  console.log(`Веб-манифест: ${out} (${base})`);
  const page = resolve(dirname(out), "mac.html");
  writeFileSync(page, macPage(base));
  console.log(`Страница установки для Mac: ${page}`);
}
