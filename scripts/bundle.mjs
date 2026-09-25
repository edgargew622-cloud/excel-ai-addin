/**
 * Готовый комплект для Windows: всё, что нужно для работы надстройки, без
 * Node.js, npm и сборки на компьютере пользователя.
 *
 * Раскладка повторяет рабочую копию после npm run release, поэтому комплект
 * запускают те же сценарии: scripts\start-server.ps1 находит выбранный выпуск
 * по releases\current.json, а Node — в node\node.exe.
 *
 *   node\node.exe                     переносной Node (и его LICENSE)
 *   releases\current.json             выбранный выпуск
 *   releases\<id>\panel\              собранная панель
 *   server\releases\<id>\dist\        собранный сервер без тестов
 *   server\node_modules\              зависимости сервера без dev-пакетов
 *   scripts\*.ps1                     установка, запуск, автозапуск, регистрация, диагностика
 *   INSTALL.md                        установка для пользователя
 *   Установить.cmd                    установка двойным щелчком
 *   bundle.json                       что и из чего собрано
 *
 * Запуск после npm run check:
 *   node scripts/bundle.mjs [--out bundle] [--node <путь к node.exe>]
 * По умолчанию берётся тот Node, которым запущен скрипт: в CI это официальная
 * сборка для Windows из actions/setup-node.
 */

import { createHash } from "node:crypto";
import { execFileSync, execSync } from "node:child_process";
import { cpSync, existsSync, mkdirSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs";
import { dirname, join, relative, resolve } from "node:path";

const root = resolve(import.meta.dirname, "..");
const APP_ID = "excel-ai-addin";
const BUNDLE_NAME = "ExcelAI";
const SCRIPTS = ["install.ps1", "panel-token.ps1", "start-server.ps1", "register-autostart.ps1", "register-local-catalog.ps1", "diagnose.ps1"];
const FILES = ["manifest.xml", "LICENSE", "NOTICE", "README.md", ".env.example"];

function option(name, fallback) {
  const index = process.argv.indexOf(`--${name}`);
  return index === -1 ? fallback : process.argv[index + 1];
}

function fail(message) {
  console.error(`Комплект не собран: ${message}`);
  process.exit(1);
}

const outRoot = resolve(root, option("out", "bundle"));
const out = join(outRoot, BUNDLE_NAME);
const nodeExe = resolve(option("node", process.execPath));

const panelSource = join(root, "dist");
const serverSource = join(root, "server", "dist");
if (!existsSync(join(panelSource, "taskpane.html")) || !existsSync(join(serverSource, "server.js"))) {
  fail("нет сборки панели или сервера. Сначала выполните npm run check.");
}
if (!existsSync(nodeExe)) fail(`не найден Node: ${nodeExe}`);

// Удаляем только собственный прежний комплект: чужой каталог с тем же именем
// не трогаем.
if (existsSync(out)) {
  if (!existsSync(join(out, "bundle.json"))) fail(`${out} существует и не похож на прежний комплект.`);
  rmSync(out, { recursive: true, force: true });
}

const stamp = new Date().toISOString().replace(/[-:]/g, "").replace(/\.\d+Z$/, "Z");
const hash = createHash("sha256").update(readFileSync(join(serverSource, "server.js"))).digest("hex").slice(0, 8);
const id = `${stamp}-${hash}`;
if (!/^[0-9]{8}T[0-9]{6}Z-[a-f0-9]{8}$/.test(id)) fail(`неверный идентификатор выпуска ${id}`);

cpSync(panelSource, join(out, "releases", id, "panel"), { recursive: true });
cpSync(serverSource, join(out, "server", "releases", id, "dist"), {
  recursive: true,
  filter: (path) => !/\.test\.js(\.map)?$/.test(path)
});
// Тот же формат, что пишет scripts\release.ps1: UTF-8 без BOM, одна строка.
writeFileSync(
  join(out, "releases", "current.json"),
  JSON.stringify({ id, previous: "", activatedAt: new Date().toISOString() })
);

// Сервер читает версию из package.json в корне проекта.
const pkg = JSON.parse(readFileSync(join(root, "package.json"), "utf8"));
writeFileSync(
  join(out, "package.json"),
  JSON.stringify({ name: pkg.name, version: pkg.version, license: pkg.license, private: true }, null, 2) + "\n"
);

for (const file of ["package.json", "package-lock.json"]) {
  cpSync(join(root, "server", file), join(out, "server", file));
}
// Строго по lock-файлу и без install-скриптов пакетов: в комплект попадает
// ровно то, что проверено тестами.
execSync("npm ci --omit=dev --ignore-scripts --no-audit --no-fund", {
  cwd: join(out, "server"),
  stdio: "inherit"
});

for (const file of FILES) cpSync(join(root, file), join(out, file));
// Инструкция пользователю комплекта — в корне, рядом с README разработчика.
cpSync(join(root, "docs", "INSTALL.md"), join(out, "INSTALL.md"));
// Установка двойным щелчком — в корне, чтобы её было видно сразу после распаковки.
cpSync(join(root, "packaging", "Установить.cmd"), join(out, "Установить.cmd"));
mkdirSync(join(out, "scripts"), { recursive: true });
for (const script of SCRIPTS) cpSync(join(root, "scripts", script), join(out, "scripts", script));

// Лицензия Node распространяется вместе с его исполняемым файлом.
mkdirSync(join(out, "node"), { recursive: true });
cpSync(nodeExe, join(out, "node", "node.exe"));
const nodeLicense = join(dirname(nodeExe), "LICENSE");
if (existsSync(nodeLicense)) cpSync(nodeLicense, join(out, "node", "LICENSE"));
else if (process.env.CI) fail(`рядом с ${nodeExe} нет LICENSE Node.js — без него комплект распространять нельзя.`);
else console.warn(`Внимание: рядом с ${nodeExe} нет LICENSE Node.js.`);

const nodeVersion = execFileSync(nodeExe, ["--version"], { encoding: "utf8" }).trim();
let commit = process.env.GITHUB_SHA ?? "";
if (!commit) {
  try { commit = execFileSync("git", ["rev-parse", "HEAD"], { cwd: root, encoding: "utf8" }).trim(); } catch { commit = "unknown"; }
}
const info = { app: APP_ID, version: pkg.version, release: id, node: nodeVersion, commit, builtAt: new Date().toISOString() };
writeFileSync(join(out, "bundle.json"), JSON.stringify(info, null, 2) + "\n");

function sizeOf(path) {
  const stats = statSync(path);
  if (!stats.isDirectory()) return stats.size;
  return readdirSync(path).reduce((sum, name) => sum + sizeOf(join(path, name)), 0);
}

console.log(`Комплект собран: ${relative(root, out) || out}`);
console.log(`  выпуск ${id}, версия ${pkg.version}, Node ${nodeVersion}, коммит ${commit.slice(0, 12)}`);
console.log(`  размер ${(sizeOf(out) / 1024 / 1024).toFixed(1)} МБ`);
