/**
 * Каталог надстроек через WebDAV — «общая папка» без прав администратора.
 *
 * Excel показывает надстройку из каталога надёжных надстроек, а каталог —
 * это сетевая папка. Прежде ею была служебная \\localhost\C$\…, открытая
 * только администраторам. Windows умеет открывать как сетевую папку и адрес
 * WebDAV: \\localhost@3080\catalog читает служба WebClient, и это доступно
 * любому пользователю. Проверено 29.09.2026: Excel нашёл надстройку в
 * «Общей папке» и открыл панель.
 *
 * Отдаём, только когда регистрация выбрала этот путь: скрипт регистрации
 * кладёт метку server/catalog-dav в закрытую папку надстройки, если C$
 * недоступна. Без метки — 404 на всё: при регистрации через C$ манифест с
 * токеном другим учётным записям этого компьютера по WebDAV не виден.
 *
 * Отдаём только файлы из папки catalog, только чтение, только на 127.0.0.1.
 * В манифесте — токен панели, поэтому Host должен быть своим: иначе страница
 * в браузере могла бы подменой DNS (rebinding) стать «своей» для этого порта
 * и прочитать манифест. Заголовков CORS нет — чужие страницы ответа не видят.
 */

import http from "node:http";
import { readdirSync, readFileSync, statSync, type Stats } from "node:fs";
import { basename, join } from "node:path";

export const CATALOG_DAV_PREFIX = "/catalog";
export const DEFAULT_CATALOG_DAV_PORT = 3080;

function xmlEscape(value: string): string {
  return value.replace(/[<>&"']/g, (c) => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;", "\"": "&quot;", "'": "&apos;" })[c]!);
}

function davEntry(href: string, stat: Stats): string {
  const name = xmlEscape(href.split("/").filter(Boolean).pop() ?? "");
  const kind = stat.isDirectory()
    ? "<D:resourcetype><D:collection/></D:resourcetype>"
    : `<D:resourcetype/><D:getcontentlength>${stat.size}</D:getcontentlength><D:getcontenttype>text/xml</D:getcontenttype>`;
  return `<D:response><D:href>${xmlEscape(href)}</D:href><D:propstat><D:prop>` +
    `<D:displayname>${name}</D:displayname>${kind}` +
    `<D:getlastmodified>${stat.mtime.toUTCString()}</D:getlastmodified>` +
    `<D:creationdate>${stat.birthtime.toISOString()}</D:creationdate>` +
    `</D:prop><D:status>HTTP/1.1 200 OK</D:status></D:propstat></D:response>`;
}

/** Только файлы прямо в папке каталога: без подпапок, скрытых файлов и выхода наверх. */
function catalogFiles(root: string): string[] {
  return readdirSync(root, { withFileTypes: true })
    .filter((item) => item.isFile() && !item.name.startsWith("."))
    .map((item) => item.name);
}

export function catalogDavHandler(root: string, port: number, enabled: () => boolean = () => true): http.RequestListener {
  const allowedHosts = new Set([`localhost:${port}`, `127.0.0.1:${port}`]);
  return (req, res) => {
    if (!enabled()) {
      res.writeHead(404);
      res.end();
      return;
    }
    res.setHeader("DAV", "1");
    res.setHeader("MS-Author-Via", "DAV");
    res.setHeader("Cache-Control", "no-store");
    if (!allowedHosts.has(String(req.headers.host ?? "").toLowerCase())) {
      res.writeHead(403);
      res.end();
      return;
    }
    if (req.method === "OPTIONS") {
      res.setHeader("Allow", "OPTIONS, GET, HEAD, PROPFIND");
      res.end();
      return;
    }

    let path: string;
    try {
      path = decodeURIComponent(String(req.url ?? "/").split("?")[0]);
    } catch {
      res.writeHead(400);
      res.end();
      return;
    }
    // Корень сервера — чтобы клиент Windows нашёл папку catalog; внутри неё — файлы.
    const atRoot = path === "/" || path === "";
    const inCatalog = path === CATALOG_DAV_PREFIX || path.startsWith(`${CATALOG_DAV_PREFIX}/`);
    if (!atRoot && !inCatalog) {
      res.writeHead(404);
      res.end();
      return;
    }
    const name = inCatalog ? path.slice(CATALOG_DAV_PREFIX.length).replace(/^\/+/, "").replace(/\/+$/, "") : "";

    let files: string[];
    let rootStat: Stats;
    try {
      files = catalogFiles(root);
      rootStat = statSync(root);
    } catch {
      res.writeHead(404);
      res.end();
      return;
    }
    if (name && (name !== basename(name) || !files.includes(name))) {
      res.writeHead(404);
      res.end();
      return;
    }

    if (req.method === "PROPFIND") {
      const depth = String(req.headers.depth ?? "1");
      let body: string;
      if (atRoot) {
        body = davEntry("/", rootStat) + (depth === "0" ? "" : davEntry(`${CATALOG_DAV_PREFIX}/`, rootStat));
      } else if (!name) {
        body = davEntry(`${CATALOG_DAV_PREFIX}/`, rootStat);
        if (depth !== "0") for (const file of files) body += davEntry(`${CATALOG_DAV_PREFIX}/${file}`, statSync(join(root, file)));
      } else {
        body = davEntry(`${CATALOG_DAV_PREFIX}/${name}`, statSync(join(root, name)));
      }
      res.writeHead(207, { "Content-Type": "text/xml; charset=utf-8" });
      res.end(`<?xml version="1.0" encoding="utf-8"?><D:multistatus xmlns:D="DAV:">${body}</D:multistatus>`);
      return;
    }

    if ((req.method === "GET" || req.method === "HEAD") && name) {
      const file = join(root, name);
      const data = readFileSync(file);
      res.writeHead(200, {
        "Content-Type": "text/xml; charset=utf-8",
        "Content-Length": data.length,
        "Last-Modified": statSync(file).mtime.toUTCString()
      });
      res.end(req.method === "HEAD" ? undefined : data);
      return;
    }

    res.setHeader("Allow", "OPTIONS, GET, HEAD, PROPFIND");
    res.writeHead(405);
    res.end();
  };
}

/** Порт WebDAV-каталога из server/.env; "off" выключает его. */
export function catalogDavPort(value: string | undefined): number | null {
  const text = (value ?? "").trim();
  if (/^(off|0|false|no)$/i.test(text)) return null;
  if (!text) return DEFAULT_CATALOG_DAV_PORT;
  const port = Number(text);
  return Number.isInteger(port) && port > 0 && port < 65536 ? port : DEFAULT_CATALOG_DAV_PORT;
}

export function startCatalogDav(root: string, port: number, enabled: () => boolean, onError: (error: NodeJS.ErrnoException) => void): http.Server {
  const server = http.createServer(catalogDavHandler(root, port, enabled));
  server.on("error", onError);
  server.listen(port, "127.0.0.1");
  return server;
}
