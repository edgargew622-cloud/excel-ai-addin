/**
 * Безопасное чтение страницы из интернета (этап 8, 8.7).
 *
 * Сервер работает на компьютере пользователя, поэтому ссылка из интернета
 * (от модели, со страницы, из файла) не должна заставить его обратиться к
 * самому компьютеру, роутеру или устройствам домашней сети (SSRF). Правила:
 * - только http и https, только порты 80 и 443, без логина в адресе;
 * - адрес проверяется на каждом соединении — в собственной функции поиска
 *   IP, через которую идёт подключение, а не заранее: иначе имя могло бы
 *   сначала показать публичный адрес, а при подключении — внутренний
 *   (DNS rebinding);
 * - перенаправления — вручную, каждое снова через те же проверки;
 * - предел времени и размера ответа.
 */

import { lookup as dnsLookup, type LookupAddress } from "node:dns";
import { request as httpRequest, type IncomingMessage } from "node:http";
import { request as httpsRequest } from "node:https";
import { isIP } from "node:net";
import { brotliDecompressSync, gunzipSync, inflateSync } from "node:zlib";

export const FETCH_LIMITS = { bytes: 3 * 1024 * 1024, timeoutMs: 15_000, redirects: 5 };

export class WebFetchError extends Error {}

function ipv4Private(ip: string): boolean {
  const [a, b] = ip.split(".").map(Number);
  return a === 0 || a === 10 || a === 127 || a >= 224 ||
    (a === 100 && b >= 64 && b <= 127) ||
    (a === 169 && b === 254) ||
    (a === 172 && b >= 16 && b <= 31) ||
    (a === 192 && b === 168) ||
    (a === 192 && b === 0) ||
    (a === 198 && (b === 18 || b === 19));
}

/** Публичный ли адрес: не петля, не частная сеть, не служебный диапазон. */
export function isPublicAddress(ip: string): boolean {
  const kind = isIP(ip);
  if (kind === 4) return !ipv4Private(ip);
  if (kind !== 6) return false;
  const lower = ip.toLowerCase();
  const mapped = /^(?:::ffff:|64:ff9b::)(\d+\.\d+\.\d+\.\d+)$/.exec(lower);
  if (mapped) return !ipv4Private(mapped[1]);
  if (/^::ffff:[0-9a-f]{1,4}:[0-9a-f]{1,4}$/.test(lower)) {
    const [high, low] = lower.slice(7).split(":").map((part) => parseInt(part, 16));
    return !ipv4Private(`${high >> 8}.${high & 255}.${low >> 8}.${low & 255}`);
  }
  if (lower === "::" || lower === "::1") return false;
  const first = parseInt(lower.split(":")[0] || "0", 16);
  // fc00::/7 — частные, fe80::/10 — локальные в сегменте, ff00::/8 — групповые.
  if ((first & 0xfe00) === 0xfc00 || (first & 0xffc0) === 0xfe80 || (first & 0xff00) === 0xff00) return false;
  return true;
}

/** Проверка адреса до запроса: схема, порт, логин, IP-литерал. */
export function checkUrl(raw: string): URL {
  let url: URL;
  try { url = new URL(raw); } catch { throw new WebFetchError(`«${raw}» — не адрес страницы.`); }
  if (url.protocol !== "http:" && url.protocol !== "https:") throw new WebFetchError(`Открываются только страницы http и https, а не ${url.protocol}`);
  if (url.username || url.password) throw new WebFetchError("Адрес с логином и паролем не открывается.");
  if (url.port && url.port !== "80" && url.port !== "443") throw new WebFetchError(`Порт ${url.port} не открывается: только обычные 80 и 443.`);
  const host = url.hostname.replace(/^\[|\]$/g, "");
  if (isIP(host) && !isPublicAddress(host)) throw new WebFetchError("Адрес ведёт на этот компьютер или во внутреннюю сеть — такие страницы не открываются.");
  if (/^(localhost|.*\.localhost|.*\.local|.*\.internal|.*\.lan|.*\.home)$/i.test(host)) throw new WebFetchError("Адрес ведёт во внутреннюю сеть — такие страницы не открываются.");
  return url;
}

/** Поиск IP с проверкой: подключение идёт только к публичному адресу. */
function guardedLookup(hostname: string, options: any, callback: (...args: any[]) => void) {
  dnsLookup(hostname, { all: true, family: options?.family ?? 0 }, (error, addresses: LookupAddress[]) => {
    if (error) return callback(error);
    const blocked = addresses.find((item) => !isPublicAddress(item.address));
    if (blocked || !addresses.length) {
      return callback(new WebFetchError("Адрес ведёт на этот компьютер или во внутреннюю сеть — такие страницы не открываются."));
    }
    if (options?.all) return callback(null, addresses);
    callback(null, addresses[0].address, addresses[0].family);
  });
}

export interface FetchedPage {
  url: string;
  status: number;
  contentType: string;
  body: Buffer;
}

function once(url: URL): Promise<IncomingMessage> {
  return new Promise((resolve, reject) => {
    const request = (url.protocol === "https:" ? httpsRequest : httpRequest)(url, {
      method: "GET",
      lookup: guardedLookup as any,
      headers: {
        "User-Agent": "am.AI for Excel (+https://github.com/edgargew622-cloud/excel-ai-addin)",
        "Accept": "text/html,application/xhtml+xml,text/plain,application/pdf;q=0.9,*/*;q=0.5",
        "Accept-Language": "ru,en;q=0.8",
        "Accept-Encoding": "gzip, deflate, br"
      },
      timeout: FETCH_LIMITS.timeoutMs
    }, resolve);
    request.on("timeout", () => request.destroy(new WebFetchError(`Страница не ответила за ${FETCH_LIMITS.timeoutMs / 1000} секунд.`)));
    request.on("error", reject);
    request.end();
  });
}

export async function safeFetch(raw: string): Promise<FetchedPage> {
  let url = checkUrl(raw);
  for (let hop = 0; hop <= FETCH_LIMITS.redirects; hop++) {
    let response: IncomingMessage;
    try {
      response = await once(url);
    } catch (error: any) {
      if (error instanceof WebFetchError) throw error;
      throw new WebFetchError(`Страница не открылась: ${error?.code === "ENOTFOUND" ? "такого сайта нет" : error?.message ?? error}.`);
    }
    const status = response.statusCode ?? 0;
    if (status >= 300 && status < 400 && response.headers.location) {
      response.resume();
      url = checkUrl(new URL(response.headers.location, url).href);
      continue;
    }
    const chunks: Buffer[] = [];
    let size = 0;
    await new Promise<void>((resolve, reject) => {
      response.on("data", (chunk: Buffer) => {
        size += chunk.length;
        if (size > FETCH_LIMITS.bytes) { response.destroy(); reject(new WebFetchError(`Страница больше ${FETCH_LIMITS.bytes / 1024 / 1024} МБ — не читается.`)); return; }
        chunks.push(chunk);
      });
      response.on("end", resolve);
      response.on("error", reject);
    });
    let body = Buffer.concat(chunks);
    const encoding = String(response.headers["content-encoding"] ?? "").toLowerCase();
    try {
      if (encoding === "gzip") body = gunzipSync(body, { maxOutputLength: FETCH_LIMITS.bytes * 4 });
      else if (encoding === "br") body = brotliDecompressSync(body, { maxOutputLength: FETCH_LIMITS.bytes * 4 });
      else if (encoding === "deflate") body = inflateSync(body, { maxOutputLength: FETCH_LIMITS.bytes * 4 });
    } catch {
      throw new WebFetchError("Страница пришла сжатой и не распаковалась.");
    }
    if (status >= 400) throw new WebFetchError(`Сайт ответил ошибкой ${status}${status === 403 || status === 429 ? " — возможно, он не пускает программы" : ""}.`);
    return { url: url.href, status, contentType: String(response.headers["content-type"] ?? ""), body };
  }
  throw new WebFetchError(`Больше ${FETCH_LIMITS.redirects} перенаправлений — страница не открыта.`);
}
