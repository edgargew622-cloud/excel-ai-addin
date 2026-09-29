/**
 * fetch для модели на этом компьютере — без предела на ожидание ответа.
 *
 * Встроенный fetch ждёт заголовков не дольше пяти минут. Ollama присылает их
 * только прочитав весь запрос, а на процессоре или встроенной видеокарте это
 * 20 тысяч токенов со скоростью 13–40 в секунду — от десяти минут. Обычный
 * fetch обрывал такой шаг с «fetch failed». Здесь тот же запрос через
 * node:http: заголовков ждём сколько нужно, время задачи ограничивает панель
 * (taskBudgetMinutes), отмена — через тот же AbortSignal.
 */

import http from "node:http";
import https from "node:https";
import { Readable } from "node:stream";

export function fetchWithoutHeaderTimeout(url: string, init: RequestInit = {}): Promise<Response> {
  return new Promise((resolve, reject) => {
    const target = new URL(url);
    const signal = init.signal ?? undefined;
    if (signal?.aborted) return reject(signal.reason ?? new DOMException("Aborted", "AbortError"));

    const headers: Record<string, string> = {};
    new Headers(init.headers).forEach((value, key) => { headers[key] = value; });
    const request = (target.protocol === "https:" ? https : http).request(target, {
      method: init.method ?? "GET",
      headers
    }, (response) => {
      const responseHeaders = new Headers();
      for (const [key, value] of Object.entries(response.headers)) {
        if (value !== undefined) responseHeaders.set(key, Array.isArray(value) ? value.join(", ") : value);
      }
      const status = response.statusCode ?? 502;
      resolve(new Response(Readable.toWeb(response) as ReadableStream, {
        status,
        statusText: response.statusMessage,
        headers: responseHeaders
      }));
    });

    const onAbort = () => request.destroy(signal?.reason ?? new DOMException("Aborted", "AbortError"));
    signal?.addEventListener("abort", onAbort, { once: true });
    request.on("close", () => signal?.removeEventListener("abort", onAbort));
    // Ошибка соединения — как у fetch: TypeError с причиной, чтобы повтор
    // при обрыве (server.ts) узнавал её по cause.code.
    request.on("error", (error) => reject(signal?.aborted ? error : Object.assign(new TypeError("fetch failed"), { cause: error })));
    if (init.body !== undefined && init.body !== null) request.write(init.body as string);
    request.end();
  });
}
