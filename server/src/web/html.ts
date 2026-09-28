/**
 * Текст страницы из HTML (этап 8, 8.7): заголовок и содержимое без скриптов,
 * стилей и разметки. Таблицы — строками «ячейка | ячейка», чтобы числа из
 * финансовых таблиц не слипались.
 */

/** Кодировка из заголовка или <meta charset>; по умолчанию UTF-8. */
export function decodeHtml(body: Buffer, contentType: string): string {
  const head = body.subarray(0, 4096).toString("latin1");
  const charset = (/charset=["']?([\w-]+)/i.exec(contentType)?.[1] ?? /<meta[^>]+charset=["']?([\w-]+)/i.exec(head)?.[1] ?? "utf-8").toLowerCase();
  try {
    return new TextDecoder(charset === "cp1251" ? "windows-1251" : charset).decode(body);
  } catch {
    return new TextDecoder("utf-8").decode(body);
  }
}

const ENTITIES: Record<string, string> = { nbsp: " ", amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", laquo: "«", raquo: "»", mdash: "—", ndash: "–", hellip: "…", copy: "©", reg: "®", euro: "€", rub: "₽" };

function entities(text: string): string {
  return text
    .replace(/&#x([0-9a-f]+);/gi, (_, hex) => { try { return String.fromCodePoint(parseInt(hex, 16)); } catch { return ""; } })
    .replace(/&#(\d+);/g, (_, dec) => { try { return String.fromCodePoint(Number(dec)); } catch { return ""; } })
    .replace(/&([a-z]+);/gi, (whole, name) => ENTITIES[name.toLowerCase()] ?? whole);
}

export function htmlToText(html: string): { title: string; text: string } {
  const title = entities(/<title[^>]*>([\s\S]*?)<\/title>/i.exec(html)?.[1] ?? "").replace(/\s+/g, " ").trim();
  let body = html
    .replace(/<!--[\s\S]*?-->/g, " ")
    .replace(/<(script|style|noscript|svg|template|iframe|head|canvas|button|select)\b[\s\S]*?<\/\1>/gi, " ");
  // Переводы строк исходника ничего не значат: строку таблицы ЦБ, записанную
  // по ячейке на строку, иначе было бы не собрать (замер 28.09.2026).
  body = body
    .replace(/\s+/g, " ")
    .replace(/<\/(td|th)>/gi, " | ")
    .replace(/<\/tr>/gi, "\n")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/?(p|div|section|article|li|ul|ol|h[1-6]|table|tbody|thead|tr|header|footer|nav|aside|main|blockquote|pre|dd|dt|dl)\b[^>]*>/gi, "\n")
    .replace(/<[^>]+>/g, " ");
  const text = entities(body)
    .split("\n")
    .map((line) => line.replace(/[ \t ]+/g, " ").replace(/\s*\|\s*$/, "").trim())
    .filter((line) => line && line !== "|")
    .join("\n")
    .replace(/\n{3,}/g, "\n\n");
  return { title, text };
}
