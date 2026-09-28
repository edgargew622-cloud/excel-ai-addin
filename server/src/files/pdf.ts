/**
 * PDF (этап 8, 8.6): текст по страницам через pdf.js (Mozilla, Apache 2.0).
 *
 * PDF — самый опасный из форматов: в нём бывают скрипты и шрифты-ловушки.
 * Поэтому pdf.js работает без исполнения кода (isEvalSupported: false —
 * закрывает CVE-2024-4367), без загрузки шрифтов и системных шрифтов; нужен
 * только текстовый слой. Скан без текстового слоя не распознаётся — это
 * называется, а не выдаётся пустым ответом.
 */

import { createRequire } from "node:module";
import { dirname, join } from "node:path";
import { pathToFileURL } from "node:url";
import { FILE_LIMITS, limitText, type ParsedFile } from "./types.js";
import { FileParseError } from "./zip.js";

export async function parsePdf(buffer: Buffer, name: string): Promise<ParsedFile> {
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  // Таблицы символов для шрифтов CID (китайский, японский, корейский) — из самого пакета.
  const root = dirname(createRequire(import.meta.url).resolve("pdfjs-dist/package.json"));
  const folder = (name: string) => pathToFileURL(join(root, name)).href + "/";
  const warnings: string[] = [];
  let document: any;
  const task = pdfjs.getDocument({
    data: new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength),
    isEvalSupported: false,
    disableFontFace: true,
    useSystemFonts: false,
    enableXfa: false,
    stopAtErrors: false,
    verbosity: 0,
    cMapUrl: folder("cmaps"),
    cMapPacked: true,
    standardFontDataUrl: folder("standard_fonts")
  } as any);
  try {
    document = await task.promise;
  } catch (error: any) {
    await task.destroy();
    const message = String(error?.name ?? "") === "PasswordException"
      ? "PDF защищён паролем: без пароля его не прочитать."
      : `PDF не открылся: ${error?.message ?? error}. Возможно, файл повреждён.`;
    throw new FileParseError(message);
  }
  try {
    const total = document.numPages as number;
    const count = Math.min(total, FILE_LIMITS.pages);
    if (total > count) warnings.push(`В PDF ${total} страниц: разобраны первые ${count}.`);
    const pages: string[] = [];
    let empty = 0;
    for (let number = 1; number <= count; number++) {
      const page = await document.getPage(number);
      const content = await page.getTextContent();
      let text = "";
      for (const item of content.items as any[]) {
        if (typeof item.str !== "string") continue;
        text += item.str;
        if (item.hasEOL) text += "\n";
      }
      text = text.replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
      if (!text) empty += 1;
      pages.push(text);
      page.cleanup();
    }
    if (empty === pages.length) {
      throw new FileParseError("В PDF нет текстового слоя — похоже на скан. Распознавание текста на картинках панель не делает.");
    }
    if (empty) warnings.push(`Страниц без текста (картинки или сканы): ${empty} — их содержимое не прочитано.`);
    warnings.push("Таблицы в PDF приходят текстом: столбцы не размечены, перенос в книгу — после разбора текста.");
    return { kind: "pdf", name, size: buffer.length, tables: [], text: limitText(pages, warnings), warnings };
  } finally {
    await task.destroy();
  }
}
