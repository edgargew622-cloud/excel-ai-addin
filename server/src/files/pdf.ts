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
import { buildSheetLayout, type PdfPageRaw, type PdfSegment, type PdfText, type SheetLayout } from "./pdfLayout.js";

const LAYOUT_PAGES = 30;
const BOLD_FONT = /bold|black|heavy|semibold|demibold|extrabold/i;

type Matrix = [number, number, number, number, number, number];
const multiply = (m: Matrix, n: Matrix): Matrix => [
  m[0] * n[0] + m[2] * n[1], m[1] * n[0] + m[3] * n[1],
  m[0] * n[2] + m[2] * n[3], m[1] * n[2] + m[3] * n[3],
  m[0] * n[4] + m[2] * n[5] + m[4], m[1] * n[4] + m[3] * n[5] + m[5]
];

/**
 * Текст и линии одной страницы в координатах «от левого верхнего угла».
 * Линии — из контуров операторов рисования: обведённые отрезки вдоль осей и
 * тонкие залитые прямоугольники (так рисуют рамки Excel и Word); крупные
 * заливки — фон, не рамки. Преобразования страницы (cm, форм-объекты)
 * учитываются.
 */
async function pageGeometry(pdfjs: any, document: any, number: number): Promise<PdfPageRaw> {
  const page = await document.getPage(number);
  try {
    const [vx0, vy0, vx1, vy1] = page.view as number[];
    const height = vy1 - vy0;
    const ops = await page.getOperatorList();
    const content = await page.getTextContent();
    const O = pdfjs.OPS;
    const toTop = (x: number, y: number) => ({ x: x - vx0, y: height - (y - vy0) });

    const segments: PdfSegment[] = [];
    let ctm: Matrix = [1, 0, 0, 1, 0, 0];
    let lineWidth = 1;
    // Белые и почти белые линии — фон, а не рамки (путевой лист 01.10.2026:
    // белые обводки фоновых прямоугольников давали ложные рамки).
    let strokeLight = false;
    let fillLight = false;
    const light = (color: unknown) => {
      const hex = /^#?([0-9a-f]{6})$/i.exec(String(Array.isArray(color) ? color[0] : color ?? ""));
      if (!hex) return false;
      const n = parseInt(hex[1], 16);
      return [n >> 16, (n >> 8) & 255, n & 255].every((part) => part >= 0xe6);
    };
    const stack: Array<{ ctm: Matrix; lineWidth: number; strokeLight: boolean; fillLight: boolean }> = [];
    const apply = (x: number, y: number) => toTop(ctm[0] * x + ctm[2] * y + ctm[4], ctm[1] * x + ctm[3] * y + ctm[5]);
    const scale = () => Math.sqrt(Math.abs(ctm[0] * ctm[3] - ctm[1] * ctm[2])) || 1;
    const addEdge = (a: { x: number; y: number }, b: { x: number; y: number }, width: number) => {
      if (Math.abs(a.x - b.x) < 0.8 || Math.abs(a.y - b.y) < 0.8) {
        segments.push({ x1: Math.min(a.x, b.x), y1: Math.min(a.y, b.y), x2: Math.max(a.x, b.x), y2: Math.max(a.y, b.y), width });
      }
    };
    // Текст — из самих команд вывода (Tj/TJ): pdf.js в getTextContent склеивает
    // надписи одной строки и заменяет просвет между колонками одним пробелом,
    // а здесь у каждой надписи своё точное место (01.10.2026, «Кол-во  Ед.»).
    const opTexts: PdfText[] = [];
    let skippedRotated = 0;
    type TextState = { font: string; size: number; charSpacing: number; wordSpacing: number; hScale: number; leading: number; rise: number };
    let ts: TextState = { font: "", size: 10, charSpacing: 0, wordSpacing: 0, hScale: 1, leading: 0, rise: 0 };
    let tm: Matrix = [1, 0, 0, 1, 0, 0];
    let lm: Matrix = [1, 0, 0, 1, 0, 0];
    const textStack: TextState[] = [];
    const fontInfo = new Map<string, { scale: number; ascent: number; descent: number; bold: boolean }>();
    const fontOf = (id: string) => {
      if (!fontInfo.has(id)) {
        let font: any = null;
        try { font = page.commonObjs.get(id); } catch { /* не загружен */ }
        fontInfo.set(id, {
          scale: Array.isArray(font?.fontMatrix) || ArrayBuffer.isView(font?.fontMatrix) ? Number(font.fontMatrix[0]) || 0.001 : 0.001,
          ascent: Number(font?.ascent) || 0.8,
          descent: Number(font?.descent) || -0.2,
          bold: BOLD_FONT.test(String(font?.name ?? "")) || Boolean(font?.bold) || Boolean(font?.black)
        });
      }
      return fontInfo.get(id)!;
    };
    const moveText = (tx: number, ty: number) => {
      lm = [lm[0], lm[1], lm[2], lm[3], lm[4] + tx * lm[0] + ty * lm[2], lm[5] + tx * lm[1] + ty * lm[3]];
      tm = lm;
    };
    const showGlyphs = (glyphs: any[]) => {
      if (!Array.isArray(glyphs)) return;
      const font = fontOf(ts.font);
      const full = multiply(ctm, tm);
      if (Math.abs(full[1]) > 0.01 * Math.abs(full[0]) || Math.abs(full[2]) > 0.01 * Math.abs(full[3])) {
        skippedRotated += 1;
      }
      const sizeOnPage = Math.abs(ts.size * full[3]) || Math.abs(ts.size * full[0]) || ts.size;
      let x = 0;
      let run: { str: string; start: number; end: number } | null = null;
      const flush = () => {
        if (run && run.str.trim()) {
          const a = toTop(full[0] * run.start + full[2] * ts.rise + full[4], full[1] * run.start + full[3] * ts.rise + full[5]);
          const b = toTop(full[0] * run.end + full[2] * ts.rise + full[4], full[1] * run.end + full[3] * ts.rise + full[5]);
          const lead = run.str.length - run.str.trimStart().length;
          opTexts.push({
            str: run.str.trim(),
            x0: Math.min(a.x, b.x) + (lead ? (Math.abs(b.x - a.x) * lead) / run.str.length : 0),
            x1: Math.max(a.x, b.x),
            top: a.y - font.ascent * sizeOnPage,
            bottom: a.y - font.descent * sizeOnPage,
            size: sizeOnPage,
            bold: font.bold
          });
        }
        run = null;
      };
      for (const glyph of glyphs) {
        if (typeof glyph === "number") {
          const shift = (-glyph / 1000) * ts.size * ts.hScale;
          if (shift > 0.5 * ts.size) flush();
          x += shift;
          continue;
        }
        if (!glyph) continue;
        const space = glyph.isSpace || glyph.unicode === " ";
        const advance = ((Number(glyph.width) || 0) * font.scale * ts.size + ts.charSpacing + (space ? ts.wordSpacing : 0)) * ts.hScale;
        if (space && advance > 0.5 * ts.size) { flush(); x += advance; continue; }
        if (!run) run = { str: "", start: x, end: x };
        run.str += String(glyph.unicode ?? "");
        x += advance;
        run.end = x;
      }
      flush();
      tm = [tm[0], tm[1], tm[2], tm[3], tm[4] + x * tm[0], tm[5] + x * tm[1]];
    };
    const strokes = new Set([O.stroke, O.closeStroke, O.fillStroke, O.eoFillStroke, O.closeFillStroke, O.closeEOFillStroke]);
    const fills = new Set([O.fill, O.eoFill]);
    for (let i = 0; i < ops.fnArray.length; i++) {
      const fn = ops.fnArray[i];
      const args = ops.argsArray[i];
      if (fn === O.save) { stack.push({ ctm, lineWidth, strokeLight, fillLight }); textStack.push({ ...ts }); }
      else if (fn === O.restore) { ({ ctm, lineWidth, strokeLight, fillLight } = stack.pop() ?? { ctm, lineWidth, strokeLight, fillLight }); ts = textStack.pop() ?? ts; }
      else if (fn === O.setStrokeRGBColor) strokeLight = light(args);
      else if (fn === O.setFillRGBColor) fillLight = light(args);
      else if (fn === O.beginText) { tm = [1, 0, 0, 1, 0, 0]; lm = tm; }
      else if (fn === O.setFont) { ts.font = String(args[0]); ts.size = Number(args[1]) || ts.size; }
      else if (fn === O.setCharSpacing) ts.charSpacing = Number(args[0]) || 0;
      else if (fn === O.setWordSpacing) ts.wordSpacing = Number(args[0]) || 0;
      else if (fn === O.setHScale) ts.hScale = (Number(args[0]) || 100) / 100;
      else if (fn === O.setLeading) ts.leading = Number(args[0]) || 0;
      else if (fn === O.setTextRise) ts.rise = Number(args[0]) || 0;
      else if (fn === O.setTextMatrix) { tm = Array.from(args as ArrayLike<number>).slice(0, 6) as Matrix; lm = tm; }
      else if (fn === O.moveText) moveText(Number(args[0]), Number(args[1]));
      else if (fn === O.setLeadingMoveText) { ts.leading = -Number(args[1]); moveText(Number(args[0]), Number(args[1])); }
      else if (fn === O.nextLine) moveText(0, -ts.leading);
      else if (fn === O.showText || fn === O.showSpacedText) showGlyphs(args[0]);
      else if (fn === O.nextLineShowText) { moveText(0, -ts.leading); showGlyphs(args[0]); }
      else if (fn === O.nextLineSetSpacingShowText) { ts.wordSpacing = Number(args[0]) || 0; ts.charSpacing = Number(args[1]) || 0; moveText(0, -ts.leading); showGlyphs(args[2]); }
      else if (fn === O.transform) ctm = multiply(ctm, args as Matrix);
      else if (fn === O.setLineWidth) lineWidth = Number(args[0]) || 1;
      else if (fn === O.paintFormXObjectBegin) { stack.push({ ctm, lineWidth, strokeLight, fillLight }); if (Array.isArray(args?.[0]) || ArrayBuffer.isView(args?.[0])) ctm = multiply(ctm, Array.from(args[0] as ArrayLike<number>) as Matrix); }
      else if (fn === O.paintFormXObjectEnd) ({ ctm, lineWidth, strokeLight, fillLight } = stack.pop() ?? { ctm, lineWidth, strokeLight, fillLight });
      else if (fn === O.constructPath) {
        const paint = args[0];
        const stroke = strokes.has(paint) && !strokeLight;
        const fill = (fills.has(paint) || (strokes.has(paint) && paint !== O.stroke && paint !== O.closeStroke)) && !fillLight;
        if (!stroke && !fill) continue;
        const width = Math.max(0.1, lineWidth * scale());
        for (const data of (Array.isArray(args[1]) ? args[1] : [args[1]]) as ArrayLike<number>[]) {
          if (!data) continue;
          let start: { x: number; y: number } | null = null;
          let current: { x: number; y: number } | null = null;
          let points: Array<{ x: number; y: number }> = [];
          const flushFill = () => {
            if (fill && points.length >= 3) {
              const xs = points.map((p) => p.x);
              const ys = points.map((p) => p.y);
              const [x1, x2, y1, y2] = [Math.min(...xs), Math.max(...xs), Math.min(...ys), Math.max(...ys)];
              const [w, h] = [x2 - x1, y2 - y1];
              if (Math.min(w, h) <= 2.5 && Math.max(w, h) >= 3) {
                if (w >= h) segments.push({ x1, x2, y1: (y1 + y2) / 2, y2: (y1 + y2) / 2, width: Math.max(0.1, h) });
                else segments.push({ x1: (x1 + x2) / 2, x2: (x1 + x2) / 2, y1, y2, width: Math.max(0.1, w) });
              }
            }
            points = [];
          };
          for (let k = 0; k < data.length;) {
            const op = data[k];
            if (op === 0) { flushFill(); start = current = apply(data[k + 1], data[k + 2]); points.push(current); k += 3; }
            else if (op === 1) { const next = apply(data[k + 1], data[k + 2]); if (stroke && current) addEdge(current, next, width); current = next; points.push(next); k += 3; }
            else if (op === 2) { current = apply(data[k + 5], data[k + 6]); points.push(current); k += 7; }
            else if (op === 3) { current = apply(data[k + 3], data[k + 4]); points.push(current); k += 5; }
            else if (op === 4) { if (stroke && current && start) addEdge(current, start, width); current = start; k += 1; }
            else break;
          }
          flushFill();
        }
      }
    }

    const texts: PdfText[] = [];
    const styles = content.styles as Record<string, { ascent?: number; descent?: number }>;
    for (const item of content.items as any[]) {
      if (typeof item.str !== "string" || !item.str.trim()) continue;
      const t = item.transform as number[];
      const size = Math.hypot(t[2], t[3]) || item.height || 10;
      const style = styles[item.fontName] ?? {};
      const ascent = style.ascent ?? 0.8;
      const descent = style.descent ?? -0.2;
      const baseline = toTop(t[4], t[5]);
      let bold = false;
      try { bold = BOLD_FONT.test(String(page.commonObjs.get(item.fontName)?.name ?? "")); } catch { /* шрифт не загружен */ }
      texts.push({
        str: item.str,
        x0: baseline.x,
        x1: baseline.x + (item.width || size * 0.5 * item.str.length),
        top: baseline.y - ascent * size,
        bottom: baseline.y - descent * size,
        size,
        bold
      });
    }
    if (opTexts.length && !skippedRotated) return { width: vx1 - vx0, height, texts: opTexts, segments };
    return { width: vx1 - vx0, height, texts: opTexts.length >= texts.length ? opTexts : texts, segments };
  } finally {
    page.cleanup();
  }
}

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
    warnings.push("Таблицы в PDF приходят текстом; перенести файл в книгу «как есть» — со столбцами, объединениями и рамками — может import_file_layout.");
    // Разметка «как есть» (pdfLayout.ts): первые LAYOUT_PAGES страниц.
    let layout: SheetLayout | undefined;
    let layoutError: string | undefined;
    try {
      const raw: PdfPageRaw[] = [];
      for (let number = 1; number <= Math.min(count, LAYOUT_PAGES); number++) raw.push(await pageGeometry(pdfjs, document, number));
      layout = buildSheetLayout(raw);
      if (count > LAYOUT_PAGES) layout.warnings.push(`Как есть переносятся первые ${LAYOUT_PAGES} страниц из ${count}.`);
    } catch (error: any) {
      layoutError = String(error?.message ?? error);
    }
    return { kind: "pdf", name, size: buffer.length, tables: [], text: limitText(pages, warnings), warnings, ...(layout ? { layout } : {}), ...(layoutError ? { layoutError } : {}) };
  } finally {
    await task.destroy();
  }
}
