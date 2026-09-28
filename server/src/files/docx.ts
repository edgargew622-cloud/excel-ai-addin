/**
 * DOCX (этап 8, 8.6): текст документа и таблицы.
 *
 * Абзацы идут в текст по порядку; на месте таблицы в тексте — пометка
 * «[Таблица N]», а сама таблица — отдельно, её можно перенести в книгу.
 * Вложенная таблица ячейки попадает в текст этой ячейки. Колонтитулы,
 * сноски и надписи не читаются — это сказано в предупреждениях.
 */

import { chunkText, FILE_LIMITS, limitText, type CellValue, type FileTable, type ParsedFile } from "./types.js";
import { FileParseError, xmlText, zipReader } from "./zip.js";

/** Текст одного абзаца <w:p>: прогоны <w:t>, табуляции и переносы строк. */
function paragraphText(xml: string): string {
  let text = "";
  for (const match of xml.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>|<w:t\b[^>]*\/>|<w:(tab|br|cr)\b[^>]*\/>/g)) {
    if (match[2] === "tab") text += "\t";
    else if (match[2]) text += "\n";
    else text += xmlText(match[1] ?? "");
  }
  return text;
}

/**
 * Тело документа по верхнему уровню: абзацы и таблицы в порядке следования.
 * Таблицы ищутся по вложенности тегов, а не регулярным выражением целиком:
 * у вложенной таблицы свой </w:tbl>.
 */
function topLevelBlocks(body: string): ({ kind: "p"; xml: string } | { kind: "tbl"; xml: string })[] {
  const blocks: ({ kind: "p"; xml: string } | { kind: "tbl"; xml: string })[] = [];
  const tag = /<(\/?)w:(tbl|p)\b[^>]*?(\/?)>/g;
  let depth = 0;
  let tableStart = -1;
  let paragraphStart = -1;
  let match: RegExpExecArray | null;
  while ((match = tag.exec(body))) {
    const [whole, closing, name, selfClosing] = match;
    if (name === "tbl") {
      if (!closing) { if (depth === 0) tableStart = match.index; depth += 1; }
      else { depth -= 1; if (depth === 0 && tableStart >= 0) { blocks.push({ kind: "tbl", xml: body.slice(tableStart, match.index + whole.length) }); tableStart = -1; } }
      continue;
    }
    if (depth > 0) continue;
    if (selfClosing) { blocks.push({ kind: "p", xml: "" }); continue; }
    if (!closing) paragraphStart = match.index;
    else if (paragraphStart >= 0) { blocks.push({ kind: "p", xml: body.slice(paragraphStart, match.index + whole.length) }); paragraphStart = -1; }
  }
  return blocks;
}

/** Строки и ячейки таблицы верхнего уровня; вложенная таблица — текстом ячейки. */
function tableRows(xml: string): string[][] {
  const inner = xml.replace(/^<w:tbl\b[^>]*>/, "").replace(/<\/w:tbl>$/, "");
  const rows: string[][] = [];
  const tag = /<(\/?)w:(tbl|tr|tc)\b[^>]*?>/g;
  let depth = 0;
  let row: string[] | null = null;
  let cellStart = -1;
  let match: RegExpExecArray | null;
  while ((match = tag.exec(inner))) {
    const [whole, closing, name] = match;
    if (name === "tbl") { depth += closing ? -1 : 1; continue; }
    if (depth > 0) continue;
    if (name === "tr") {
      if (!closing) row = [];
      else if (row) { rows.push(row); row = null; }
    } else if (name === "tc") {
      if (!closing) cellStart = match.index + whole.length;
      else if (row && cellStart >= 0) {
        const cellXml = inner.slice(cellStart, match.index);
        const paragraphs = [...cellXml.matchAll(/<w:p\b[^>]*>([\s\S]*?)<\/w:p>/g)].map((item) => paragraphText(item[1]));
        row.push(paragraphs.join("\n").trim());
        cellStart = -1;
      }
    }
  }
  return rows;
}

export function parseDocx(buffer: Buffer, name: string): ParsedFile {
  const zip = zipReader(buffer);
  const document = zip.text("word/document.xml");
  if (!document) throw new FileParseError("В архиве нет word/document.xml: это не документ Word DOCX.");
  const warnings: string[] = [];
  const body = /<w:body\b[^>]*>([\s\S]*)<\/w:body>/.exec(document)?.[1] ?? "";
  const lines: string[] = [];
  const tables: FileTable[] = [];
  let cells = 0;
  for (const block of topLevelBlocks(body)) {
    if (block.kind === "p") { lines.push(paragraphText(block.xml)); continue; }
    const rows = tableRows(block.xml);
    if (!rows.length) continue;
    const columns = Math.max(...rows.map((row) => row.length));
    if (cells + rows.length * columns > FILE_LIMITS.cells) { warnings.push("Таблиц в документе больше предела: разобраны первые."); break; }
    cells += rows.length * columns;
    const grid: CellValue[][] = rows.map((row) => Array.from({ length: columns }, (_, c) => row[c] ?? ""));
    tables.push({ name: `Таблица ${tables.length + 1}`, rows: grid.length, columns, cells: grid });
    lines.push(`[Таблица ${tables.length}: ${grid.length} × ${columns} — читается отдельно]`);
  }
  if (tables.length && /<w:(gridSpan|vMerge)/.test(body)) warnings.push("В таблицах есть объединённые ячейки: значение стоит в первой из них, строки выровнены по самой длинной.");
  if (zip.names().some((entry) => /^word\/(header|footer)\d*\.xml$/.test(entry))) warnings.push("Колонтитулы документа не читаются.");
  if (zip.names().some((entry) => /^word\/(footnotes|endnotes)\.xml$/.test(entry))) warnings.push("Сноски документа не читаются.");
  const text = lines.join("\n").replace(/\n{3,}/g, "\n\n").trim();
  if (!text && !tables.length) throw new FileParseError("В документе нет ни текста, ни таблиц.");
  return { kind: "docx", name, size: buffer.length, tables, text: limitText(chunkText(text), warnings), warnings };
}
