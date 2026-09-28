/**
 * CSV и TXT (этап 8, 8.6).
 *
 * Кодировка: метка BOM, иначе UTF-8, если байты ею читаются, иначе
 * Windows-1251 — так сохраняет русский Excel. Разделитель — тот из «; , Tab |»,
 * что встречается одинаково часто в первых строках вне кавычек. Кавычки —
 * по RFC 4180: поле в кавычках может содержать разделитель, перевод строки
 * и удвоенную кавычку.
 *
 * Значения остаются текстом: «1 200,50», «01.02.2026», «007» разобрать
 * однозначно нельзя, это делает перенос в книгу по правилам 7.2.3.
 */

import { chunkText, FILE_LIMITS, limitText, type FileTable, type ParsedFile } from "./types.js";
import { FileParseError } from "./zip.js";

export function decodeText(buffer: Buffer): { text: string; encoding: string } {
  if (buffer[0] === 0xef && buffer[1] === 0xbb && buffer[2] === 0xbf) return { text: buffer.subarray(3).toString("utf8"), encoding: "UTF-8" };
  if (buffer[0] === 0xff && buffer[1] === 0xfe) return { text: new TextDecoder("utf-16le").decode(buffer.subarray(2)), encoding: "UTF-16" };
  if (buffer[0] === 0xfe && buffer[1] === 0xff) return { text: new TextDecoder("utf-16be").decode(buffer.subarray(2)), encoding: "UTF-16" };
  try {
    return { text: new TextDecoder("utf-8", { fatal: true }).decode(buffer), encoding: "UTF-8" };
  } catch {
    return { text: new TextDecoder("windows-1251").decode(buffer), encoding: "Windows-1251" };
  }
}

const DELIMITERS = [";", ",", "\t", "|"];

/** Сколько раз разделитель стоит вне кавычек в строке. */
function countOutsideQuotes(line: string, delimiter: string): number {
  let count = 0;
  let quoted = false;
  for (const char of line) {
    if (char === '"') quoted = !quoted;
    else if (char === delimiter && !quoted) count += 1;
  }
  return count;
}

export function detectDelimiter(text: string): string {
  const lines = text.split(/\r?\n/).filter((line) => line.trim()).slice(0, 20);
  let best = ",";
  let bestScore = -1;
  for (const delimiter of DELIMITERS) {
    const counts = lines.map((line) => countOutsideQuotes(line, delimiter));
    const first = counts[0] ?? 0;
    if (!first) continue;
    // Одинаковое число разделителей в строках — признак настоящего разделителя.
    const consistent = counts.filter((count) => count === first).length;
    const score = consistent * 100 + first;
    if (score > bestScore) { bestScore = score; best = delimiter; }
  }
  return best;
}

export function parseCsvRows(text: string, delimiter: string, maxCells: number): { rows: string[][]; truncated: boolean } {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let quoted = false;
  let cells = 0;
  const pushField = () => { row.push(field); field = ""; cells += 1; };
  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    if (quoted) {
      if (char === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; } else quoted = false;
      } else field += char;
      continue;
    }
    if (char === '"' && field === "") { quoted = true; continue; }
    if (char === delimiter) { pushField(); continue; }
    if (char === "\n" || char === "\r") {
      if (char === "\r" && text[i + 1] === "\n") i++;
      pushField();
      rows.push(row);
      row = [];
      if (cells >= maxCells) return { rows, truncated: i < text.length - 1 };
      continue;
    }
    field += char;
  }
  if (field !== "" || row.length) { pushField(); rows.push(row); }
  return { rows, truncated: false };
}

export function parseCsv(buffer: Buffer, name: string): ParsedFile {
  const warnings: string[] = [];
  const { text, encoding } = decodeText(buffer);
  if (/\u0000/.test(text.slice(0, 4096))) throw new FileParseError("Файл двоичный, а не текстовый CSV: разобрать нельзя.");
  const delimiter = detectDelimiter(text);
  const { rows, truncated } = parseCsvRows(text, delimiter, FILE_LIMITS.cells);
  while (rows.length && rows[rows.length - 1].every((cell) => cell === "")) rows.pop();
  if (!rows.length) throw new FileParseError("В CSV нет ни одной строки данных.");
  if (truncated) warnings.push(`В файле больше ${FILE_LIMITS.cells.toLocaleString("ru-RU")} ячеек: разобрано начало, ${rows.length} строк.`);
  const columns = Math.max(...rows.map((row) => row.length));
  const uneven = rows.filter((row) => row.length !== columns).length;
  if (uneven) warnings.push(`В ${uneven} строках меньше полей, чем в самой длинной (${columns}): недостающие ячейки пустые.`);
  const cells = rows.map((row) => Array.from({ length: columns }, (_, c) => row[c] ?? ""));
  const delimiterName = delimiter === "\t" ? "табуляция" : `«${delimiter}»`;
  warnings.push(`Кодировка ${encoding}, разделитель ${delimiterName}.`);
  const table: FileTable = { name, rows: cells.length, columns, cells };
  return { kind: "csv", name, size: buffer.length, tables: [table], text: [], warnings };
}

export function parseTxt(buffer: Buffer, name: string): ParsedFile {
  const warnings: string[] = [];
  const { text, encoding } = decodeText(buffer);
  if (/\u0000/.test(text.slice(0, 4096))) throw new FileParseError("Файл двоичный, а не текстовый: разобрать нельзя.");
  warnings.push(`Кодировка ${encoding}.`);
  return { kind: "txt", name, size: buffer.length, tables: [], text: limitText(chunkText(text.replace(/\r\n?/g, "\n")), warnings), warnings };
}
