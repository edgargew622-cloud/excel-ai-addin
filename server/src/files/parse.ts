/**
 * Разбор загруженного файла по виду содержимого (этап 8, 8.6).
 *
 * Вид определяется по первым байтам, а не только по расширению: .xlsx,
 * переименованный в .csv, или исполняемый файл, названный .pdf, не пройдут
 * мимо своего разборщика.
 */

import { parseCsv, parseTxt } from "./csv.js";
import { parseDocx } from "./docx.js";
import { parsePdf } from "./pdf.js";
import type { FileKind, ParsedFile } from "./types.js";
import { parseXlsx } from "./xlsx.js";
import { FileParseError } from "./zip.js";

export function detectKind(buffer: Buffer, name: string): FileKind {
  const extension = (/\.([a-z0-9]+)$/i.exec(name)?.[1] ?? "").toLowerCase();
  const isZip = buffer.length > 4 && buffer.readUInt32LE(0) === 0x04034b50;
  const isPdf = buffer.subarray(0, 1024).toString("latin1").includes("%PDF-");
  if (isPdf) return "pdf";
  if (isZip) {
    const head = buffer.toString("latin1");
    if (head.includes("xl/workbook.xml") || extension === "xlsx") return "xlsx";
    if (head.includes("word/document.xml") || extension === "docx") return "docx";
    throw new FileParseError("Это архив, но не XLSX и не DOCX. Поддерживаются CSV, TXT, XLSX, DOCX и PDF.");
  }
  if (buffer.length >= 8 && buffer.readUInt32BE(0) === 0xd0cf11e0) {
    throw new FileParseError("Это файл старого формата Office (XLS или DOC). Сохраните его в Excel или Word как XLSX или DOCX.");
  }
  if (extension === "csv" || extension === "tsv") return "csv";
  if (extension === "txt" || extension === "md") return "txt";
  if (["xlsx", "docx", "pdf"].includes(extension)) throw new FileParseError(`Файл назван .${extension}, но внутри не ${extension.toUpperCase()}: возможно, он повреждён.`);
  throw new FileParseError(`Формат .${extension || "без расширения"} не поддерживается. Поддерживаются CSV, TXT, XLSX, DOCX и PDF.`);
}

export async function parseFile(buffer: Buffer, name: string): Promise<ParsedFile> {
  const kind = detectKind(buffer, name);
  switch (kind) {
    case "csv": return parseCsv(buffer, name);
    case "txt": return parseTxt(buffer, name);
    case "xlsx": return parseXlsx(buffer, name);
    case "docx": return parseDocx(buffer, name);
    case "pdf": return parsePdf(buffer, name);
  }
}
