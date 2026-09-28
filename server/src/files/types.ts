/**
 * Разобранный файл пользователя (этап 8, 8.6): таблицы и текст.
 *
 * Всё, что здесь лежит, — недоверенные данные: модель получает их частями
 * как содержимое, а не как указания.
 */

export type FileKind = "csv" | "xlsx" | "docx" | "pdf" | "txt";

export type CellValue = string | number | boolean | null;

export interface FileTable {
  /** Имя листа XLSX, «Таблица 2» для DOCX, имя файла для CSV. */
  name: string;
  rows: number;
  columns: number;
  cells: CellValue[][];
  /** Числовой формат даты по ячейке «r,c» — только для дат из XLSX (там дата — число). */
  dateFormats?: Record<string, string>;
}

export interface ParsedFile {
  kind: FileKind;
  name: string;
  size: number;
  tables: FileTable[];
  /** Текст частями: страницы PDF, куски DOCX/TXT. */
  text: string[];
  /** Что было при разборе упрощено или отброшено — пересказать пользователю. */
  warnings: string[];
}

export const FILE_LIMITS = {
  /** Размер загружаемого файла. */
  bytes: 20 * 1024 * 1024,
  /** Ячеек во всех таблицах файла. */
  cells: 300_000,
  /** Страниц PDF. */
  pages: 300,
  /** Знаков текста во всём файле. */
  textChars: 2_000_000,
  /** Размер одного куска текста DOCX/TXT. */
  chunkChars: 4_000
};

/** Текст → куски по абзацам, не длиннее chunkChars. */
export function chunkText(text: string, size = FILE_LIMITS.chunkChars): string[] {
  const chunks: string[] = [];
  let current = "";
  for (const paragraph of text.split(/\n/)) {
    if (current && current.length + paragraph.length + 1 > size) {
      chunks.push(current);
      current = "";
    }
    if (paragraph.length > size) {
      for (let i = 0; i < paragraph.length; i += size) chunks.push(paragraph.slice(i, i + size));
      continue;
    }
    current = current ? `${current}\n${paragraph}` : paragraph;
  }
  if (current.trim()) chunks.push(current);
  return chunks;
}

/** Обрезка по общему пределу текста — с предупреждением, а не молча. */
export function limitText(chunks: string[], warnings: string[]): string[] {
  let total = 0;
  const kept: string[] = [];
  for (const chunk of chunks) {
    if (total + chunk.length > FILE_LIMITS.textChars) {
      warnings.push(`Текст длиннее ${FILE_LIMITS.textChars.toLocaleString("ru-RU")} знаков: разобрано только начало (${kept.length} частей из ${chunks.length}).`);
      break;
    }
    total += chunk.length;
    kept.push(chunk);
  }
  return kept;
}
