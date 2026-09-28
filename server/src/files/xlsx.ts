/**
 * XLSX (этап 8, 8.6): значения всех листов, как их сохранил Excel.
 *
 * Формулы не переносятся: у ячейки с формулой берётся последнее посчитанное
 * значение, и это называется. Дата в XLSX — число со «стилем даты»: по стилю
 * ячейки она узнаётся и переносится числом с тем же видом. Книга с системой
 * дат 1904 года (старые файлы Mac) переводится в обычную.
 */

import { FILE_LIMITS, type CellValue, type FileTable, type ParsedFile } from "./types.js";
import { attr, FileParseError, xmlText, zipReader } from "./zip.js";

/** Встроенные форматы Excel, которые показывают дату или время. */
const BUILTIN_DATE_FORMATS: Record<number, string> = {
  14: "dd.mm.yyyy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy", 18: "h:mm AM/PM", 19: "h:mm:ss AM/PM",
  20: "h:mm", 21: "h:mm:ss", 22: "dd.mm.yyyy h:mm", 45: "mm:ss", 46: "[h]:mm:ss", 47: "mm:ss.0"
};

/** Код формата — дата или время: есть d, m, y, h, s вне кавычек и скобок, и это не «Основной». */
export function isDateCode(code: string): boolean {
  const bare = code.replace(/"[^"]*"/g, "").replace(/\\./g, "").replace(/\[(?!h\]|m\]|s\])[^\]]*\]/gi, "").toLowerCase();
  if (!bare || bare === "general" || bare === "@") return false;
  return /[dmyhs]/.test(bare) && !/^[#0.,%\s]*$/.test(bare);
}

function columnIndex(ref: string): number {
  const letters = /^[A-Z]+/.exec(ref)?.[0] ?? "A";
  return [...letters].reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0) - 1;
}

function rowIndex(ref: string): number {
  return Number(/\d+$/.exec(ref)?.[0] ?? "1") - 1;
}

/** Текст узла <si>/<is>: простой <t> или части <r><t>, без фонетики <rPh>. */
function stringItem(xml: string): string {
  const withoutPhonetic = xml.replace(/<rPh\b[\s\S]*?<\/rPh>/g, "");
  let text = "";
  for (const match of withoutPhonetic.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>|<t\b[^>]*\/>/g)) text += xmlText(match[1] ?? "");
  return text;
}

export function parseXlsx(buffer: Buffer, name: string): ParsedFile {
  const zip = zipReader(buffer);
  const warnings: string[] = [];
  const workbook = zip.text("xl/workbook.xml");
  if (!workbook) throw new FileParseError("В архиве нет xl/workbook.xml: это не книга Excel XLSX.");
  const date1904 = /<workbookPr\b[^>]*\bdate1904="(1|true)"/.test(workbook);

  const rels = new Map<string, string>();
  for (const match of (zip.text("xl/_rels/workbook.xml.rels") ?? "").matchAll(/<Relationship\b[^>]*>/g)) {
    const id = attr(match[0], "Id");
    const target = attr(match[0], "Target");
    if (id && target) rels.set(id, target.startsWith("/") ? target.slice(1) : `xl/${target.replace(/^\.\//, "")}`);
  }

  const shared: string[] = [];
  for (const match of (zip.text("xl/sharedStrings.xml") ?? "").matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)) shared.push(stringItem(match[1]));

  // Стили: индекс стиля ячейки → код формата, если это дата.
  const styles = zip.text("xl/styles.xml") ?? "";
  const customFormats = new Map<number, string>();
  for (const match of styles.matchAll(/<numFmt\b[^>]*>/g)) {
    const id = Number(attr(match[0], "numFmtId"));
    const code = attr(match[0], "formatCode");
    if (code !== null) customFormats.set(id, code);
  }
  const cellXfs = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles)?.[1] ?? "";
  const styleDate: (string | null)[] = [...cellXfs.matchAll(/<xf\b[^>]*\/?>/g)].map((match) => {
    const id = Number(attr(match[0], "numFmtId") ?? "0");
    if (BUILTIN_DATE_FORMATS[id]) return BUILTIN_DATE_FORMATS[id];
    const code = customFormats.get(id);
    return code && isDateCode(code) ? code : null;
  });

  const tables: FileTable[] = [];
  let totalCells = 0;
  let formulas = 0;
  let truncated = false;
  for (const match of workbook.matchAll(/<sheet\b[^>]*\/?>/g)) {
    const sheetName = attr(match[0], "name") ?? `Лист ${tables.length + 1}`;
    const state = attr(match[0], "state");
    const path = rels.get(attr(match[0], "r:id") ?? "");
    if (!path) continue;
    const xml = zip.text(path);
    if (xml === null) { warnings.push(`Лист «${sheetName}» не найден в архиве.`); continue; }
    const grid = new Map<number, Map<number, CellValue>>();
    const dateFormats: Record<string, string> = {};
    let maxRow = -1;
    let maxColumn = -1;
    const data = /<sheetData\b[^>]*>([\s\S]*)<\/sheetData>/.exec(xml)?.[1] ?? "";
    for (const cell of data.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      if (totalCells >= FILE_LIMITS.cells) { truncated = true; break; }
      const head = ` ${cell[1]}`;
      const ref = attr(head, "r");
      if (!ref) continue;
      const body = cell[2] ?? "";
      const type = attr(head, "t") ?? "n";
      const raw = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(body)?.[1];
      if (/<f\b/.test(body)) formulas += 1;
      let value: CellValue = null;
      if (type === "s") value = raw === undefined ? null : shared[Number(raw)] ?? "";
      else if (type === "inlineStr") value = stringItem(/<is\b[^>]*>([\s\S]*?)<\/is>/.exec(body)?.[1] ?? "");
      else if (type === "str") value = raw === undefined ? "" : xmlText(raw);
      else if (type === "b") value = raw === "1";
      else if (type === "e") value = raw === undefined ? null : xmlText(raw);
      else if (raw !== undefined && raw !== "") {
        const number = Number(raw);
        value = Number.isFinite(number) ? number : xmlText(raw);
      }
      if (value === null || value === "") continue;
      const r = rowIndex(ref);
      const c = columnIndex(ref);
      const style = styleDate[Number(attr(head, "s") ?? "0")];
      if (typeof value === "number" && style) {
        if (date1904) value += 1462;
        dateFormats[`${r},${c}`] = style;
      }
      if (!grid.has(r)) grid.set(r, new Map());
      grid.get(r)!.set(c, value);
      totalCells += 1;
      if (r > maxRow) maxRow = r;
      if (c > maxColumn) maxColumn = c;
    }
    if (maxRow < 0) continue;
    const cells: CellValue[][] = Array.from({ length: maxRow + 1 }, (_, r) =>
      Array.from({ length: maxColumn + 1 }, (_, c) => grid.get(r)?.get(c) ?? null));
    tables.push({
      name: sheetName + (state === "hidden" || state === "veryHidden" ? " (скрытый)" : ""),
      rows: cells.length,
      columns: maxColumn + 1,
      cells,
      ...(Object.keys(dateFormats).length ? { dateFormats } : {})
    });
    if (truncated) break;
  }
  if (truncated) warnings.push(`В книге больше ${FILE_LIMITS.cells.toLocaleString("ru-RU")} заполненных ячеек: разобрано начало.`);
  if (formulas) warnings.push(`Формул в файле: ${formulas}. Переносятся их последние посчитанные значения, сами формулы — нет.`);
  if (date1904) warnings.push("Книга считает даты от 1904 года (старый Mac): даты переведены в обычную систему.");
  if (!tables.length) throw new FileParseError("В книге нет ни одной заполненной ячейки.");
  return { kind: "xlsx", name, size: buffer.length, tables, text: [], warnings };
}
