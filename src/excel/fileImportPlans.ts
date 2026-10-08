/**
 * Перенос таблицы из прикреплённого файла в книгу (этап 8, 8.6).
 *
 * Данные берутся прямо из разобранного файла, а не из текста модели: модель
 * лишь называет файл, таблицу и место, и не может исказить значения при
 * перепечатке. Место обязано быть пустым: перенос ничего не затирает.
 *
 * Значения:
 * - из XLSX — как в файле: числа числами, текст текстом, даты числом с тем
 *   же видом, что в файле;
 * - из CSV и Word — текст, кроме однозначных чисел: целых без ведущего нуля и
 *   дробных с точкой или запятой, если после неё не ровно три цифры («1,500»
 *   может быть и тысячей пятьсот). «1 200,50», «007», «01.02.2026» остаются
 *   текстом — их переводит convert_values по правилам книги (7.2.3).
 * Текст всегда пишется текстом: «=HYPERLINK(…)» из файла не станет формулой.
 */

import { intersects, parseA1Rect, type A1Rect } from "./a1";
import {
  assertPlanWorkbook,
  checkAddress,
  deepFreeze,
  preflightToolArgs,
  readTableRanges,
  ToolError,
  ToolExecutionError,
  valuesForLiteralWrite,
  fitNewTable
} from "./excelTools";
import { columnLetters } from "./formulaFill";
import { checkSheetName, freeSheetName } from "./sheetRules";
import { action, getStructuralRevision, isCustomUndoAvailable, push } from "./undo";
import { captureTarget, type WorkbookTarget } from "./workbookContext";
import { fetchFullTable, type FileKind } from "../taskpane/api/files";

export const MAX_IMPORT_CELLS = 100_000;
const WRITE_CELLS_PER_SYNC = 5_000;
const TOLERANCE = 1e-9;

type Cell = string | number | boolean;

/** Однозначное число из текста файла — или null. */
export function plainNumber(text: string): number | null {
  const trimmed = text.trim();
  if (/^-?(0|[1-9]\d{0,14})$/.test(trimmed)) return Number(trimmed);
  const decimal = /^(-?(?:0|[1-9]\d{0,14}))[.,](\d+)$/.exec(trimmed);
  if (decimal && decimal[2].length !== 3) return Number(`${decimal[1]}.${decimal[2]}`);
  return null;
}

export interface ImportConversion {
  values: Cell[][];
  /** Формат даты по ячейке «r,c» блока. */
  dateFormats: Record<string, string>;
  numbersFromText: number;
  /** Похоже на число или дату, но оставлено текстом — для convert_values. */
  keptAsText: { count: number; examples: string[] };
}

export function convertTable(
  kind: FileKind,
  cells: readonly (readonly (string | number | boolean | null)[])[],
  dateFormats: Record<string, string> = {},
  firstRow = 0
): ImportConversion {
  const out: ImportConversion = { values: [], dateFormats: {}, numbersFromText: 0, keptAsText: { count: 0, examples: [] } };
  const convertText = kind !== "xlsx";
  // Реквизиты — не числа: ИНН, КПП, счёт, БИК, номер документа остаются текстом,
  // как в файле. Иначе в одном столбце ИНН часть становилась числом, а с ведущим
  // нулём оставалась текстом («Книга11», 08.10.2026).
  const header = firstRow === 0 ? cells[0] ?? [] : [];
  const identifier = header.map((name) => typeof name === "string" && /инн|кпп|огрн|бик|сч[её]т|номер|№|код|телефон|паспорт|снилс|артикул/i.test(name));
  cells.forEach((row, r) => {
    out.values.push(row.map((value, c) => {
      if (identifier[c] && r > 0 && typeof value === "string") return value;
      if (value === null || value === undefined) return "";
      if (typeof value === "number" || typeof value === "boolean") {
        const format = dateFormats[`${r + firstRow},${c}`];
        if (format && typeof value === "number") out.dateFormats[`${r},${c}`] = format;
        return value;
      }
      if (!convertText || value === "") return value;
      const number = plainNumber(value);
      if (number !== null) { out.numbersFromText += 1; return number; }
      if (/^[-+]?[\d\s .,]+$/.test(value.trim()) || /^\d{1,4}[./-]\d{1,2}[./-]\d{1,4}$/.test(value.trim())) {
        out.keptAsText.count += 1;
        if (out.keptAsText.examples.length < 5) out.keptAsText.examples.push(value);
      }
      return value;
    }));
  });
  return out;
}

export interface ImportFilePlan {
  readonly kind: "import_file_table";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly fileId: string;
  readonly fileName: string;
  readonly fileKind: FileKind;
  readonly tableName: string;
  readonly sourceRows: string;
  readonly destSheet: string;
  readonly destSheetId: string;
  readonly newSheet?: true;
  readonly destArea: string;
  readonly rows: number;
  readonly columns: number;
  readonly values: readonly (readonly Cell[])[];
  readonly dateFormats: Readonly<Record<string, string>>;
  readonly numbersFromText: number;
  readonly keptAsText: { readonly count: number; readonly examples: readonly string[] };
  readonly preview: readonly string[];
  readonly undoAvailable: boolean;
  readonly undoNote?: string;
  readonly createdAt: string;
}

function areaOf(cell: string, rows: number, columns: number): { rect: A1Rect; address: string } {
  const start = parseA1Rect(cell)!;
  const rect: A1Rect = { kind: "cells", rowStart: start.rowStart, columnStart: start.columnStart, rowEnd: start.rowStart + rows - 1, columnEnd: start.columnStart + columns - 1 };
  return { rect, address: `${columnLetters(rect.columnStart)}${rect.rowStart}:${columnLetters(rect.columnEnd)}${rect.rowEnd}` };
}

/** Почему место не годится, или null. */
async function placeProblem(ctx: Excel.RequestContext, sheet: Excel.Worksheet, area: { rect: A1Rect; address: string }): Promise<string | null> {
  sheet.load("name");
  try { sheet.protection?.load("protected"); } catch { /* нет сведений */ }
  const place = sheet.getRange(area.address);
  place.load(["formulas", "values"]);
  await ctx.sync();
  if ((sheet as any).protection?.protected) return `Лист ${sheet.name} защищён: записать на него нельзя.`;
  const busy = (place.formulas as unknown[][]).flat().filter((value) => value !== "" && value !== null).length;
  if (busy) return `Место ${sheet.name}!${area.address} занято: там ${busy} непустых ячеек — перенос их затёр бы.`;
  const tables = await readTableRanges(ctx, sheet);
  const hit = tables.find((table) => { const rect = parseA1Rect(table.address.replace(/^.*!/, "")); return rect && intersects(rect, area.rect); });
  if (hit) return `Место ${sheet.name}!${area.address} задевает таблицу ${hit.name}.`;
  return null;
}

export async function prepareImportFilePlan(args: unknown): Promise<ImportFilePlan> {
  preflightToolArgs("import_file_table", args);
  const a = args as { fileId: string; table?: number; sheet?: string; destAddress?: string; newSheet?: string; firstRow?: number; lastRow?: number };
  if (a.newSheet?.trim() && a.destAddress?.trim()) throw new ToolError("newSheet не сочетается с destAddress: на новом листе таблица встаёт в A1.");
  let dest = "A1";
  if (a.destAddress?.trim()) {
    const rect = parseA1Rect(checkAddress(a.destAddress));
    if (!rect || rect.kind !== "cells") throw new ToolError(`destAddress — левая верхняя ячейка, например A1; получено «${a.destAddress}».`);
    dest = `${columnLetters(rect.columnStart)}${rect.rowStart}`;
  }
  const table = await fetchFullTable(a.fileId, a.table ?? 0);
  const first = Math.max(1, Math.floor(a.firstRow ?? 1));
  const last = Math.min(table.rows, Math.floor(a.lastRow ?? table.rows));
  if (first > last) throw new ToolError(`Строки ${first}–${last} вне таблицы: в ней ${table.rows} строк.`);
  const slice = table.cells.slice(first - 1, last);
  // Пустые столбцы справа не переносятся.
  let columns = table.columns;
  while (columns > 1 && slice.every((row) => row[columns - 1] === null || row[columns - 1] === "")) columns -= 1;
  const rows = slice.length;
  if (rows * columns > MAX_IMPORT_CELLS) {
    throw new ToolError(`В таблице ${rows * columns} ячеек, за один перенос — до ${MAX_IMPORT_CELLS}. Перенесите частями: firstRow и lastRow.`);
  }
  const conversion = convertTable(table.kind, slice.map((row) => row.slice(0, columns)), table.dateFormats ?? {}, first - 1);
  const target = await captureTarget(a.sheet);

  const prepared = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    sheet.load(["id", "name"]);
    const all = ctx.workbook.worksheets;
    all.load("items/name");
    await ctx.sync();
    let destSheet = sheet.name;
    let destSheetId = sheet.id;
    let newSheet = false;
    if (a.newSheet?.trim()) {
      try {
        destSheet = checkSheetName(a.newSheet, all.items.map((item) => item.name));
      } catch (error: any) {
        const message = String(error?.message ?? error);
        throw new ToolError(/уже есть/.test(message) ? `${message} Свободно, например, «${freeSheetName(String(a.newSheet), all.items.map((item) => item.name))}».` : message);
      }
      destSheetId = "";
      newSheet = true;
    }
    const area = areaOf(dest, rows, columns);
    if (!newSheet) {
      const problem = await placeProblem(ctx, sheet, area);
      if (problem) {
        const used = sheet.getUsedRangeOrNullObject(true);
        used.load(["isNullObject", "columnIndex", "columnCount", "rowIndex"]);
        await ctx.sync();
        const free = used.isNullObject ? "A1" : `${columnLetters(used.columnIndex + used.columnCount + 2)}${used.rowIndex + 1}`;
        throw new ToolError(`${problem} Операция не выполнялась. Можно на новый лист (newSheet) или правее данных — destAddress: "${free}".`);
      }
    }
    const shown = (value: Cell) => (typeof value === "string" ? value : String(value));
    const undo = isCustomUndoAvailable();
    return {
      kind: "import_file_table" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      target: { ...target, sheetName: sheet.name },
      fileId: a.fileId,
      fileName: table.file,
      fileKind: table.kind,
      tableName: table.name,
      sourceRows: `${first}–${last}`,
      destSheet,
      destSheetId,
      ...(newSheet ? { newSheet: true as const } : {}),
      destArea: area.address,
      rows,
      columns,
      values: conversion.values,
      dateFormats: conversion.dateFormats,
      numbersFromText: conversion.numbersFromText,
      keptAsText: conversion.keptAsText,
      preview: conversion.values.slice(0, 5).map((row) => row.slice(0, 8).map(shown).join(" | ") + (row.length > 8 ? " | …" : "")),
      undoAvailable: undo,
      ...(undo ? {} : { undoNote: "Отмена недоступна: монитор изменений Excel не активен." }),
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

function sameValue(actual: unknown, expected: Cell): boolean {
  if (typeof expected === "number") return typeof actual === "number" && Math.abs(actual - expected) <= TOLERANCE * Math.max(1, Math.abs(expected));
  if (typeof expected === "boolean") return actual === expected;
  return String(actual ?? "") === expected;
}

export async function executeImportFilePlan(plan: ImportFilePlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    let sheet: Excel.Worksheet;
    let createdSheetId: string | null = null;
    const area = areaOf(plan.destArea.split(":")[0], plan.rows, plan.columns);
    if (plan.newSheet) {
      const all = ctx.workbook.worksheets;
      all.load("items/name");
      await ctx.sync();
      if (all.items.some((item) => item.name.trim().toLowerCase() === plan.destSheet.toLowerCase())) {
        throw new ToolExecutionError(`Лист «${plan.destSheet}» появился после предпросмотра. Перенос не выполнялся — выберите другое имя.`, "failed_before_write");
      }
      try {
        sheet = ctx.workbook.worksheets.add(plan.destSheet);
        sheet.load(["id", "name"]);
        await ctx.sync();
        createdSheetId = sheet.id;
      } catch (error: any) {
        throw new ToolExecutionError(`Excel отказал в создании листа «${plan.destSheet}»: ${error?.message ?? error}.`, "unknown");
      }
    } else {
      sheet = ctx.workbook.worksheets.getItem(plan.destSheetId);
      sheet.load(["id", "name"]);
      const problem = await placeProblem(ctx, sheet, area);
      if (problem) throw new ToolExecutionError(`${problem} Это появилось после предпросмотра. Перенос не выполнялся.`, "failed_before_write");
    }
    const where = `${plan.destSheet}!${plan.destArea}`;
    const rowsPerSync = Math.max(1, Math.floor(WRITE_CELLS_PER_SYNC / plan.columns));
    try {
      for (let start = 0; start < plan.rows; start += rowsPerSync) {
        const rows = plan.values.slice(start, start + rowsPerSync);
        const chunk = sheet.getRangeByIndexes(area.rect.rowStart - 1 + start, area.rect.columnStart - 1, rows.length, plan.columns);
        chunk.values = valuesForLiteralWrite(rows) as any[][];
        const formats = rows.map((row, r) => row.map((_, c) => plan.dateFormats[`${start + r},${c}`]));
        if (formats.some((row) => row.some(Boolean))) {
          formats.forEach((row, r) => row.forEach((format, c) => { if (format) chunk.getCell(r, c).numberFormat = [[format]] as any[][]; }));
        }
        await ctx.sync();
      }
    } catch (error: any) {
      throw new ToolExecutionError(`Перенос на ${where} прервался: ${error?.message ?? error}. Часть данных могла записаться — перечитайте область.`, "unknown");
    }

    // Перенесённая таблица — сразу по содержимому (даты, суммы, длинные названия).
    await fitNewTable(ctx, sheet, area.rect.columnStart - 1, plan.columns);
    const written = sheet.getRange(plan.destArea);
    written.load(["values", "formulas"]);
    await ctx.sync();
    const values = written.values as unknown[][];
    const formulas = written.formulas as unknown[][];
    const problems: string[] = [];
    plan.values.forEach((row, r) => row.forEach((expected, c) => {
      const cell = `${columnLetters(area.rect.columnStart + c)}${area.rect.rowStart + r}`;
      // Текст «=…» из файла, ставший формулой, дал бы здесь её результат, а не сам текст.
      if (!sameValue(values[r]?.[c], expected)) problems.push(`${cell}: ${JSON.stringify(values[r]?.[c])} вместо ${JSON.stringify(expected)}`);
    }));

    let undoRecorded = false;
    if (plan.undoAvailable) {
      const signature = JSON.stringify(formulas);
      const sheetId = createdSheetId ?? plan.destSheetId;
      undoRecorded = push(action(`перенос «${plan.fileName}» на ${where}`, async () => {
        const revision = getStructuralRevision();
        await Excel.run(async (undoCtx) => {
          const undoSheet = undoCtx.workbook.worksheets.getItem(sheetId);
          const range = undoSheet.getRange(plan.destArea);
          range.load("formulas");
          await undoCtx.sync();
          if (JSON.stringify(range.formulas) !== signature) throw new Error(`Данные на ${where} изменили после переноса. Отмена остановлена, чтобы не затереть правку.`);
          if (getStructuralRevision() !== revision) throw new Error("Структура книги изменилась во время отмены. Отмена остановлена.");
          if (createdSheetId) {
            const used = undoSheet.getUsedRangeOrNullObject(true);
            used.load(["isNullObject", "address"]);
            await undoCtx.sync();
            // Лист удаляется, только если на нём нет ничего, кроме перенесённого.
            if (used.isNullObject || String(used.address).replace(/^.*!/, "") === plan.destArea) {
              undoSheet.delete();
              await undoCtx.sync();
              return;
            }
          }
          range.clear("All" as any);
          await undoCtx.sync();
        });
      }));
    }

    if (problems.length) {
      throw new ToolExecutionError(
        `Таблица перенесена на ${where}, но ${problems.length} ячеек расходятся с файлом: ${problems.slice(0, 6).join("; ")}. ` +
        (undoRecorded ? "Перенос можно отменить кнопкой «Отменить»." : "Проверьте область."),
        "applied"
      );
    }
    return {
      ok: true,
      executionState: "verified",
      file: plan.fileName,
      table: plan.tableName,
      sourceRows: plan.sourceRows,
      address: where,
      ...(plan.newSheet ? { newSheet: plan.destSheet } : {}),
      cells: plan.rows * plan.columns,
      numbersFromText: plan.numbersFromText,
      ...(plan.keptAsText.count ? {
        keptAsText: plan.keptAsText,
        keptAsTextNote: "Эти значения похожи на числа или даты, но записаны текстом: разделители и порядок даты неоднозначны. Предложи convert_values — он переведёт однозначные по правилам книги и спросит про остальные."
      } : {}),
      note: "Каждая ячейка сверена с файлом; текст из файла записан текстом, формулой он не стал.",
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: plan.undoNote ?? "Автоматическая отмена этой операции недоступна." })
    };
  });
}
