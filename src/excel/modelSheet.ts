/**
 * Общее у моделей на отдельном листе (этап 7, 7.5): трёхотчётной, DCF.
 *
 * Модель пишется на новый лист одной записью формул и форматов; допущения —
 * синим текстом. После записи лист читается обратно, и каждая ячейка
 * сверяется с тем, что посчитала панель. Отмена удаляет лист, только если
 * его содержимое не меняли после операции.
 */

import { ToolExecutionError } from "./excelTools";
import { columnLetters } from "./formulaFill";
import { action, getStructuralRevision, push } from "./undo";

export interface ModelCell {
  formula: string | number;
  expected: number | string;
  format?: string;
  input?: boolean;
}

export const AMOUNT = "#,##0;-#,##0";
export const SHARE = "0.0%";

/** Строка из подписи и ячеек, дополненная пустыми до ширины. */
export function line(columns: number, label: string, cells: ModelCell[] = []): ModelCell[] {
  const row: ModelCell[] = [{ formula: label, expected: label }, ...cells];
  while (row.length < columns) row.push({ formula: "", expected: "" });
  return row;
}

export const text = (value: string): ModelCell => ({ formula: value, expected: value });

/** Расхождения прочитанного с расчётом: числа с допуском, текст как текст. */
export function modelMismatches(rows: readonly (readonly ModelCell[])[], values: readonly (readonly unknown[])[]): string[] {
  const problems: string[] = [];
  rows.forEach((row, r) => row.forEach((cell, c) => {
    const actual = values[r]?.[c];
    const name = `${columnLetters(1 + c)}${r + 1}`;
    if (typeof cell.expected === "number") {
      if (typeof actual !== "number" || Math.abs(actual - cell.expected) > 1e-9 * Math.max(1, Math.abs(cell.expected))) problems.push(`${name}: ${String(actual)} вместо ${cell.expected}`);
    } else if (String(actual ?? "") !== cell.expected) {
      problems.push(`${name}: «${String(actual)}» вместо «${cell.expected}»`);
    }
  }));
  return problems;
}

export const sheetAddress = (rows: readonly (readonly ModelCell[])[], columns: number) => `A1:${columnLetters(columns)}${rows.length}`;

/**
 * Новый лист с моделью: запись, чтение обратно, отмена. Возвращает
 * прочитанные значения — сверку и контроль делает вызывающий.
 */
export async function writeModelSheet(
  ctx: Excel.RequestContext,
  options: { sheetName: string; rows: readonly (readonly ModelCell[])[]; columns: number; undoAvailable: boolean; label: string }
): Promise<{ sheetName: string; values: unknown[][]; undoRecorded: boolean }> {
  const address = sheetAddress(options.rows, options.columns);
  const existing = ctx.workbook.worksheets.getItemOrNullObject(options.sheetName);
  existing.load("isNullObject");
  await ctx.sync();
  if (!existing.isNullObject) throw new ToolExecutionError(`Лист «${options.sheetName}» появился после предпросмотра. Модель не строилась.`, "failed_before_write");

  let sheet: Excel.Worksheet;
  try {
    sheet = ctx.workbook.worksheets.add(options.sheetName);
    sheet.load(["id", "name"]);
    const range = sheet.getRange(address);
    range.formulas = options.rows.map((row) => row.map((cell) => cell.formula)) as any[][];
    range.numberFormat = options.rows.map((row) => row.map((cell) => cell.format ?? "General")) as any[][];
    options.rows.forEach((row, r) => row.forEach((cell, c) => {
      if (cell.input) range.getCell(r, c).format.font.color = "#0000FF";
    }));
    range.getRow(0).format.font.bold = true;
    sheet.getRange("A:A").format.columnWidth = 260;
    await ctx.sync();
  } catch (error: any) {
    throw new ToolExecutionError(`Не удалось определить, построилась ли модель на листе «${options.sheetName}»: ${error?.message ?? error}. Посмотрите на книгу.`, "unknown");
  }
  const range = sheet.getRange(address);
  range.load(["values", "formulas"]);
  await ctx.sync();
  const sheetId = sheet.id;

  let undoRecorded = false;
  if (options.undoAvailable) {
    const written = JSON.stringify(range.formulas);
    undoRecorded = push(action(`${options.label} на листе «${sheet.name}»`, async () => {
      const revision = getStructuralRevision();
      await Excel.run(async (undoCtx) => {
        const target = undoCtx.workbook.worksheets.getItemOrNullObject(sheetId);
        target.load(["isNullObject", "name"]);
        await undoCtx.sync();
        if (target.isNullObject) throw new Error("Листа модели уже нет. Отменять нечего.");
        const used = target.getUsedRangeOrNullObject(true);
        used.load(["isNullObject", "address"]);
        const block = target.getRange(address);
        block.load("formulas");
        await undoCtx.sync();
        const outside = !used.isNullObject && String(used.address).replace(/^.*!/, "") !== address;
        if (JSON.stringify(block.formulas) !== written || outside) {
          throw new Error(`Лист «${target.name}» изменили после операции агента. Отмена остановлена: удаление унесло бы правки. Удалите лист вручную, если он не нужен.`);
        }
        if (getStructuralRevision() !== revision) throw new Error("Структура книги изменилась во время отмены. Отмена остановлена.");
        target.delete();
        await undoCtx.sync();
      });
    }));
  }
  return { sheetName: sheet.name, values: range.values as unknown[][], undoRecorded };
}
