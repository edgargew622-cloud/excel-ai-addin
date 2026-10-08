/**
 * «Широкая таблица → длинная» (срез после 1.0.36, 08.10.2026): замена шага
 * «Отменить свёртывание столбцов» (Unpivot) из Power Query, которого
 * надстройкам Excel не дают. Месяцы в столбцах превращаются в плоскую базу:
 * ключевые столбцы + «Показатель» (заголовок столбца) + «Значение».
 *
 * Результат пишется значениями на новый лист — источник не меняется. Сверка
 * после записи: число строк и сумма чисел совпадают с исходной таблицей.
 */

import { assertPlanWorkbook, checkAddress, deepFreeze, fitNewTable, preflightToolArgs, ToolError, ToolExecutionError, valuesForLiteralWrite } from "./excelTools";
import { checkSheetName, freeSheetName } from "./sheetRules";
import { action, getStructuralRevision, isCustomUndoAvailable, push } from "./undo";
import { captureTarget, type WorkbookTarget } from "./workbookContext";
import { columnLetters } from "./formulaFill";

export const MAX_UNPIVOT_CELLS = 100_000;
const WRITE_CELLS_PER_SYNC = 5_000;
const isBlank = (value: unknown) => value === "" || value === null || value === undefined;

export interface UnpivotResult {
  header: unknown[];
  rows: unknown[][];
  skipped: number;
  sum: number;
}

/** Разворот без Excel: первые idCount столбцов — ключи, остальные — показатели. */
export function unpivotRows(values: readonly (readonly unknown[])[], idCount: number, names: { attribute: string; value: string }, skipBlanks: boolean): UnpivotResult {
  const [head, ...body] = values;
  const header = [...head.slice(0, idCount), names.attribute, names.value];
  const rows: unknown[][] = [];
  let skipped = 0;
  let sum = 0;
  for (const row of body) {
    if (row.every(isBlank)) continue;
    for (let column = idCount; column < head.length; column++) {
      const value = row[column];
      if (skipBlanks && isBlank(value)) { skipped++; continue; }
      if (typeof value === "number" && Number.isFinite(value)) sum += value;
      rows.push([...row.slice(0, idCount), head[column], value ?? ""]);
    }
  }
  return { header, rows, skipped, sum };
}

export interface UnpivotPlan {
  readonly kind: "unpivot_range";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly sourceSheet: string;
  readonly sourceAddress: string;
  readonly destSheet: string;
  readonly idColumns: readonly string[];
  readonly valueColumns: number;
  readonly header: readonly unknown[];
  readonly rows: readonly (readonly unknown[])[];
  /** Форматы чисел столбцов результата: ключи — как в источнике, показатель — как шапка, значение — как первое значение. */
  readonly formats: readonly string[];
  readonly skipped: number;
  readonly sum: number;
  readonly preview: readonly string[];
  readonly undoAvailable: boolean;
  readonly createdAt: string;
}

export async function prepareUnpivotPlan(args: unknown): Promise<UnpivotPlan> {
  preflightToolArgs("unpivot_range", args);
  const a = args as { sheet?: string; address: string; idColumns?: number; attributeName?: string; valueName?: string; newSheet?: string; keepBlanks?: boolean };
  const address = checkAddress(a.address);
  const target = await captureTarget(a.sheet);
  const prepared = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    sheet.load("name");
    const range = sheet.getRange(address);
    range.load(["address", "values", "numberFormat", "rowCount", "columnCount", "columnIndex"]);
    const all = ctx.workbook.worksheets;
    all.load("items/name");
    await ctx.sync();
    const values = range.values as unknown[][];
    const idCount = Math.round(a.idColumns ?? 1);
    if (range.rowCount < 2) throw new ToolError("Нужна шапка и хотя бы одна строка данных.");
    if (idCount < 1 || idCount >= range.columnCount) throw new ToolError(`idColumns — сколько первых столбцов оставить ключами: от 1 до ${range.columnCount - 1}.`);
    const head = values[0];
    const blankHead = head.map((value, index) => (isBlank(value) ? columnLetters(range.columnIndex + index + 1) : "")).filter(Boolean);
    if (blankHead.length) throw new ToolError(`В шапке пустые ячейки (столбцы ${blankHead.join(", ")}): по шапке называются показатели. Выделите область с шапкой.`);
    const names = { attribute: (a.attributeName ?? "Показатель").trim() || "Показатель", value: (a.valueName ?? "Значение").trim() || "Значение" };
    const result = unpivotRows(values, idCount, names, a.keepBlanks !== true);
    if (!result.rows.length) throw new ToolError("После разворота не осталось ни одной строки: все значения пусты.");
    const cells = (result.rows.length + 1) * result.header.length;
    if (cells > MAX_UNPIVOT_CELLS) throw new ToolError(`Получится ${cells} ячеек, за раз — до ${MAX_UNPIVOT_CELLS}. Разверните часть строк.`);
    const wanted = a.newSheet?.trim() || `${sheet.name} — длинная`;
    let destSheet: string;
    try {
      destSheet = checkSheetName(wanted, all.items.map((item) => item.name));
    } catch (error: any) {
      const message = String(error?.message ?? error);
      throw new ToolError(/уже есть/.test(message) ? `${message} Свободно, например, «${freeSheetName(wanted, all.items.map((item) => item.name))}».` : message);
    }
    const formats = range.numberFormat as unknown[][];
    const firstRow = formats[1] ?? [];
    const resultFormats = [...firstRow.slice(0, idCount), formats[0]?.[idCount], firstRow[idCount]].map((item) => String(item ?? "General"));
    const shown = (row: readonly unknown[]) => row.map((item) => String(item ?? "")).join(" | ");
    const undo = isCustomUndoAvailable();
    return {
      kind: "unpivot_range" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      target: { ...target, sheetName: sheet.name },
      sourceSheet: sheet.name,
      sourceAddress: String(range.address).replace(/^.*!/, ""),
      destSheet,
      idColumns: head.slice(0, idCount).map(String),
      valueColumns: head.length - idCount,
      header: result.header,
      rows: result.rows,
      formats: resultFormats,
      skipped: result.skipped,
      sum: result.sum,
      preview: [
        `Ключи: ${head.slice(0, idCount).map((item) => `«${item}»`).join(", ")}; ${head.length - idCount} столбцов-показателей (${head.slice(idCount, idCount + 4).map((item) => `«${item}»`).join(", ")}${head.length - idCount > 4 ? ", …" : ""}) станут строками «${names.attribute}» / «${names.value}».`,
        `Получится ${result.rows.length} строк на новом листе «${destSheet}»${result.skipped ? `; пустых значений пропущено: ${result.skipped}` : ""}. Источник не меняется.`,
        shown(result.header),
        ...result.rows.slice(0, 4).map(shown)
      ],
      undoAvailable: undo,
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executeUnpivotPlan(plan: UnpivotPlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const all = ctx.workbook.worksheets;
    all.load("items/name");
    await ctx.sync();
    if (all.items.some((item) => item.name.trim().toLowerCase() === plan.destSheet.toLowerCase())) {
      throw new ToolExecutionError(`Лист «${plan.destSheet}» появился после предпросмотра. Ничего не записано — выберите другое имя.`, "failed_before_write");
    }
    let sheet: Excel.Worksheet;
    try {
      sheet = ctx.workbook.worksheets.add(plan.destSheet);
      sheet.load(["id", "name"]);
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(`Excel отказал в создании листа «${plan.destSheet}»: ${error?.message ?? error}.`, "unknown");
    }
    const width = plan.header.length;
    const total = plan.rows.length + 1;
    const lines = [plan.header, ...plan.rows];
    const perSync = Math.max(1, Math.floor(WRITE_CELLS_PER_SYNC / width));
    try {
      for (let start = 0; start < total; start += perSync) {
        const part = lines.slice(start, start + perSync);
        const chunk = sheet.getRangeByIndexes(start, 0, part.length, width);
        chunk.values = valuesForLiteralWrite(part as unknown[][]) as any[][];
        await ctx.sync();
      }
      // Форматы по столбцам: даты в ключах и показателях не превращаются в числа.
      plan.formats.forEach((format, column) => {
        if (format && format !== "General" && format !== "Общий") sheet.getRangeByIndexes(1, column, plan.rows.length, 1).numberFormat = Array.from({ length: plan.rows.length }, () => [format]) as any;
      });
      sheet.getRangeByIndexes(0, 0, 1, width).format.font.bold = true;
      sheet.freezePanes.freezeRows(1);
      await ctx.sync();
      await fitNewTable(ctx, sheet, 0, width);
    } catch (error: any) {
      throw new ToolExecutionError(`Запись на лист «${plan.destSheet}» прервалась: ${error?.message ?? error}. Часть строк могла записаться.`, "unknown");
    }
    const address = `A1:${columnLetters(width)}${total}`;
    const written = sheet.getRange(address);
    written.load(["values", "formulas"]);
    await ctx.sync();
    const back = written.values as unknown[][];
    const sum = back.slice(1).reduce((total, row) => total + (typeof row[width - 1] === "number" ? row[width - 1] as number : 0), 0);
    const problems: string[] = [];
    if (back.length !== total) problems.push(`строк ${back.length - 1} вместо ${plan.rows.length}`);
    if (Math.abs(sum - plan.sum) > 1e-6 * Math.max(1, Math.abs(plan.sum))) problems.push(`сумма значений ${sum} вместо ${plan.sum}`);

    let undoRecorded = false;
    if (plan.undoAvailable) {
      const signature = JSON.stringify(written.formulas);
      const sheetId = sheet.id;
      undoRecorded = push(action(`длинная таблица на листе ${plan.destSheet}`, async () => {
        const revision = getStructuralRevision();
        await Excel.run(async (undoCtx) => {
          const undoSheet = undoCtx.workbook.worksheets.getItem(sheetId);
          const range = undoSheet.getRange(address);
          range.load("formulas");
          await undoCtx.sync();
          if (JSON.stringify(range.formulas) !== signature) throw new Error(`Лист «${plan.destSheet}» изменили после разворота. Отмена остановлена, чтобы не затереть правку.`);
          if (getStructuralRevision() !== revision) throw new Error("Структура книги изменилась во время отмены. Отмена остановлена.");
          undoSheet.delete();
          await undoCtx.sync();
        });
      }));
    }
    if (problems.length) {
      throw new ToolExecutionError(`Длинная таблица записана на «${plan.destSheet}», но сверка расходится: ${problems.join("; ")}.` + (undoRecorded ? " Её можно убрать кнопкой «Отменить»." : ""), "applied");
    }
    return {
      ok: true,
      executionState: "verified",
      source: `${plan.sourceSheet}!${plan.sourceAddress}`,
      sheet: plan.destSheet,
      address,
      rows: plan.rows.length,
      header: plan.header,
      ...(plan.skipped ? { skippedBlanks: plan.skipped } : {}),
      sum: plan.sum,
      note: "Сверено: число строк и сумма значений совпадают с исходной таблицей. Дальше можно сделать таблицу Excel (create_table) и сводную по ней.",
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: "Автоматическая отмена недоступна: лист удаляется вручную." })
    };
  });
}
