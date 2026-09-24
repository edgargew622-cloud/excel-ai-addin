/**
 * Вставка и удаление столбцов (этап 7, 7.3.2).
 *
 * Те же правила риска, что у строк этапа 6: отката нет, история отмены
 * очищается, последствия для формул книги называются до операции и
 * сверяются после. Замер в Excel 24 сентября 2026 года: удаление столбца C
 * превращает `=C2*10` и `=SUM(C:C)` в `#ССЫЛКА!`, молча сужает `SUM(B:D)`
 * до `B:C` и `SUM(A2:E2)` до `A2:D2`; вставка внутрь диапазона его
 * расширяет, а вплотную к краю — нет, и итог молча не охватывает новый
 * столбец.
 */

import {
  assertPlanWorkbook,
  collectRowRisks,
  deepFreeze,
  describeTableChanges,
  MAX_IO_CELLS,
  preflightToolArgs,
  probeMergedAreas,
  readTableRanges,
  REF_ERROR,
  scanWorkbookFormulas,
  ToolError,
  ToolExecutionError,
  type RowFormulaRisk,
  type TableRange
} from "./excelTools";
import { columnLetters } from "./formulaFill";
import { countFilled, countRefErrors, rowBand } from "./rowOps";
import { invalidateAfterStructuralChange } from "./undo";
import { lastWorkbookBackup } from "./workbookBackup";
import { captureTarget, officeCapabilities, type WorkbookTarget } from "./workbookContext";

const MAX_COLUMNS_PER_OP = 100;
const MAX_PREVIEW = 8;
const MAX_BROKEN_LISTED = 20;

export interface ColumnOpPlan {
  readonly kind: "insert_columns" | "delete_columns";
  readonly id: string;
  readonly target: WorkbookTarget;
  /** Первый столбец полосы буквами и номером (A — 1). */
  readonly startColumn: string;
  readonly startIndex: number;
  readonly count: number;
  /** Адрес полосы столбцов, например C:D. */
  readonly columnsAddress: string;
  readonly usedRangeAddress?: string;
  readonly usedRangeRows: number;
  /** Содержимое полосы для предпросмотра — по строкам занятой области. */
  readonly preview: readonly (readonly unknown[])[];
  readonly previewTruncated: boolean;
  readonly filledCells: number;
  readonly bandSignature: string | null;
  readonly formulaRisks: readonly RowFormulaRisk[];
  readonly riskOverflow: number;
  readonly unscannedSheets: readonly string[];
  readonly tableFormulaSheets: readonly string[];
  readonly refErrorsBefore: number;
  readonly tablesBefore: readonly TableRange[];
  readonly tableWarning?: string;
  readonly mergeWarning?: string;
  readonly backup: { name: string; at: string } | null;
  readonly undoAvailable: false;
  readonly undoNote: string;
  readonly createdAt: string;
}

/** Номер столбца по буквам: A — 1, AA — 27; null — не буквы столбца. */
export function columnNumber(letters: unknown): number | null {
  if (typeof letters !== "string" || !/^[A-Za-z]{1,3}$/.test(letters.trim())) return null;
  const value = [...letters.trim().toUpperCase()].reduce((total, ch) => total * 26 + ch.charCodeAt(0) - 64, 0);
  return value >= 1 && value <= 16_384 ? value : null;
}

/** Таблицы, которые полоса столбцов задевает: у таблицы Excel свои правила столбцов. */
function tablesInBand(tables: readonly TableRange[], start: number, end: number): string[] {
  return tables.filter((table) => {
    const [from, to = from] = table.address.replace(/^.*!/, "").replace(/\$/g, "").split(":");
    const a = columnNumber(/^[A-Z]+/i.exec(from)?.[0]);
    const b = columnNumber(/^[A-Z]+/i.exec(to)?.[0]);
    return a !== null && b !== null && a <= end && b >= start;
  }).map((table) => table.name);
}

async function prepareColumnOpPlan(mode: "insert_columns" | "delete_columns", args: unknown): Promise<ColumnOpPlan> {
  preflightToolArgs(mode, args);
  const a = args as { sheet?: string; startColumn: string; count: number };
  const start = columnNumber(a.startColumn);
  if (start === null) throw new ToolError(`startColumn — буквы столбца, например C; получено «${a.startColumn}».`);
  const count = a.count;
  if (!Number.isInteger(count) || count < 1 || count > MAX_COLUMNS_PER_OP) throw new ToolError(`count — от 1 до ${MAX_COLUMNS_PER_OP}.`);
  if (start + count - 1 > 16_384) throw new ToolError("Полоса выходит за последний столбец листа.");
  const band = rowBand(start, count);
  const address = `${columnLetters(start)}:${columnLetters(start + count - 1)}`;
  const target = await captureTarget(a.sheet);

  const prepared = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    sheet.load(["id", "name"]);
    try { sheet.protection?.load("protected"); } catch { /* среда без сведений о защите */ }
    const used = officeCapabilities().usedRangeOrNull ? sheet.getUsedRangeOrNullObject(true) : sheet.getUsedRange(true);
    used.load(["address", "rowIndex", "columnIndex", "rowCount", "columnCount", "isNullObject"]);
    await ctx.sync();
    if (sheet.protection?.protected) {
      throw new ToolError(`Лист ${sheet.name} защищён: столбцы вставить или удалить нельзя, и операция не выполнялась. Снимите защиту листа.`);
    }
    const empty = Boolean((used as any).isNullObject);
    const tables = await readTableRanges(ctx, sheet);
    // Столбец через таблицу Excel меняет саму таблицу — её столбцы и формулы.
    // Здесь это не поддержано: отказ до карточки, а не неожиданность после.
    const touched = tablesInBand(tables, band.startRow, band.endRow);
    if (touched.length) {
      throw new ToolError(`Столбцы ${address} проходят через таблицу Excel ${touched.map((name) => `«${name}»`).join(", ")}: это меняет саму таблицу и здесь не поддержано. Операция не выполнялась.`);
    }

    // Содержимое полосы — в пределах занятых строк: целый столбец — миллион ячеек.
    let preview: unknown[][] = [];
    let previewTruncated = false;
    let filledCells = 0;
    let bandSignature: string | null = null;
    const overlapStart = Math.max(band.startRow, empty ? 1 : used.columnIndex + 1);
    const overlapEnd = Math.min(band.endRow, empty ? 0 : used.columnIndex + used.columnCount);
    if (!empty && overlapEnd >= overlapStart) {
      const columns = overlapEnd - overlapStart + 1;
      if (columns * used.rowCount <= MAX_IO_CELLS) {
        const bandRange = sheet.getRangeByIndexes(used.rowIndex, overlapStart - 1, used.rowCount, columns);
        bandRange.load(["values", "formulas"]);
        await ctx.sync();
        const values = bandRange.values as unknown[][];
        filledCells = countFilled(values);
        bandSignature = JSON.stringify(bandRange.formulas);
        preview = values.slice(0, MAX_PREVIEW).map((row) => row.slice(0, MAX_PREVIEW));
        previewTruncated = values.length > MAX_PREVIEW || columns > MAX_PREVIEW;
      } else {
        previewTruncated = true;
      }
    }

    const probeRows = Math.max(1, Math.min(empty ? 1 : used.rowCount, 100));
    const bandForProbe = sheet.getRangeByIndexes(empty ? 0 : used.rowIndex, start - 1, probeRows, count);
    bandForProbe.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
    await ctx.sync();
    const merged = await probeMergedAreas(ctx, sheet, bandForProbe);
    const scan = await scanWorkbookFormulas(ctx);
    const { risks, overflow, tableFormulaSheets } = collectRowRisks(mode, scan.sheets, sheet.name, band);
    const backup = lastWorkbookBackup();

    return {
      kind: mode,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      target: { ...target, sheetName: sheet.name },
      startColumn: columnLetters(start),
      startIndex: start,
      count,
      columnsAddress: address,
      ...(empty ? {} : { usedRangeAddress: String(used.address) }),
      usedRangeRows: empty ? 0 : used.rowCount,
      preview,
      previewTruncated,
      filledCells,
      bandSignature,
      formulaRisks: risks,
      riskOverflow: overflow,
      unscannedSheets: scan.unscanned,
      tableFormulaSheets,
      refErrorsBefore: scan.sheets.reduce((total, item) => total + countRefErrors(item.values), 0),
      tablesBefore: tables,
      ...(merged.areas.length > 0 || merged.unresolvedAnchors.length > 0
        ? { mergeWarning: "Полосу задевают объединённые ячейки. Excel может отказать в операции или разорвать объединение." }
        : {}),
      backup: backup ? { name: backup.name, at: backup.at } : null,
      undoAvailable: false as const,
      undoNote: mode === "delete_columns"
        ? "Отмены нет: удалённые столбцы не восстанавливаются ни кнопкой отмены панели, ни повтором операции. " +
          (backup ? `Вернуться можно только к резервной копии ${backup.name}.` : "Резервной копии в этом сеансе не создавалось — возвращаться будет не к чему.")
        : "Отмены нет: вставленные столбцы панель удалить обратно не может, а вся предыдущая история отмены будет очищена.",
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared) as ColumnOpPlan;
}

export const prepareInsertColumnsPlan = (args: unknown) => prepareColumnOpPlan("insert_columns", args);
export const prepareDeleteColumnsPlan = (args: unknown) => prepareColumnOpPlan("delete_columns", args);

export async function executeColumnOpPlan(plan: ColumnOpPlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(plan.target.sheetId);
    sheet.load(["id", "name"]);
    const used = officeCapabilities().usedRangeOrNull ? sheet.getUsedRangeOrNullObject(true) : sheet.getUsedRange(true);
    used.load(["rowIndex", "columnIndex", "rowCount", "columnCount", "isNullObject"]);
    await ctx.sync();

    // Отката нет: правка полосы после предпросмотра обязана отменить операцию.
    if (plan.bandSignature !== null) {
      const empty = Boolean((used as any).isNullObject);
      const overlapStart = Math.max(plan.startIndex, empty ? 1 : used.columnIndex + 1);
      const overlapEnd = Math.min(plan.startIndex + plan.count - 1, empty ? 0 : used.columnIndex + used.columnCount);
      const columns = overlapEnd - overlapStart + 1;
      if (empty || columns < 1 || used.rowCount !== plan.usedRangeRows) {
        throw new ToolExecutionError(`Занятая область листа ${sheet.name} изменилась после предпросмотра. Столбцы не трогали — сделайте новый предпросмотр.`, "failed_before_write");
      }
      const bandRange = sheet.getRangeByIndexes(used.rowIndex, overlapStart - 1, used.rowCount, columns);
      bandRange.load("formulas");
      await ctx.sync();
      if (JSON.stringify(bandRange.formulas) !== plan.bandSignature) {
        throw new ToolExecutionError(`Столбцы ${sheet.name}!${plan.columnsAddress} изменились после предпросмотра. Операция не выполнялась — сделайте новый предпросмотр.`, "failed_before_write");
      }
    }

    const range = sheet.getRange(plan.columnsAddress);
    try {
      if (plan.kind === "insert_columns") range.insert(Excel.InsertShiftDirection.right);
      else range.delete(Excel.DeleteShiftDirection.left);
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(
        `Excel отказал в операции со столбцами ${sheet.name}!${plan.columnsAddress}: ${error?.message ?? error}. ` +
        "Неизвестно, успела ли она примениться — перечитайте лист, прежде чем что-либо менять.",
        "unknown"
      );
    }

    // Сами столбцы после операции всегда выглядят правильно: смотреть надо на книгу.
    const scan = await scanWorkbookFormulas(ctx);
    const refErrorsAfter = scan.sheets.reduce((total, item) => total + countRefErrors(item.values), 0);
    const newRefErrors = refErrorsAfter - plan.refErrorsBefore;
    const brokenCells: string[] = [];
    for (const item of scan.sheets) {
      item.values.forEach((row, r) => row.forEach((value, c) => {
        if (typeof value === "string" && REF_ERROR.test(value.trim()) && brokenCells.length < MAX_BROKEN_LISTED) {
          brokenCells.push(`${item.name}!${columnLetters(item.columnIndex + c + 1)}${item.rowIndex + r + 1}`);
        }
      }));
    }
    const tableChanges = describeTableChanges(plan.tablesBefore, await readTableRanges(ctx, sheet));
    const invalidatedUndo = invalidateAfterStructuralChange();
    const predictedBroken = plan.formulaRisks.filter((risk) => risk.kind === "broken").length + plan.riskOverflow;

    const result = {
      ok: true,
      executionState: "verified",
      sheet: sheet.name,
      columns: plan.columnsAddress,
      ...(plan.kind === "insert_columns" ? { inserted: plan.count, at: plan.startColumn } : { deleted: plan.count, from: plan.startColumn, lostFilledCells: plan.filledCells }),
      refErrorsBefore: plan.refErrorsBefore,
      refErrorsAfter,
      ...(plan.formulaRisks.length
        ? {
            affectedFormulas: plan.formulaRisks,
            ...(plan.riskOverflow > 0 ? { affectedFormulasOmitted: plan.riskOverflow } : {}),
            affectedFormulasNote: plan.kind === "insert_columns"
              ? "Эти формулы не охватывают вставленные столбцы: ошибки не будет, итог просто посчитан без них. Назови их пользователю поимённо."
              : "Эти формулы ссылались на удалённые столбцы. Перечисли их пользователю: часть станет #ССЫЛКА!, а суженные диапазоны молча считают по меньшему числу столбцов."
          }
        : {}),
      ...(plan.tableFormulaSheets.length ? { tableFormulaSheets: plan.tableFormulaSheets, tableFormulaNote: "На этих листах есть формулы со ссылками на таблицы; они не разбирались." } : {}),
      ...(newRefErrors > 0 ? { newRefErrors, brokenCells, refNote: `В книге появилось ${newRefErrors} ошибок ссылок — это нужно назвать пользователю.` } : {}),
      ...(scan.unscanned.length ? { unscannedSheets: scan.unscanned, scanNote: "Эти листы слишком велики для обхода формул: про них ничего не проверено." } : {}),
      ...(tableChanges.length ? { tableChanges, tableNote: "Excel изменил границы таблицы из-за этой операции; в отчёте это нужно назвать." } : {}),
      undoable: false,
      undoNote: plan.undoNote,
      invalidatedUndo
    };
    // Больше новых ошибок ссылок, чем предсказал разбор, — разбор нашёл не всё.
    if (newRefErrors > predictedBroken) {
      throw new ToolExecutionError(
        `Операция со столбцами ${sheet.name}!${plan.columnsAddress} выполнена, но новых ошибок ссылок ${newRefErrors}, а предсказано ${predictedBroken}` +
        `${brokenCells.length ? ` (${brokenCells.slice(0, 5).join(", ")})` : ""}: разбор формул нашёл не всё. Проверьте книгу; отмены нет.`,
        "applied"
      );
    }
    return result;
  });
}
