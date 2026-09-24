/**
 * Группировка строк и столбцов (этап 7, 7.3.3).
 *
 * Настоящая группа Excel (структура с кнопками «+» и «−»), а не скрытие.
 * Уровни группировки API не отдаёт, поэтому группа доказывается действием,
 * как её проверил бы человек: свернуть только эту группу и убедиться, что
 * строки скрылись. Замер 24 сентября 2026 года (ExcelApi 1.10):
 *
 * - свернуть несгруппированные строки нельзя — ничего не скрывается;
 *   поэтому свёртка и отличает группу от её отсутствия;
 * - разворот той же полосы строки обратно не показывает, а после
 *   разгруппировки они так и остаются скрытыми. Поэтому видимость после
 *   проверки возвращается построчно — ровно такой, какой была до операции;
 * - вложенная группа сворачивается отдельно от внешней.
 */

import { assertPlanWorkbook, deepFreeze, preflightToolArgs, ToolError, ToolExecutionError } from "./excelTools";
import { columnLetters } from "./formulaFill";
import { action, isCustomUndoAvailable, push } from "./undo";
import { captureTarget, type WorkbookTarget } from "./workbookContext";

/** Столько строк или столбцов за операцию: видимость читается по одной. */
export const MAX_OUTLINE_LINES = 1000;

export interface OutlineBand {
  axis: "rows" | "columns";
  /** Номера первой и последней строки или столбца, с 1. */
  start: number;
  end: number;
  /** Адрес полосы: «3:5» или «C:D». */
  address: string;
}

const letterNumber = (text: string) => [...text.toUpperCase()].reduce((total, ch) => total * 26 + ch.charCodeAt(0) - 64, 0);

/** Полоса по адресу: «3:5» — строки, «C:D» — столбцы. */
export function parseOutlineBand(address: unknown): OutlineBand {
  const text = String(address ?? "").trim().replace(/\$/g, "");
  const rows = /^(\d{1,7}):(\d{1,7})$/.exec(text);
  if (rows) {
    const [start, end] = [Number(rows[1]), Number(rows[2])].sort((a, b) => a - b);
    if (start < 1 || end > 1_048_576) throw new ToolError(`Строки ${text} вне листа.`);
    return { axis: "rows", start, end, address: `${start}:${end}` };
  }
  const columns = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(text);
  if (columns) {
    const [start, end] = [letterNumber(columns[1]), letterNumber(columns[2])].sort((a, b) => a - b);
    if (end > 16_384) throw new ToolError(`Столбцы ${text} вне листа.`);
    return { axis: "columns", start, end, address: `${columnLetters(start)}:${columnLetters(end)}` };
  }
  throw new ToolError(`Полоса для группировки — целые строки «3:5» или столбцы «C:D»; получено «${text}».`);
}

export interface GroupPlan {
  readonly kind: "group_rows_columns";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly band: OutlineBand;
  /** Свернуть группу после создания. */
  readonly collapse: boolean;
  /** Видимость каждой строки или столбца до операции — её же и вернуть. */
  readonly hiddenBefore: readonly boolean[];
  readonly undoAvailable: boolean;
  readonly undoNote: string;
  readonly createdAt: string;
}

function line(sheet: Excel.Worksheet, band: OutlineBand, index: number): Excel.Range {
  const n = band.start + index;
  return sheet.getRange(band.axis === "rows" ? `${n}:${n}` : `${columnLetters(n)}:${columnLetters(n)}`);
}

async function readHidden(ctx: Excel.RequestContext, sheet: Excel.Worksheet, band: OutlineBand): Promise<boolean[]> {
  const lines = Array.from({ length: band.end - band.start + 1 }, (_, index) => {
    const range = line(sheet, band, index);
    range.load(band.axis === "rows" ? "rowHidden" : "columnHidden");
    return range;
  });
  await ctx.sync();
  return lines.map((range) => Boolean(band.axis === "rows" ? range.rowHidden : range.columnHidden));
}

async function setHidden(ctx: Excel.RequestContext, sheet: Excel.Worksheet, band: OutlineBand, hidden: readonly boolean[]): Promise<void> {
  hidden.forEach((value, index) => {
    const range = line(sheet, band, index);
    if (band.axis === "rows") range.rowHidden = value;
    else range.columnHidden = value;
  });
  await ctx.sync();
}

const option = (band: OutlineBand) => (band.axis === "rows" ? "ByRows" : "ByColumns") as any;

export async function prepareGroupPlan(args: unknown): Promise<GroupPlan> {
  preflightToolArgs("group_rows_columns", args);
  const a = args as { sheet?: string; address: string; collapse?: boolean };
  const band = parseOutlineBand(a.address);
  if (band.end - band.start + 1 > MAX_OUTLINE_LINES) throw new ToolError(`Группа до ${MAX_OUTLINE_LINES} строк или столбцов за операцию.`);
  const target = await captureTarget(a.sheet);
  const prepared = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    sheet.load(["id", "name"]);
    try { sheet.protection?.load("protected"); } catch { /* среда без сведений о защите */ }
    await ctx.sync();
    if (sheet.protection?.protected) throw new ToolError(`Лист ${sheet.name} защищён: группировать нельзя. Операция не выполнялась.`);
    const hiddenBefore = await readHidden(ctx, sheet, band);
    const undo = isCustomUndoAvailable();
    return {
      kind: "group_rows_columns" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      target: { ...target, sheetName: sheet.name },
      band,
      collapse: a.collapse === true,
      hiddenBefore,
      undoAvailable: undo,
      undoNote: undo
        ? "Отмена снимет один уровень группировки с этой полосы и вернёт прежнюю видимость."
        : "Отмена недоступна: монитор изменений Excel не активен. Группу можно снять вручную.",
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executeGroupPlan(plan: GroupPlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(plan.target.sheetId);
    sheet.load(["id", "name"]);
    await ctx.sync();
    const band = plan.band;
    const where = `${sheet.name}!${band.address}`;
    const hiddenNow = await readHidden(ctx, sheet, band);
    if (JSON.stringify(hiddenNow) !== JSON.stringify(plan.hiddenBefore)) {
      throw new ToolExecutionError(`Видимость ${where} изменилась после предпросмотра. Группа не создавалась — сделайте новый предпросмотр.`, "failed_before_write");
    }
    const range = sheet.getRange(band.address);
    try {
      range.group(option(band));
      await ctx.sync();
    } catch (error: any) {
      // Excel откажет, например, на девятом уровне вложенности.
      throw new ToolExecutionError(`Excel отказал в группировке ${where}: ${error?.message ?? error}. Проверьте структуру листа.`, "unknown");
    }

    // Группа доказывается свёрткой: без неё строки не скрылись бы.
    let proven: boolean | null;
    try {
      range.hideGroupDetails(option(band));
      await ctx.sync();
      const collapsed = await readHidden(ctx, sheet, band);
      const visibleBefore = plan.hiddenBefore.map((hidden, index) => ({ hidden, index })).filter((item) => !item.hidden);
      proven = visibleBefore.length ? visibleBefore.every((item) => collapsed[item.index]) : null;
      // Разворот той же полосы видимость не возвращает — возвращаем сами.
      await setHidden(ctx, sheet, band, plan.collapse ? collapsed.map(() => true) : plan.hiddenBefore);
    } catch (error: any) {
      throw new ToolExecutionError(`Группа ${where} создана, но проверить её свёрткой не удалось: ${error?.message ?? error}. Проверьте видимость строк.`, "applied");
    }
    const finalHidden = await readHidden(ctx, sheet, band);
    const expected = plan.collapse ? finalHidden.map(() => true) : plan.hiddenBefore;
    const visibilityRestored = JSON.stringify(finalHidden) === JSON.stringify(expected);

    let undoRecorded = false;
    if (plan.undoAvailable) {
      undoRecorded = push(action(`группировка ${where}`, async () => {
        await Excel.run(async (undoCtx) => {
          const undoSheet = undoCtx.workbook.worksheets.getItem(plan.target.sheetId);
          undoSheet.getRange(band.address).ungroup(option(band));
          await undoCtx.sync();
          await setHidden(undoCtx, undoSheet, band, plan.hiddenBefore);
          const restored = await readHidden(undoCtx, undoSheet, band);
          if (JSON.stringify(restored) !== JSON.stringify(plan.hiddenBefore)) {
            throw new Error(`Группа ${where} снята, но видимость строк не вернулась к прежней. Проверьте её вручную.`);
          }
        });
      }));
    }

    if (proven === false || !visibilityRestored) {
      throw new ToolExecutionError(
        (proven === false ? `После группировки ${where} свёртка не скрыла строки — группа, похоже, не создалась. ` : "") +
        (!visibilityRestored ? `Видимость ${where} после проверки не совпала с ожидаемой. ` : "") +
        (undoRecorded ? "Операцию можно отменить кнопкой «Отменить»." : "Проверьте лист."),
        "applied"
      );
    }
    const unit = band.axis === "rows" ? "строк" : "столбцов";
    return {
      ok: true,
      executionState: proven === null ? "applied" : "verified",
      sheet: sheet.name,
      grouped: band.address,
      axis: band.axis,
      collapsed: plan.collapse,
      note: proven === null
        ? `Все ${unit} полосы были скрыты ещё до операции: свёрткой доказать группу нельзя, она не проверена.`
        : `Группа проверена свёрткой: ${unit} скрывались. ${plan.collapse ? "Группа оставлена свёрнутой." : "Видимость возвращена прежней."}`,
      existingGroupsNote: "Уровни прежних групп Excel через API не сообщает: если на этой полосе уже была группа, новая добавилась к ней уровнем.",
      undoable: undoRecorded,
      undoNote: plan.undoNote
    };
  });
}
