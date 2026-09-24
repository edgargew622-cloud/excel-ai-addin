/**
 * Проверяемые шаблоны расчёта (этап 7, 7.5.2): доли и рост.
 *
 * Шаблон пишет формулы, а не числа, посчитанные моделью: блок ссылается на
 * исходную таблицу и пересчитывается вместе с ней. Панель сама считает, что
 * должно получиться, и после записи сверяет каждую ячейку с этим расчётом.
 * Контрольное равенство — сумма долей каждого периода равна 100 % — входит
 * в сверку: без него «готово» не выдаётся.
 *
 * Рост от нуля или от пустого значения не определён: ячейка остаётся пустой,
 * и ответ называет такие места, а не прячет их.
 */

import { parseA1Rect } from "./a1";
import {
  assertPlanWorkbook,
  assertTargetWritable,
  checkAddress,
  deepFreeze,
  preflightToolArgs,
  rangeOf,
  ToolError,
  ToolExecutionError
} from "./excelTools";
import { columnLetters } from "./formulaFill";
import { action, isCustomUndoAvailable, push } from "./undo";
import { captureTarget, type WorkbookTarget } from "./workbookContext";

const TOLERANCE = 1e-9;
const MAX_ITEMS = 200;
const MAX_PERIODS = 40;

export interface SourceBlock {
  /** Номер первой строки (шапки) и первого столбца (подписей), с 1. */
  top: number;
  left: number;
  labels: string[];
  periods: unknown[];
  /** Числа: строки — статьи, столбцы — периоды. null — пусто. */
  numbers: (number | null)[][];
}

export interface TemplateCell {
  formula: string;
  /** Что должно получиться: число, "" для пустого, строка для подписи. */
  expected: number | string | null;
}

export interface TemplateLayout {
  top: number;
  left: number;
  rows: TemplateCell[][];
  /** Строка контроля «сумма долей» внутри блока, с 0. */
  controlRow: number;
  /** Ячейки роста, оставленные пустыми: база ноль или пусто. */
  undefinedGrowth: string[];
  percentRows: number[];
}

/** Исходная таблица: шапка с периодами, подписи статей слева, числа. */
export function readSourceBlock(values: readonly (readonly unknown[])[], top: number, left: number): SourceBlock {
  if (values.length < 2 || (values[0]?.length ?? 0) < 3) {
    throw new ToolError("Нужна таблица: первая строка — периоды, первый столбец — статьи; хотя бы одна строка данных и два периода.");
  }
  const periods = values[0].slice(1);
  if (periods.length > MAX_PERIODS) throw new ToolError(`Периодов больше ${MAX_PERIODS}: разбейте таблицу.`);
  const items = values.slice(1);
  if (items.length > MAX_ITEMS) throw new ToolError(`Статей больше ${MAX_ITEMS}: разбейте таблицу.`);
  const labels = items.map((row) => String(row[0] ?? "").trim());
  const numbers = items.map((row, r) => row.slice(1).map((value, c) => {
    if (value === "" || value === null || value === undefined) return null;
    if (typeof value !== "number") {
      throw new ToolError(`В ${columnLetters(left + c + 1)}${top + r + 1} не число: «${String(value)}». Доли и рост считаются по числам — сначала приведите данные (convert_values).`);
    }
    return value;
  }));
  if (labels.some((label) => !label)) throw new ToolError("У каждой строки данных нужна подпись статьи в первом столбце.");
  return { top, left, labels, periods, numbers };
}

/**
 * Блок долей и роста под исходной таблицей: формулы и ожидаемые значения.
 * Все адреса — абсолютные строки исходника, чтобы блок читался без догадок.
 */
export function shareGrowthLayout(source: SourceBlock, top: number, left: number): TemplateLayout {
  const n = source.labels.length;
  const firstRow = source.top + 1;
  const lastRow = source.top + n;
  const col = (index: number) => columnLetters(source.left + 1 + index);
  const labelCol = columnLetters(source.left);
  const header = (): TemplateCell[] => [
    { formula: "", expected: "" },
    // Пустая шапка по ссылке дала бы 0: такую ячейку оставляем пустой.
    ...source.periods.map((period, j) => (period === "" || period === null || period === undefined
      ? { formula: "", expected: "" }
      : { formula: `=${col(j)}$${source.top}`, expected: period as number | string }))
  ];
  const label = (i: number): TemplateCell => ({ formula: `=$${labelCol}${firstRow + i}`, expected: source.labels[i] });
  const totals = source.periods.map((_, j) => source.numbers.reduce((sum, row) => sum + (row[j] ?? 0), 0));
  const rows: TemplateCell[][] = [];
  const percentRows: number[] = [];
  const undefinedGrowth: string[] = [];

  rows.push([{ formula: "Доля в итоге периода", expected: "Доля в итоге периода" }, ...source.periods.map(() => ({ formula: "", expected: "" }))]);
  rows.push(header());
  for (let i = 0; i < n; i++) {
    percentRows.push(rows.length);
    rows.push([label(i), ...source.periods.map((_, j) => {
      const sum = `SUM(${col(j)}$${firstRow}:${col(j)}$${lastRow})`;
      return {
        formula: `=IF(${sum}=0,"",${col(j)}${firstRow + i}/${sum})`,
        expected: totals[j] === 0 ? "" : (source.numbers[i][j] ?? 0) / totals[j]
      };
    })]);
  }
  const controlRow = rows.length;
  percentRows.push(rows.length);
  rows.push([{ formula: "Контроль: сумма долей", expected: "Контроль: сумма долей" }, ...source.periods.map((_, j) => ({
    // Доли стоят в строках top+2 … top+1+n этого же блока.
    formula: `=SUM(${columnLetters(left + 1 + j)}${top + 2}:${columnLetters(left + 1 + j)}${top + 1 + n})`,
    expected: totals[j] === 0 ? 0 : 1
  }))]);
  rows.push([{ formula: "", expected: "" }, ...source.periods.map(() => ({ formula: "", expected: "" }))]);
  rows.push([{ formula: "Рост к прошлому периоду", expected: "Рост к прошлому периоду" }, ...source.periods.map(() => ({ formula: "", expected: "" }))]);
  rows.push(header());
  const growth = (current: number | null, previous: number | null) =>
    previous === null || previous === 0 ? "" : (current ?? 0) / previous - 1;
  const growthRowStart = rows.length;
  for (let i = 0; i < n; i++) {
    percentRows.push(rows.length);
    rows.push([label(i), ...source.periods.map((_, j) => {
      if (j === 0) return { formula: "", expected: "" };
      const previous = source.numbers[i][j - 1];
      if (previous === null || previous === 0) undefinedGrowth.push(`${columnLetters(left + 1 + j)}${top + growthRowStart + i}`);
      return {
        formula: `=IF(N(${col(j - 1)}${firstRow + i})=0,"",${col(j)}${firstRow + i}/${col(j - 1)}${firstRow + i}-1)`,
        expected: growth(source.numbers[i][j], previous)
      };
    })]);
  }
  percentRows.push(rows.length);
  rows.push([{ formula: "Итого", expected: "Итого" }, ...source.periods.map((_, j) => {
    if (j === 0) return { formula: "", expected: "" };
    const current = `SUM(${col(j)}$${firstRow}:${col(j)}$${lastRow})`;
    const previous = `SUM(${col(j - 1)}$${firstRow}:${col(j - 1)}$${lastRow})`;
    return { formula: `=IF(${previous}=0,"",${current}/${previous}-1)`, expected: totals[j - 1] === 0 ? "" : totals[j] / totals[j - 1] - 1 };
  })]);
  return { top, left, rows, controlRow, undefinedGrowth, percentRows };
}

/** Совпадает ли прочитанное с ожидаемым: числа — с допуском, подписи — как текст. */
export function templateMismatches(layout: TemplateLayout, values: readonly (readonly unknown[])[]): string[] {
  const problems: string[] = [];
  layout.rows.forEach((row, r) => row.forEach((cell, c) => {
    const actual = values[r]?.[c];
    const name = `${columnLetters(layout.left + c)}${layout.top + r}`;
    if (typeof cell.expected === "number") {
      if (typeof actual !== "number" || Math.abs(actual - cell.expected) > TOLERANCE * Math.max(1, Math.abs(cell.expected))) {
        problems.push(`${name}: ${String(actual)} вместо ${cell.expected}`);
      }
    } else if (String(actual ?? "") !== String(cell.expected ?? "")) {
      problems.push(`${name}: «${String(actual)}» вместо «${String(cell.expected)}»`);
    }
  }));
  return problems;
}

/* --- план ------------------------------------------------------------------------ */

export interface ShareGrowthPlan {
  readonly kind: "add_share_growth";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly sourceAddress: string;
  readonly sourceSignature: string;
  readonly destAddress: string;
  readonly layout: TemplateLayout;
  readonly items: number;
  readonly periods: number;
  readonly numberFormatsBefore: unknown[][];
  readonly undoAvailable: boolean;
  readonly undoNote: string;
  readonly createdAt: string;
}

export async function prepareShareGrowthPlan(args: unknown): Promise<ShareGrowthPlan> {
  preflightToolArgs("add_share_growth", args);
  const a = args as { sheet?: string; address: string; destAddress?: string };
  const address = checkAddress(a.address);
  const target = await captureTarget(a.sheet);
  const prepared = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    const range = await rangeOf(ctx, sheet, address);
    range.load(["address", "rowIndex", "columnIndex", "rowCount", "columnCount", "values", "formulas"]);
    sheet.load(["id", "name"]);
    try { sheet.protection?.load(["protected", "options"]); } catch { /* среда без сведений о защите */ }
    await ctx.sync();
    const source = readSourceBlock(range.values as unknown[][], range.rowIndex + 1, range.columnIndex + 1);
    const rowsNeeded = 2 * source.labels.length + 7;
    let top = range.rowIndex + range.rowCount + 2;
    let left = range.columnIndex + 1;
    if (a.destAddress) {
      const rect = parseA1Rect(checkAddress(a.destAddress));
      if (!rect) throw new ToolError(`destAddress «${a.destAddress}» — не адрес ячейки.`);
      top = rect.rowStart;
      left = rect.columnStart;
    }
    const layout = shareGrowthLayout(source, top, left);
    const destAddress = `${columnLetters(left)}${top}:${columnLetters(left + source.periods.length)}${top + rowsNeeded - 1}`;
    const dest = sheet.getRange(destAddress);
    dest.load(["formulas", "numberFormat"]);
    try { dest.format?.protection?.load("locked"); } catch { /* нет сведений */ }
    const used = sheet.getUsedRangeOrNullObject(true);
    used.load(["isNullObject", "rowIndex", "rowCount"]);
    await ctx.sync();
    assertTargetWritable(sheet, dest);
    const busy = (dest.formulas as unknown[][]).some((row) => row.some((value) => value !== "" && value !== null));
    if (busy) {
      const free = used.isNullObject ? top : used.rowIndex + used.rowCount + 2;
      throw new ToolError(`Место под блок ${destAddress} занято. Свободно ниже данных листа: начиная с ${columnLetters(left)}${free} — передайте его в destAddress.`);
    }
    const undo = isCustomUndoAvailable();
    return {
      kind: "add_share_growth" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      target: { ...target, sheetName: sheet.name },
      sourceAddress: String(range.address).replace(/^.*!/, ""),
      sourceSignature: JSON.stringify(range.formulas),
      destAddress,
      layout,
      items: source.labels.length,
      periods: source.periods.length,
      numberFormatsBefore: dest.numberFormat as unknown[][],
      undoAvailable: undo,
      undoNote: undo ? "Отмена очистит блок и вернёт прежний формат чисел." : "Отмена недоступна: монитор изменений Excel не активен.",
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executeShareGrowthPlan(plan: ShareGrowthPlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(plan.target.sheetId);
    sheet.load("name");
    const source = sheet.getRange(plan.sourceAddress);
    const dest = sheet.getRange(plan.destAddress);
    source.load("formulas");
    dest.load("formulas");
    await ctx.sync();
    const where = `${sheet.name}!${plan.destAddress}`;
    if (JSON.stringify(source.formulas) !== plan.sourceSignature) {
      throw new ToolExecutionError(`Исходная таблица ${sheet.name}!${plan.sourceAddress} изменилась после предпросмотра. Блок не записывался — сделайте новый предпросмотр.`, "failed_before_write");
    }
    if ((dest.formulas as unknown[][]).some((row) => row.some((value) => value !== "" && value !== null))) {
      throw new ToolExecutionError(`Место ${where} заняли после предпросмотра. Блок не записывался.`, "failed_before_write");
    }
    const formulas = plan.layout.rows.map((row) => row.map((cell) => cell.formula));
    const formats = plan.layout.rows.map((row, r) => row.map((_, c) =>
      plan.layout.percentRows.includes(r) && c > 0 ? "0.0%" : (plan.numberFormatsBefore[r]?.[c] as string ?? "General")));
    try {
      dest.formulas = formulas as any[][];
      dest.numberFormat = formats as any[][];
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(`Не удалось определить, записался ли блок на ${where}: ${error?.message ?? error}. Перечитайте область.`, "unknown");
    }
    dest.load(["values", "formulas"]);
    await ctx.sync();

    let undoRecorded = false;
    if (plan.undoAvailable) {
      const written = JSON.stringify(dest.formulas);
      undoRecorded = push(action(`доли и рост ${where}`, async () => {
        await Excel.run(async (undoCtx) => {
          const range = undoCtx.workbook.worksheets.getItem(plan.target.sheetId).getRange(plan.destAddress);
          range.load("formulas");
          await undoCtx.sync();
          if (JSON.stringify(range.formulas) !== written) throw new Error(`Блок ${where} изменили после операции агента. Отмена остановлена, чтобы не затереть правку.`);
          range.formulas = plan.layout.rows.map((row) => row.map(() => "")) as any[][];
          range.numberFormat = plan.numberFormatsBefore as any[][];
          await undoCtx.sync();
          range.load("formulas");
          await undoCtx.sync();
          if ((range.formulas as unknown[][]).some((row) => row.some((value) => value !== ""))) throw new Error(`Отмена очистила ${where} не полностью. Проверьте область.`);
        });
      }));
    }

    const values = dest.values as unknown[][];
    const mismatches = templateMismatches(plan.layout, values);
    const control = (values[plan.layout.controlRow] ?? []).slice(1);
    if (mismatches.length) {
      throw new ToolExecutionError(
        `Блок записан на ${where}, но значения расходятся с расчётом панели: ${mismatches.slice(0, 8).join("; ")}${mismatches.length > 8 ? ` и ещё ${mismatches.length - 8}` : ""}. ` +
          `${undoRecorded ? "Его уберёт «Отменить»." : "Проверьте блок."} Готовым он не считается.`,
        "applied"
      );
    }
    return {
      ok: true,
      executionState: "verified",
      address: where,
      source: `${sheet.name}!${plan.sourceAddress}`,
      items: plan.items,
      periods: plan.periods,
      checkedCells: plan.layout.rows.length * (plan.periods + 1),
      control: { row: `${plan.layout.top + plan.layout.controlRow}`, sumOfShares: control, note: "Сумма долей каждого периода сверена: 100 % (0 — если итог периода ноль)." },
      ...(plan.layout.undefinedGrowth.length
        ? { undefinedGrowth: plan.layout.undefinedGrowth, undefinedGrowthNote: "В этих ячейках рост не определён: в прошлом периоде ноль или пусто. Ячейка оставлена пустой — назови это пользователю." }
        : {}),
      note: "Блок — формулы со ссылками на исходную таблицу: он пересчитается при изменении данных. Значения сверены с расчётом панели.",
      undoable: undoRecorded,
      undoNote: plan.undoNote
    };
  });
}
