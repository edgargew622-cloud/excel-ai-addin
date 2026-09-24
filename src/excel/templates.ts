/**
 * Проверяемые шаблоны расчёта (этап 7, 7.5.2): доли и рост, сравнительный анализ.
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
  /** Строка контрольного равенства внутри блока, с 0; null — у шаблона её нет. */
  controlRow: number | null;
  /** Ячейки, оставленные пустыми, потому что величина не определена. */
  undefinedCells: string[];
  /** Формат чисел строки блока, с 0: «0.0%», «0». */
  rowFormats: Record<number, string>;
}

/** Исходная таблица: шапка с периодами, подписи статей слева, числа. */
export function readSourceBlock(values: readonly (readonly unknown[])[], top: number, left: number, minColumns = 2): SourceBlock {
  if (values.length < 2 || (values[0]?.length ?? 0) < minColumns + 1) {
    throw new ToolError(`Нужна таблица: первая строка — шапка, первый столбец — подписи; хотя бы одна строка данных и ${minColumns === 1 ? "один столбец чисел" : `${minColumns} столбца чисел`}.`);
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
  const rowFormats: Record<number, string> = {};
  const undefinedCells: string[] = [];

  rows.push([{ formula: "Доля в итоге периода", expected: "Доля в итоге периода" }, ...source.periods.map(() => ({ formula: "", expected: "" }))]);
  rows.push(header());
  for (let i = 0; i < n; i++) {
    rowFormats[rows.length] = "0.0%";
    rows.push([label(i), ...source.periods.map((_, j) => {
      const sum = `SUM(${col(j)}$${firstRow}:${col(j)}$${lastRow})`;
      return {
        formula: `=IF(${sum}=0,"",${col(j)}${firstRow + i}/${sum})`,
        expected: totals[j] === 0 ? "" : (source.numbers[i][j] ?? 0) / totals[j]
      };
    })]);
  }
  const controlRow = rows.length;
  rowFormats[rows.length] = "0.0%";
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
    rowFormats[rows.length] = "0.0%";
    rows.push([label(i), ...source.periods.map((_, j) => {
      if (j === 0) return { formula: "", expected: "" };
      const previous = source.numbers[i][j - 1];
      if (previous === null || previous === 0) undefinedCells.push(`${columnLetters(left + 1 + j)}${top + growthRowStart + i}`);
      return {
        formula: `=IF(N(${col(j - 1)}${firstRow + i})=0,"",${col(j)}${firstRow + i}/${col(j - 1)}${firstRow + i}-1)`,
        expected: growth(source.numbers[i][j], previous)
      };
    })]);
  }
  rowFormats[rows.length] = "0.0%";
  rows.push([{ formula: "Итого", expected: "Итого" }, ...source.periods.map((_, j) => {
    if (j === 0) return { formula: "", expected: "" };
    const current = `SUM(${col(j)}$${firstRow}:${col(j)}$${lastRow})`;
    const previous = `SUM(${col(j - 1)}$${firstRow}:${col(j - 1)}$${lastRow})`;
    return { formula: `=IF(${previous}=0,"",${current}/${previous}-1)`, expected: totals[j - 1] === 0 ? "" : totals[j] / totals[j - 1] - 1 };
  })]);
  return { top, left, rows, controlRow, undefinedCells, rowFormats };
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

/* --- сравнительный анализ -------------------------------------------------------- */

/** Медиана, как у MEDIAN в Excel: пустые не считаются. */
export function median(numbers: readonly number[]): number {
  const sorted = [...numbers].sort((a, b) => a - b);
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
}

/**
 * Блок сравнения под таблицей «объекты × показатели»: по каждому показателю
 * среднее, медиана, минимум и максимум, отклонение каждого объекта от медианы
 * и место (1 — наибольшее значение; равные значения делят место, как у RANK).
 * Объект без числа по показателю не получает ни отклонения, ни места — это
 * называется, а не прячется за нулём.
 */
export function comparisonLayout(source: SourceBlock, top: number, left: number): TemplateLayout {
  const n = source.labels.length;
  const firstRow = source.top + 1;
  const lastRow = source.top + n;
  const col = (index: number) => columnLetters(source.left + 1 + index);
  const labelCol = columnLetters(source.left);
  const empty = (): TemplateCell => ({ formula: "", expected: "" });
  const title = (text: string): TemplateCell[] => [{ formula: text, expected: text }, ...source.periods.map(empty)];
  const header = (): TemplateCell[] => [empty(), ...source.periods.map((metric, j) => (metric === "" || metric === null || metric === undefined
    ? empty()
    : { formula: `=${col(j)}$${source.top}`, expected: metric as number | string }))];
  const label = (i: number): TemplateCell => ({ formula: `=$${labelCol}${firstRow + i}`, expected: source.labels[i] });
  const columnNumbers = source.periods.map((_, j) => source.numbers.map((row) => row[j]).filter((value): value is number => value !== null));
  const area = (j: number) => `${col(j)}$${firstRow}:${col(j)}$${lastRow}`;

  const rows: TemplateCell[][] = [];
  const rowFormats: Record<number, string> = {};
  const undefinedCells: string[] = [];
  rows.push(title("Сравнение: показатели по всем объектам"));
  rows.push(header());
  const stat = (name: string, fn: string, compute: (numbers: number[]) => number) => {
    rows.push([{ formula: name, expected: name }, ...source.periods.map((_, j) => ({
      formula: `=IF(COUNT(${area(j)})=0,"",${fn}(${area(j)}))`,
      expected: columnNumbers[j].length ? compute(columnNumbers[j]) : ""
    }))]);
  };
  stat("Среднее", "AVERAGE", (numbers) => numbers.reduce((sum, value) => sum + value, 0) / numbers.length);
  stat("Медиана", "MEDIAN", median);
  stat("Минимум", "MIN", (numbers) => Math.min(...numbers));
  stat("Максимум", "MAX", (numbers) => Math.max(...numbers));
  rows.push([empty(), ...source.periods.map(empty)]);

  rows.push(title("Отклонение от медианы"));
  rows.push(header());
  for (let i = 0; i < n; i++) {
    rowFormats[rows.length] = "0.0%";
    const sheetRow = top + rows.length;
    rows.push([label(i), ...source.periods.map((_, j) => {
      const value = source.numbers[i][j];
      const middle = columnNumbers[j].length ? median(columnNumbers[j]) : 0;
      if (value === null || middle === 0) undefinedCells.push(`${columnLetters(left + 1 + j)}${sheetRow}`);
      return {
        formula: `=IF(OR(NOT(ISNUMBER(${col(j)}${firstRow + i})),COUNT(${area(j)})=0),"",IF(MEDIAN(${area(j)})=0,"",${col(j)}${firstRow + i}/MEDIAN(${area(j)})-1))`,
        expected: value === null || middle === 0 ? "" : value / middle - 1
      };
    })]);
  }
  rows.push([empty(), ...source.periods.map(empty)]);

  rows.push(title("Место: 1 — наибольшее значение"));
  rows.push(header());
  for (let i = 0; i < n; i++) {
    rowFormats[rows.length] = "0";
    rows.push([label(i), ...source.periods.map((_, j) => {
      const value = source.numbers[i][j];
      return {
        formula: `=IF(ISNUMBER(${col(j)}${firstRow + i}),RANK(${col(j)}${firstRow + i},${area(j)},0),"")`,
        expected: value === null ? "" : 1 + columnNumbers[j].filter((other) => other > value).length
      };
    })]);
  }
  return { top, left, rows, controlRow: null, undefinedCells, rowFormats };
}

/* --- план ------------------------------------------------------------------------ */

export type TemplateKind = "add_share_growth" | "add_comparison";

interface TemplateSpec {
  build: (source: SourceBlock, top: number, left: number) => TemplateLayout;
  minColumns: number;
  label: string;
}

const SPECS: Record<TemplateKind, TemplateSpec> = {
  add_share_growth: { build: shareGrowthLayout, minColumns: 2, label: "доли и рост" },
  add_comparison: { build: comparisonLayout, minColumns: 1, label: "сравнение" }
};

export interface TemplatePlan {
  readonly kind: TemplateKind;
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly sourceAddress: string;
  readonly sourceSignature: string;
  readonly destAddress: string;
  readonly layout: TemplateLayout;
  readonly items: number;
  /** Столбцы исходника: периоды для долей и роста, показатели для сравнения. */
  readonly periods: number;
  readonly numberFormatsBefore: unknown[][];
  readonly undoAvailable: boolean;
  readonly undoNote: string;
  readonly createdAt: string;
}

export type ShareGrowthPlan = TemplatePlan;
export type ComparisonPlan = TemplatePlan;

async function prepareTemplatePlan(kind: TemplateKind, args: unknown): Promise<TemplatePlan> {
  preflightToolArgs(kind, args);
  const spec = SPECS[kind];
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
    const source = readSourceBlock(range.values as unknown[][], range.rowIndex + 1, range.columnIndex + 1, spec.minColumns);
    let top = range.rowIndex + range.rowCount + 2;
    let left = range.columnIndex + 1;
    if (a.destAddress) {
      const rect = parseA1Rect(checkAddress(a.destAddress));
      if (!rect) throw new ToolError(`destAddress «${a.destAddress}» — не адрес ячейки.`);
      top = rect.rowStart;
      left = rect.columnStart;
    }
    const layout = spec.build(source, top, left);
    const destAddress = `${columnLetters(left)}${top}:${columnLetters(left + source.periods.length)}${top + layout.rows.length - 1}`;
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
      kind,
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

async function executeTemplatePlan(plan: TemplatePlan): Promise<{ where: string; values: unknown[][]; undoRecorded: boolean; sheetName: string }> {
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
      plan.layout.rowFormats[r] && c > 0 ? plan.layout.rowFormats[r] : (plan.numberFormatsBefore[r]?.[c] as string ?? "General")));
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
      undoRecorded = push(action(`${SPECS[plan.kind].label} ${where}`, async () => {
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
    if (mismatches.length) {
      throw new ToolExecutionError(
        `Блок записан на ${where}, но значения расходятся с расчётом панели: ${mismatches.slice(0, 8).join("; ")}${mismatches.length > 8 ? ` и ещё ${mismatches.length - 8}` : ""}. ` +
          `${undoRecorded ? "Его уберёт «Отменить»." : "Проверьте блок."} Готовым он не считается.`,
        "applied"
      );
    }
    return { where, values, undoRecorded, sheetName: sheet.name };
  });
}

const BLOCK_NOTE = "Блок — формулы со ссылками на исходную таблицу: он пересчитается при изменении данных. Значения сверены с расчётом панели.";

export const prepareShareGrowthPlan = (args: unknown) => prepareTemplatePlan("add_share_growth", args);

export async function executeShareGrowthPlan(plan: TemplatePlan) {
  const { where, values, undoRecorded, sheetName } = await executeTemplatePlan(plan);
  const controlRow = plan.layout.controlRow!;
  return {
    ok: true,
    executionState: "verified",
    address: where,
    source: `${sheetName}!${plan.sourceAddress}`,
    items: plan.items,
    periods: plan.periods,
    checkedCells: plan.layout.rows.length * (plan.periods + 1),
    control: { row: `${plan.layout.top + controlRow}`, sumOfShares: (values[controlRow] ?? []).slice(1), note: "Сумма долей каждого периода сверена: 100 % (0 — если итог периода ноль)." },
    ...(plan.layout.undefinedCells.length
      ? { undefinedGrowth: plan.layout.undefinedCells, undefinedGrowthNote: "В этих ячейках рост не определён: в прошлом периоде ноль или пусто. Ячейка оставлена пустой — назови это пользователю." }
      : {}),
    note: BLOCK_NOTE,
    undoable: undoRecorded,
    undoNote: plan.undoNote
  };
}

export const prepareComparisonPlan = (args: unknown) => prepareTemplatePlan("add_comparison", args);

export async function executeComparisonPlan(plan: TemplatePlan) {
  const { where, undoRecorded, sheetName } = await executeTemplatePlan(plan);
  return {
    ok: true,
    executionState: "verified",
    address: where,
    source: `${sheetName}!${plan.sourceAddress}`,
    objects: plan.items,
    metrics: plan.periods,
    checkedCells: plan.layout.rows.length * (plan.periods + 1),
    ...(plan.layout.undefinedCells.length
      ? { undefinedDeviation: plan.layout.undefinedCells, undefinedDeviationNote: "Здесь отклонение не определено: у объекта нет числа по показателю или медиана равна нулю. Ячейка оставлена пустой — назови это пользователю." }
      : {}),
    rankNote: "Место 1 — наибольшее значение. Если для показателя лучше меньшее (затраты, срок), место читается наоборот — скажи это пользователю.",
    note: BLOCK_NOTE,
    undoable: undoRecorded,
    undoNote: plan.undoNote
  };
}
