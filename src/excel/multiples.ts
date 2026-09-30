/**
 * Сравнение компаний по мультипликаторам (этап 8, 8.4.1).
 *
 * Блок формул под таблицей компаний: EV (стоимость компании), EV/EBITDA,
 * EV/Выручка, P/E по каждой компании и медиана, среднее, 1-й и 3-й квартили
 * (КВАРТИЛЬ.ВКЛ, этап 10), минимум, максимум по группе. Как в шаблонах 7.5: пишутся формулы со ссылками на ячейки
 * таблицы, а не числа модели; панель сама считает каждое значение и после
 * записи сверяет с ним каждую ячейку.
 *
 * EV = капитализация + чистый долг, если нет готового столбца EV.
 * Мультипликатор при нулевом или отрицательном знаменателе (убыток, отрицательная
 * EBITDA) не имеет смысла: ячейка остаётся пустой и называется, а в медиану
 * группы такие компании не входят — как принято в сравнительном анализе.
 */

import { parseA1Rect } from "./a1";
import {
  assertTargetWritable,
  checkAddress,
  deepFreeze,
  preflightToolArgs,
  rangeOf,
  ToolError
} from "./excelTools";
import { columnLetters } from "./formulaFill";
import { executeTemplatePlan, median, type TemplateCell, type TemplateLayout, type TemplatePlan } from "./templates";
import { quantileInc } from "./dataAnalysis";
import { isCustomUndoAvailable } from "./undo";
import { captureTarget } from "./workbookContext";

const MAX_COMPANIES = 200;
/** «12,3x»: буква в кавычках, чтобы Excel не принял её за код формата. */
const MULTIPLE_FORMAT = '0.0"x"';

export type MultipleKind = "evEbitda" | "evRevenue" | "pe";

export const MULTIPLE_TEXT: Record<MultipleKind, string> = { evEbitda: "EV/EBITDA", evRevenue: "EV/Выручка", pe: "P/E" };

/** Какие столбцы таблицы что значат — по заголовкам, как их назвал пользователь. */
export interface MultiplesColumns {
  marketCap?: string;
  netDebt?: string;
  enterpriseValue?: string;
  ebitda?: string;
  revenue?: string;
  netIncome?: string;
}

interface Resolved {
  /** Номер столбца листа, с 1. */
  marketCap?: number;
  netDebt?: number;
  enterpriseValue?: number;
  ebitda?: number;
  revenue?: number;
  netIncome?: number;
}

const num = (value: unknown): number | null => (typeof value === "number" && Number.isFinite(value) ? value : null);

/**
 * Раскладка блока. source — значения таблицы вместе с шапкой; top/left —
 * где она на листе (с 1); blockTop/blockLeft — где встанет блок.
 */
export function multiplesLayout(
  values: readonly (readonly unknown[])[],
  top: number,
  left: number,
  columns: Resolved,
  blockTop: number,
  blockLeft: number
): { layout: TemplateLayout; multiples: MultipleKind[]; skipped: string[] } {
  const companies = values.slice(1).map((row) => String(row[0] ?? "").trim());
  const n = companies.length;
  const firstRow = top + 1;
  const cellOf = (column: number, i: number) => `${columnLetters(column)}${firstRow + i}`;
  const valueOf = (column: number | undefined, i: number) => (column === undefined ? null : num(values[1 + i]?.[column - left]));

  const multiples: MultipleKind[] = [
    ...(columns.ebitda !== undefined ? ["evEbitda" as const] : []),
    ...(columns.revenue !== undefined ? ["evRevenue" as const] : []),
    ...(columns.netIncome !== undefined && columns.marketCap !== undefined ? ["pe" as const] : [])
  ];
  const evColumn = blockLeft + 1;
  const multipleColumn = (index: number) => blockLeft + 2 + index;
  const blockRow = (i: number) => blockTop + 2 + i;
  const evCell = (i: number) => `${columnLetters(evColumn)}${blockRow(i)}`;

  const rows: TemplateCell[][] = [];
  const rowFormats: Record<number, string> = {};
  const undefinedCells: string[] = [];
  const skipped: string[] = [];
  const text = (value: string): TemplateCell => ({ formula: value, expected: value });
  const width = 2 + multiples.length;
  const empty = (): TemplateCell => ({ formula: "", expected: "" });

  rows.push([text(columns.enterpriseValue !== undefined ? "Мультипликаторы (EV — из таблицы)" : "Мультипликаторы (EV = капитализация + чистый долг)"), ...Array.from({ length: width - 1 }, empty)]);
  rows.push([text("Компания"), text("EV"), ...multiples.map((kind) => text(MULTIPLE_TEXT[kind]))]);

  const evs: (number | null)[] = [];
  const results: (number | null)[][] = [];
  for (let i = 0; i < n; i++) {
    let evFormula: string;
    let ev: number | null;
    if (columns.enterpriseValue !== undefined) {
      evFormula = `=IF(ISNUMBER(${cellOf(columns.enterpriseValue, i)}),${cellOf(columns.enterpriseValue, i)},"")`;
      ev = valueOf(columns.enterpriseValue, i);
    } else {
      const cap = cellOf(columns.marketCap!, i);
      const debt = cellOf(columns.netDebt!, i);
      evFormula = `=IF(AND(ISNUMBER(${cap}),ISNUMBER(${debt})),${cap}+${debt},"")`;
      const capValue = valueOf(columns.marketCap, i);
      const debtValue = valueOf(columns.netDebt, i);
      ev = capValue === null || debtValue === null ? null : capValue + debtValue;
    }
    evs.push(ev);
    if (ev === null) skipped.push(`${companies[i]}: EV не посчитать — нет чисел`);
    const row: TemplateCell[] = [{ formula: `=$${columnLetters(left)}${firstRow + i}`, expected: companies[i] }, { formula: evFormula, expected: ev ?? "" }];
    const rowResults: (number | null)[] = [];
    multiples.forEach((kind, index) => {
      const numeratorCell = kind === "pe" ? cellOf(columns.marketCap!, i) : evCell(i);
      const numerator = kind === "pe" ? valueOf(columns.marketCap, i) : ev;
      const column = kind === "evEbitda" ? columns.ebitda! : kind === "evRevenue" ? columns.revenue! : columns.netIncome!;
      const denominatorCell = cellOf(column, i);
      const denominator = valueOf(column, i);
      const defined = numerator !== null && denominator !== null && denominator > 0;
      const result = defined ? numerator! / denominator! : null;
      if (!defined) undefinedCells.push(`${columnLetters(multipleColumn(index))}${blockRow(i)}`);
      rowResults.push(result);
      row.push({
        formula: `=IF(AND(ISNUMBER(${numeratorCell}),ISNUMBER(${denominatorCell})),IF(${denominatorCell}>0,${numeratorCell}/${denominatorCell},""),"")`,
        expected: result ?? ""
      });
    });
    results.push(rowResults);
    rowFormats[rows.length] = MULTIPLE_FORMAT;
    rows.push(row);
  }
  rows.push(Array.from({ length: width }, empty));

  const area = (column: number) => `${columnLetters(column)}$${blockRow(0)}:${columnLetters(column)}$${blockRow(n - 1)}`;
  const stat = (name: string, fn: string, compute: (numbers: number[]) => number, extra = "") => {
    const numbersIn = (list: (number | null)[]) => list.filter((value): value is number => value !== null);
    rowFormats[rows.length] = MULTIPLE_FORMAT;
    rows.push([
      text(name),
      { formula: `=IF(COUNT(${area(evColumn)})=0,"",${fn}(${area(evColumn)}${extra}))`, expected: numbersIn(evs).length ? compute(numbersIn(evs)) : "" },
      ...multiples.map((_, index) => {
        const list = numbersIn(results.map((row) => row[index]));
        return { formula: `=IF(COUNT(${area(multipleColumn(index))})=0,"",${fn}(${area(multipleColumn(index))}${extra}))`, expected: list.length ? compute(list) : "" };
      })
    ]);
  };
  stat("Медиана", "MEDIAN", median);
  stat("Среднее", "AVERAGE", (list) => list.reduce((sum, value) => sum + value, 0) / list.length);
  // Квартили (этап 10, 10.4) — КВАРТИЛЬ.ВКЛ, как в Excel; пустые ячейки
  // (мультипликатор без смысла) в них не входят, как и в медиану.
  const sortedQuartile = (p: number) => (list: number[]) => quantileInc([...list].sort((x, y) => x - y), p);
  stat("1-й квартиль", "QUARTILE.INC", sortedQuartile(0.25), ",1");
  stat("3-й квартиль", "QUARTILE.INC", sortedQuartile(0.75), ",3");
  stat("Минимум", "MIN", (list) => Math.min(...list));
  stat("Максимум", "MAX", (list) => Math.max(...list));

  // EV — деньги, а не «x»: столбец EV в строках компаний и статистики — «# ##0».
  return { layout: { top: blockTop, left: blockLeft, rows, controlRow: null, undefinedCells, rowFormats, columnFormats: { 1: "#,##0" } } as TemplateLayout, multiples, skipped };
}

export async function prepareMultiplesPlan(args: unknown): Promise<TemplatePlan> {
  preflightToolArgs("add_multiples", args);
  const a = args as { sheet?: string; address: string; destAddress?: string; columns: MultiplesColumns };
  const address = checkAddress(a.address);
  const target = await captureTarget(a.sheet);
  const prepared = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    const range = await rangeOf(ctx, sheet, address);
    range.load(["address", "rowIndex", "columnIndex", "rowCount", "columnCount", "values", "formulas"]);
    sheet.load(["id", "name"]);
    try { sheet.protection?.load(["protected", "options"]); } catch { /* среда без сведений о защите */ }
    await ctx.sync();
    const values = range.values as unknown[][];
    if (values.length < 2) throw new ToolError("Нужна таблица: первая строка — заголовки столбцов, первый столбец — компании, хотя бы одна компания.");
    if (values.length - 1 > MAX_COMPANIES) throw new ToolError(`Компаний больше ${MAX_COMPANIES}: разбейте таблицу.`);
    const top = range.rowIndex + 1;
    const left = range.columnIndex + 1;
    const headers = values[0].map((value) => String(value ?? "").trim());
    const resolve = (name: string | undefined, what: string): number | undefined => {
      if (!name?.trim()) return undefined;
      const index = headers.findIndex((header) => header.toLowerCase() === name.trim().toLowerCase());
      if (index <= 0) throw new ToolError(`${what}: столбца «${name}» в шапке нет. Заголовки: ${headers.slice(1).map((header) => `«${header}»`).join(", ")}.`);
      return left + index;
    };
    const c = a.columns ?? {};
    const columns: Resolved = {
      marketCap: resolve(c.marketCap, "Капитализация"),
      netDebt: resolve(c.netDebt, "Чистый долг"),
      enterpriseValue: resolve(c.enterpriseValue, "EV"),
      ebitda: resolve(c.ebitda, "EBITDA"),
      revenue: resolve(c.revenue, "Выручка"),
      netIncome: resolve(c.netIncome, "Чистая прибыль")
    };
    if (columns.enterpriseValue === undefined && (columns.marketCap === undefined || columns.netDebt === undefined)) {
      throw new ToolError("Для EV нужен столбец enterpriseValue или оба столбца: marketCap (капитализация) и netDebt (чистый долг; денежные средства — со знаком минус).");
    }
    if (columns.ebitda === undefined && columns.revenue === undefined && !(columns.netIncome !== undefined && columns.marketCap !== undefined)) {
      throw new ToolError("Нечего считать: нужен хотя бы один из столбцов ebitda, revenue или netIncome (для P/E ещё и marketCap).");
    }
    if (values.slice(1).some((row) => !String(row[0] ?? "").trim())) throw new ToolError("У каждой строки нужна компания в первом столбце таблицы.");
    const text = values.slice(1).flatMap((row, i) =>
      Object.values(columns).filter((column): column is number => column !== undefined)
        .filter((column) => { const value = row[column - left]; return value !== "" && value !== null && typeof value !== "number"; })
        .map((column) => `${columnLetters(column)}${top + 1 + i}`));
    if (text.length) throw new ToolError(`Не числа в ячейках ${text.slice(0, 6).join(", ")}${text.length > 6 ? " и других" : ""}: сначала приведите данные (convert_values).`);

    let blockTop = range.rowIndex + range.rowCount + 2;
    let blockLeft = left;
    if (a.destAddress) {
      const rect = parseA1Rect(checkAddress(a.destAddress));
      if (!rect) throw new ToolError(`destAddress «${a.destAddress}» — не адрес ячейки.`);
      blockTop = rect.rowStart;
      blockLeft = rect.columnStart;
    }
    const { layout, multiples, skipped } = multiplesLayout(values, top, left, columns, blockTop, blockLeft);
    const width = 2 + multiples.length;
    const destAddress = `${columnLetters(blockLeft)}${blockTop}:${columnLetters(blockLeft + width - 1)}${blockTop + layout.rows.length - 1}`;
    const dest = sheet.getRange(destAddress);
    dest.load(["formulas", "numberFormat"]);
    try { dest.format?.protection?.load("locked"); } catch { /* нет сведений */ }
    const used = sheet.getUsedRangeOrNullObject(true);
    used.load(["isNullObject", "rowIndex", "rowCount"]);
    await ctx.sync();
    assertTargetWritable(sheet, dest);
    if ((dest.formulas as unknown[][]).some((row) => row.some((value) => value !== "" && value !== null))) {
      const free = used.isNullObject ? blockTop : used.rowIndex + used.rowCount + 2;
      throw new ToolError(`Место под блок ${destAddress} занято. Свободно ниже данных листа: начиная с ${columnLetters(blockLeft)}${free} — передайте его в destAddress.`);
    }
    const undo = isCustomUndoAvailable();
    return {
      kind: "add_multiples" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      target: { ...target, sheetName: sheet.name },
      sourceAddress: String(range.address).replace(/^.*!/, ""),
      sourceSignature: JSON.stringify(range.formulas),
      destAddress,
      layout,
      items: values.length - 1,
      periods: width - 1,
      numberFormatsBefore: dest.numberFormat as unknown[][],
      multiples: multiples.map((kind) => MULTIPLE_TEXT[kind]),
      skipped,
      undoAvailable: undo,
      undoNote: undo ? "Отмена очистит блок и вернёт прежний формат чисел." : "Отмена недоступна: монитор изменений Excel не активен.",
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared) as TemplatePlan;
}

/** Названия строк статистики блока — по ним, а не по позиции, ищутся итоги. */
export const GROUP_STAT_ROWS = ["Медиана", "Среднее", "1-й квартиль", "3-й квартиль", "Минимум", "Максимум"] as const;

/**
 * Итоги группы из сверенных значений блока. Раньше медиана бралась «четвёртой
 * строкой с конца»; с квартилями (10.4) там оказался 1-й квартиль, и модель
 * назвала его медианой — живая проверка 30.09.2026. Теперь — по названию строки.
 */
export function groupStats(layout: TemplateLayout, values: readonly (readonly unknown[])[], multiples: readonly string[]) {
  const out: Record<string, Record<string, unknown>> = {};
  for (const label of GROUP_STAT_ROWS) {
    const index = layout.rows.findIndex((row) => row[0]?.expected === label);
    if (index < 0) continue;
    const row = values[index] ?? [];
    out[label] = { EV: row[1], ...Object.fromEntries(multiples.map((name, column) => [name, row[2 + column]])) };
  }
  return out;
}

export async function executeMultiplesPlan(plan: TemplatePlan) {
  const { where, values, undoRecorded, sheetName } = await executeTemplatePlan(plan);
  const multiples = (plan as any).multiples as string[];
  const stats = groupStats(plan.layout, values, multiples);
  return {
    ok: true,
    executionState: "verified",
    address: where,
    source: `${sheetName}!${plan.sourceAddress}`,
    companies: plan.items,
    multiples,
    medians: Object.fromEntries(multiples.map((name) => [name, stats["Медиана"]?.[name]])),
    // Все итоги группы, сверенные с Excel: называй их отсюда, не перечитывая блок.
    groupStats: stats,
    checkedCells: plan.layout.rows.length * (plan.periods + 1),
    ...(plan.layout.undefinedCells.length
      ? { notMeaningful: plan.layout.undefinedCells, notMeaningfulNote: "Здесь мультипликатор не имеет смысла: знаменатель ноль или отрицательный (убыток, отрицательная EBITDA) или нет числа. Ячейка пустая, в медиану не входит — назови это пользователю." }
      : {}),
    ...((plan as any).skipped?.length ? { skipped: (plan as any).skipped } : {}),
    note: "Блок — формулы со ссылками на таблицу компаний: пересчитается при изменении данных. Значения сверены с расчётом панели. Чистый долг — долг минус денежные средства; единицы (млн, млрд) должны совпадать во всех столбцах.",
    undoable: undoRecorded,
    undoNote: plan.undoNote
  };
}
