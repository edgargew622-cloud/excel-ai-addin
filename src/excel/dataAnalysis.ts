/**
 * Анализ данных (этап 10, 10.3): тренд, выбросы, корреляции и сводка по
 * числовым столбцам. Считает панель, а не модель: модель по прочитанным
 * ячейкам «на глаз» ошибается в числах и упирается в предел чтения, а здесь
 * считается весь диапазон, частями. Книга не меняется.
 *
 * Правила — как в Excel, чтобы числа совпадали с формулами пользователя:
 * квартили — КВАРТИЛЬ.ВКЛ (линейная интерполяция), отклонение — СТАНДОТКЛОН.В,
 * корреляция — КОРРЕЛ (Пирсон), тренд — НАКЛОН/ОТРЕЗОК/КВПИРСОН по x.
 */

import { ToolError, checkAddress, rangeOf, MAX_IO_CELLS } from "./excelTools";
import { isDateFormat } from "./dataProfile";
import { columnLetters } from "./formulaFill";
import { captureTarget, officeCapabilities } from "./workbookContext";

/** Больше этого за один анализ не читается: примерно 5 000 строк × 40 столбцов. */
export const MAX_ANALYSIS_CELLS = 200_000;
const MAX_OUTLIER_CELLS = 15;
const MAX_MATRIX_COLUMNS = 12;

/* --- статистика ----------------------------------------------------------------- */

/** КВАРТИЛЬ.ВКЛ / ПРОЦЕНТИЛЬ.ВКЛ: позиция (n−1)·p, линейная интерполяция. */
export function quantileInc(sorted: readonly number[], p: number): number {
  if (!sorted.length) return NaN;
  const position = (sorted.length - 1) * p;
  const low = Math.floor(position);
  const high = Math.ceil(position);
  return sorted[low] + (sorted[high] - sorted[low]) * (position - low);
}

export function mean(values: readonly number[]): number {
  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

/** СТАНДОТКЛОН.В — выборочное. */
export function sampleStd(values: readonly number[]): number {
  if (values.length < 2) return NaN;
  const m = mean(values);
  return Math.sqrt(values.reduce((sum, value) => sum + (value - m) ** 2, 0) / (values.length - 1));
}

/** КОРРЕЛ: Пирсон по парам. */
export function pearson(xs: readonly number[], ys: readonly number[]): number {
  const n = xs.length;
  if (n < 3 || ys.length !== n) return NaN;
  const mx = mean(xs);
  const my = mean(ys);
  let sxy = 0;
  let sxx = 0;
  let syy = 0;
  for (let i = 0; i < n; i++) {
    const dx = xs[i] - mx;
    const dy = ys[i] - my;
    sxy += dx * dy;
    sxx += dx * dx;
    syy += dy * dy;
  }
  return sxx === 0 || syy === 0 ? NaN : sxy / Math.sqrt(sxx * syy);
}

/** НАКЛОН, ОТРЕЗОК и КВПИРСОН. */
export function linearFit(xs: readonly number[], ys: readonly number[]): { slope: number; intercept: number; r2: number } {
  const n = xs.length;
  const mx = mean(xs);
  const my = mean(ys);
  let sxy = 0;
  let sxx = 0;
  for (let i = 0; i < n; i++) {
    sxy += (xs[i] - mx) * (ys[i] - my);
    sxx += (xs[i] - mx) ** 2;
  }
  const slope = sxx === 0 ? NaN : sxy / sxx;
  const r = pearson(xs, ys);
  return { slope, intercept: my - slope * mx, r2: Number.isNaN(r) ? NaN : r * r };
}

/** Число для ответа модели: 6 значащих цифр (сверяется с Excel до сотых), без хвостов вроде 0.30000000000000004. */
export function round(value: number, digits = 6): number | null {
  if (!Number.isFinite(value)) return null;
  if (value === 0) return 0;
  return Number(value.toPrecision(digits));
}

/* --- анализ таблицы ------------------------------------------------------------- */

export interface AnalysisInput {
  values: unknown[][];
  numberFormat?: unknown[][];
  hasHeaders: boolean;
  origin: { rowIndex: number; columnIndex: number };
  /** Только эти столбцы (по имени из шапки или букве); пусто — все числовые. */
  columns?: string[];
  /** Ось тренда: столбец дат или чисел (имя или буква); нет — первый столбец дат, иначе порядок строк. */
  xColumn?: string;
}

interface Column {
  index: number;
  letter: string;
  name: string;
  isDate: boolean;
  numbers: Array<{ row: number; value: number }>;
  nonNumeric: number;
  missing: number;
}

const isNumber = (value: unknown): value is number => typeof value === "number" && Number.isFinite(value);
const isBlank = (value: unknown) => value === null || value === undefined || (typeof value === "string" && value.trim() === "");

function pick(columns: Column[], key: string): Column | undefined {
  const k = key.trim().toLocaleLowerCase("ru");
  return columns.find((column) => column.name.toLocaleLowerCase("ru") === k) ?? columns.find((column) => column.letter.toLowerCase() === k);
}

/** Дата Excel (серийный номер) → ГГГГ-ММ-ДД. */
function serialDate(serial: number): string {
  const ms = Math.round((serial - 25569) * 86400000);
  return new Date(ms).toISOString().slice(0, 10);
}

export function analyzeGrid(input: AnalysisInput) {
  const { values, hasHeaders, origin } = input;
  if (!values.length || !values[0]?.length) throw new ToolError("Диапазон пуст: анализировать нечего.");
  const width = values[0].length;
  const body = hasHeaders ? values.slice(1) : values;
  const firstDataRow = origin.rowIndex + (hasHeaders ? 1 : 0);
  if (body.length < 3) throw new ToolError("Для анализа нужно хотя бы три строки данных.");

  const columns: Column[] = [];
  for (let c = 0; c < width; c++) {
    const letter = columnLetters(origin.columnIndex + c + 1);
    const header = hasHeaders ? values[0][c] : undefined;
    const name = !isBlank(header) ? String(header).trim() : `столбец ${letter}`;
    const numbers: Array<{ row: number; value: number }> = [];
    let nonNumeric = 0;
    let missing = 0;
    let dateFormatted = 0;
    body.forEach((row, r) => {
      const value = row[c];
      if (isBlank(value)) missing++;
      else if (isNumber(value)) {
        numbers.push({ row: r, value });
        if (isDateFormat(input.numberFormat?.[r + (hasHeaders ? 1 : 0)]?.[c])) dateFormatted++;
      } else nonNumeric++;
    });
    columns.push({ index: c, letter, name, isDate: numbers.length > 0 && dateFormatted >= numbers.length * 0.8, numbers, nonNumeric, missing });
  }

  // Числовой столбец: хотя бы 3 числа и числа — большинство заполненных ячеек.
  const numeric = columns.filter((column) => !column.isDate && column.numbers.length >= 3 && column.numbers.length >= (column.numbers.length + column.nonNumeric) * 0.5);
  let chosen = numeric;
  if (input.columns?.length) {
    chosen = input.columns.map((key) => {
      const column = pick(columns, key);
      if (!column) throw new ToolError(`Столбца «${key}» нет: есть ${columns.map((item) => `«${item.name}» (${item.letter})`).join(", ")}.`);
      if (!numeric.includes(column)) throw new ToolError(`В столбце «${column.name}» (${column.letter}) недостаточно чисел для анализа.`);
      return column;
    });
  }
  if (!chosen.length) throw new ToolError("В диапазоне нет числовых столбцов (нужно хотя бы три числа в столбце).");

  // Ось тренда.
  let xColumn: Column | undefined;
  if (input.xColumn) {
    xColumn = pick(columns, input.xColumn);
    if (!xColumn || xColumn.numbers.length < 3) throw new ToolError(`Столбец «${input.xColumn}» не годится для оси тренда: в нём нет дат или чисел.`);
  } else {
    xColumn = columns.find((column) => column.isDate && column.numbers.length >= 3);
  }
  const xOf = new Map<number, number>();
  if (xColumn) for (const item of xColumn.numbers) xOf.set(item.row, item.value);
  const xLabel = xColumn
    ? `${xColumn.isDate ? "даты" : "значения"} столбца ${xColumn.letter} («${xColumn.name}»)`
    : "порядок строк";
  const perUnit = xColumn?.isDate ? "в день" : xColumn ? "на единицу X" : "на строку";

  const address = (column: Column, row: number) => `${column.letter}${firstDataRow + row + 1}`;

  const results = chosen.map((column) => {
    const values = column.numbers.map((item) => item.value);
    const sorted = [...values].sort((a, b) => a - b);
    const q1 = quantileInc(sorted, 0.25);
    const q3 = quantileInc(sorted, 0.75);
    const iqr = q3 - q1;
    const low = q1 - 1.5 * iqr;
    const high = q3 + 1.5 * iqr;
    const outliers = column.numbers.filter((item) => item.value < low || item.value > high)
      .sort((a, b) => Math.abs(b.value - (a.value > high ? high : low)) - Math.abs(a.value - (a.value > high ? high : low)));

    // Тренд — без выбросов: один большой выброс обваливает R², и явный рост
    // выглядел «без тренда». Исключённые точки называются в ответе.
    const outlierRows = new Set(outliers.map((item) => item.row));
    const withX = column.numbers.filter((item) => !xColumn || xOf.has(item.row));
    const clean = withX.filter((item) => !outlierRows.has(item.row));
    const points = clean.length >= 3 ? clean : withX;
    const excluded = withX.length - points.length;
    const xs = points.map((item) => (xColumn ? xOf.get(item.row)! : item.row + 1));
    const ys = points.map((item) => item.value);
    let trend: Record<string, unknown> | null = null;
    if (points.length >= 3) {
      const fit = linearFit(xs, ys);
      const first = fit.intercept + fit.slope * Math.min(...xs);
      const last = fit.intercept + fit.slope * Math.max(...xs);
      const change = last - first;
      const relative = Math.abs(mean(ys)) > 0 ? change / Math.abs(mean(ys)) : NaN;
      const direction = !Number.isFinite(fit.r2) || fit.r2 < 0.3 || Math.abs(relative) < 0.05
        ? "явного тренда нет"
        : change > 0 ? "рост" : "снижение";
      trend = {
        direction,
        slope: round(fit.slope),
        slopeUnit: perUnit,
        ...(xColumn?.isDate ? { slopePerMonth: round(fit.slope * 30.4375) } : {}),
        r2: round(fit.r2, 3),
        fittedFirst: round(first),
        fittedLast: round(last),
        changeOverPeriodPct: Number.isFinite(relative) && Math.abs(first) > 0 ? round((change / Math.abs(first)) * 100, 3) : null,
        points: points.length,
        ...(excluded ? { excludedOutliers: excluded } : {}),
        ...(xColumn?.isDate ? { from: serialDate(Math.min(...xs)), to: serialDate(Math.max(...xs)) } : {})
      };
    }

    return {
      column: column.letter,
      name: column.name,
      count: values.length,
      missing: column.missing,
      ...(column.nonNumeric ? { nonNumeric: column.nonNumeric } : {}),
      mean: round(mean(values)),
      median: round(quantileInc(sorted, 0.5)),
      q1: round(q1),
      q3: round(q3),
      min: round(sorted[0]),
      max: round(sorted[sorted.length - 1]),
      std: round(sampleStd(values)),
      trend,
      outliers: {
        method: "межквартильный размах: ниже Q1 − 1,5·IQR или выше Q3 + 1,5·IQR",
        low: round(low),
        high: round(high),
        count: outliers.length,
        cells: outliers.slice(0, MAX_OUTLIER_CELLS).map((item) => ({ address: address(column, item.row), value: item.value })),
        ...(outliers.length > MAX_OUTLIER_CELLS ? { more: outliers.length - MAX_OUTLIER_CELLS } : {})
      }
    };
  });

  // Корреляции по строкам, где заполнены оба столбца. Выброс сильно двигает
  // Пирсона: если без выбросов связь заметно другая, называются обе цифры.
  const strengthOf = (r: number) => {
    const abs = Math.abs(r);
    if (!Number.isFinite(r)) return "не считается";
    const word = abs >= 0.7 ? "сильная" : abs >= 0.4 ? "умеренная" : abs >= 0.2 ? "слабая" : "нет связи";
    return `${word}${abs >= 0.2 ? (r > 0 ? ", прямая" : ", обратная") : ""}`;
  };
  const outlierRowsOf = new Map(results.map((result, index) => [chosen[index], new Set(result.outliers.cells.length === result.outliers.count
    ? result.outliers.cells.map((cell) => Number(cell.address.replace(/^[A-Z]+/, "")) - firstDataRow - 1)
    : [])]));
  const pairs: Array<{ a: string; b: string; r: number | null; n: number; strength: string; rWithoutOutliers?: number | null; strengthWithoutOutliers?: string }> = [];
  for (let i = 0; i < chosen.length; i++) {
    for (let j = i + 1; j < chosen.length; j++) {
      const bj = new Map(chosen[j].numbers.map((item) => [item.row, item.value]));
      const both = chosen[i].numbers.filter((item) => bj.has(item.row));
      const r = pearson(both.map((item) => item.value), both.map((item) => bj.get(item.row)!));
      const skip = new Set([...(outlierRowsOf.get(chosen[i]) ?? []), ...(outlierRowsOf.get(chosen[j]) ?? [])]);
      const clean = both.filter((item) => !skip.has(item.row));
      const rClean = skip.size && clean.length >= 3 ? pearson(clean.map((item) => item.value), clean.map((item) => bj.get(item.row)!)) : NaN;
      const differs = Number.isFinite(rClean) && Math.abs(rClean - r) >= 0.1;
      pairs.push({
        a: chosen[i].name,
        b: chosen[j].name,
        r: round(r, 3),
        n: both.length,
        strength: strengthOf(r),
        ...(differs ? { rWithoutOutliers: round(rClean, 3), strengthWithoutOutliers: strengthOf(rClean) } : {})
      });
    }
  }
  pairs.sort((x, y) => Math.abs(y.r ?? 0) - Math.abs(x.r ?? 0));

  const lastRow = firstDataRow + body.length;
  return {
    analyzedRows: `${firstDataRow + 1}:${lastRow}`,
    rows: body.length,
    xAxis: xLabel,
    columns: results,
    correlations: chosen.length > 1
      ? {
        method: "Пирсон (как КОРРЕЛ), по строкам, где заполнены оба столбца",
        pairs: chosen.length <= MAX_MATRIX_COLUMNS ? pairs : pairs.slice(0, 30),
        ...(chosen.length > MAX_MATRIX_COLUMNS ? { note: `Показаны 30 самых сильных пар из ${pairs.length}.` } : {})
      }
      : null,
    skipped: columns.filter((column) => !chosen.includes(column) && column !== xColumn).map((column) => ({
      column: column.letter,
      name: column.name,
      reason: numeric.includes(column)
        ? "не выбран в columns"
        : column.isDate ? "даты" : column.nonNumeric >= column.numbers.length ? "в основном текст" : "мало чисел"
    })),
    notes: [
      "Посчитано панелью по всем строкам диапазона; формулы Excel на тех же числах дадут то же: КВАРТИЛЬ.ВКЛ, СТАНДОТКЛОН.В, КОРРЕЛ, НАКЛОН, КВПИРСОН.",
      "Корреляция — это совместное изменение, а не причина: сильная связь не доказывает, что одно вызывает другое.",
      "Тренд — прямая линия по точкам без выбросов (excludedOutliers — сколько исключено); «явного тренда нет», если R² меньше 0,3 или изменение меньше 5% от среднего."
    ]
  };
}

/* --- инструмент ----------------------------------------------------------------- */

export async function analyzeRange(args: { sheet?: string; address?: string; hasHeaders?: boolean; columns?: string[]; xColumn?: string }) {
  const target = await captureTarget(args.sheet);
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    sheet.load("name");
    let area: Excel.Range;
    if (args.address?.trim()) {
      area = await rangeOf(ctx, sheet, checkAddress(args.address));
    } else {
      const used = officeCapabilities().usedRangeOrNull ? sheet.getUsedRangeOrNullObject(true) : sheet.getUsedRange(true);
      used.load("isNullObject");
      await ctx.sync();
      if ((used as any).isNullObject) return { sheet: sheet.name, empty: true, note: "Лист пуст: анализировать нечего." };
      area = used;
    }
    area.load(["address", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
    await ctx.sync();
    if (area.columnCount > 200) throw new ToolError(`В ${area.address} ${area.columnCount} столбцов — для анализа выберите нужные столбцы области.`);
    const readableRows = Math.min(area.rowCount, Math.floor(MAX_ANALYSIS_CELLS / area.columnCount));
    const rowsPerChunk = Math.max(1, Math.floor(MAX_IO_CELLS / area.columnCount));
    const values: unknown[][] = [];
    const numberFormat: unknown[][] = [];
    for (let start = 0; start < readableRows; start += rowsPerChunk) {
      const rows = Math.min(rowsPerChunk, readableRows - start);
      const chunk = sheet.getRangeByIndexes(area.rowIndex + start, area.columnIndex, rows, area.columnCount);
      chunk.load(["values", "numberFormat"]);
      await ctx.sync();
      values.push(...(chunk.values as unknown[][]));
      numberFormat.push(...(chunk.numberFormat as unknown[][]));
    }
    const incomplete = readableRows < area.rowCount;
    const analysis = analyzeGrid({
      values,
      numberFormat,
      hasHeaders: args.hasHeaders !== false,
      origin: { rowIndex: area.rowIndex, columnIndex: area.columnIndex },
      ...(args.columns?.length ? { columns: args.columns } : {}),
      ...(args.xColumn ? { xColumn: args.xColumn } : {})
    });
    return {
      sheet: sheet.name,
      address: area.address.slice(area.address.lastIndexOf("!") + 1),
      incomplete,
      ...(incomplete ? { incompleteNote: `Прочитаны первые ${readableRows} строк из ${area.rowCount}: предел анализа ${MAX_ANALYSIS_CELLS} ячеек. Для остального укажите address с нужными строками.` } : {}),
      ...analysis
    };
  });
}
