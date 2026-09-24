/**
 * Чего ждать от диаграммы до того, как её построит Excel.
 *
 * `sheet.charts.add` принимает область и сам решает, где подписи, где ряды
 * и идут ли ряды по строкам или по столбцам. Решает он молча и нередко
 * не так: шапка становится рядом данных, текстовый столбец — рядом нулей,
 * а круговая диаграмма показывает только первый ряд. Проверить по картинке
 * нельзя — диаграмма всегда выглядит как диаграмма. Поэтому ожидание
 * строится здесь по данным, а после операции сверяется с тем, что Excel
 * сообщает о рядах построенной диаграммы.
 *
 * Модуль не обращается к Excel и проверяется без него.
 */

import { columnLetters } from "./formulaFill";

export type ChartKind = "ColumnClustered" | "Line" | "Pie" | "BarClustered" | "XYScatter" | "Area" | "Doughnut";
export type SeriesBy = "columns" | "rows";

export const CHART_KINDS: readonly ChartKind[] = ["ColumnClustered", "Line", "Pie", "BarClustered", "XYScatter", "Area", "Doughnut"];

/** Выше этого ряды уже не читаются глазом, а легенда съедает диаграмму. */
export const MAX_READABLE_SERIES = 12;
/** Столько точек круговой ещё можно различить. */
export const MAX_PIE_SLICES = 12;

export interface ChartExpectation {
  seriesBy: SeriesBy;
  /** Первая строка — подписи рядов (при рядах по столбцам). */
  headerRow: boolean;
  /** Первый столбец — подписи категорий (при рядах по столбцам). */
  labelColumn: boolean;
  /** Имена рядов в том порядке, в каком их должен построить Excel. */
  seriesNames: string[];
  /** Сколько точек в каждом ряду. */
  pointCount: number;
  /** Подписи категорий — первые несколько, для предпросмотра. */
  categories: string[];
  warnings: string[];
}

const isNumber = (value: unknown) => typeof value === "number" && Number.isFinite(value);
const isBlank = (value: unknown) => value === "" || value === null || value === undefined;
const isText = (value: unknown) => typeof value === "string" && value.trim() !== "";

function transpose(values: readonly (readonly unknown[])[]): unknown[][] {
  const width = Math.max(0, ...values.map((row) => row.length));
  return Array.from({ length: width }, (_, column) => values.map((row) => row[column]));
}

/** Строка — подписи, если в ней есть текст, а под ней числа. */
function looksLikeHeader(first: readonly unknown[], rest: readonly (readonly unknown[])[]): boolean {
  if (!rest.length) return false;
  const textCells = first.filter(isText).length;
  if (!textCells) return false;
  // Хотя бы в одном столбце под текстом стоят числа — значит, это шапка.
  return first.some((value, column) => isText(value) && rest.some((row) => isNumber(row[column])));
}

/**
 * Какую диаграмму Excel должен построить из этих данных.
 *
 * Ряды по столбцам — привычный вид таблицы: столбец — ряд, строка — точка.
 * По строкам — то же в транспонированном виде. Всё остальное считается
 * одинаково, поэтому данные при рядах по строкам просто поворачиваются.
 */
export function expectChart(
  values: readonly (readonly unknown[])[],
  kind: ChartKind,
  seriesBy: SeriesBy,
  origin: { rowIndex: number; columnIndex: number }
): ChartExpectation {
  const grid = seriesBy === "columns" ? values.map((row) => [...row]) : transpose(values);
  const warnings: string[] = [];

  const headerRow = grid.length > 1 && looksLikeHeader(grid[0], grid.slice(1));
  const body = headerRow ? grid.slice(1) : grid;
  const firstColumn = body.map((row) => row[0]);
  // Первый столбец — подписи, если в нём текст, а числа есть правее.
  const labelColumn = (grid[0]?.length ?? 0) > 1 && firstColumn.some(isText) && !firstColumn.some(isNumber);

  const seriesColumns = Array.from({ length: grid[0]?.length ?? 0 }, (_, index) => index).filter((index) => !(labelColumn && index === 0));
  const cellName = (column: number) =>
    seriesBy === "columns"
      ? `${columnLetters(origin.columnIndex + column + 1)}`
      : `${origin.rowIndex + column + 1}`;

  const seriesNames: string[] = [];
  const textOnly: string[] = [];
  for (const column of seriesColumns) {
    const cells = body.map((row) => row[column]);
    const name = headerRow && !isBlank(grid[0][column])
      ? String(grid[0][column])
      : seriesBy === "columns" ? `столбец ${cellName(column)}` : `строка ${cellName(column)}`;
    if (!cells.some(isNumber)) {
      // Excel построит такой ряд из нулей — это почти всегда ошибка выбора области.
      if (cells.some(isText)) textOnly.push(name);
    }
    seriesNames.push(name);
  }
  if (textOnly.length) {
    warnings.push(`В рядах ${textOnly.map((name) => `«${name}»`).join(", ")} нет чисел — Excel построит их как нули. Вероятно, область выбрана шире данных.`);
  }

  const categories = labelColumn ? firstColumn.map((value) => (isBlank(value) ? "" : String(value))) : [];
  const pointCount = body.length;

  if (!seriesNames.length || !body.some((row) => seriesColumns.some((column) => isNumber(row[column])))) {
    warnings.unshift("В области нет чисел: строить не из чего.");
  }
  if ((kind === "Pie" || kind === "Doughnut") && seriesNames.length > 1) {
    warnings.push(
      kind === "Pie"
        ? `Круговая диаграмма показывает только первый ряд «${seriesNames[0]}»; остальные (${seriesNames.length - 1}) Excel молча не нарисует.`
        : `Кольцевая рисует каждый ряд отдельным кольцом: ${seriesNames.length} колец читаются плохо.`
    );
  }
  if ((kind === "Pie" || kind === "Doughnut") && body.some((row) => seriesColumns.some((column) => isNumber(row[column]) && (row[column] as number) < 0))) {
    warnings.push("Среди значений есть отрицательные: доля не может быть меньше нуля, Excel нарисует их по модулю.");
  }
  if ((kind === "Pie" || kind === "Doughnut") && pointCount > MAX_PIE_SLICES) {
    warnings.push(`Секторов будет ${pointCount} — больше ${MAX_PIE_SLICES} их не различить. Лучше столбчатая диаграмма.`);
  }
  if (seriesNames.length > MAX_READABLE_SERIES) {
    warnings.push(`Рядов будет ${seriesNames.length}: на диаграмме их не различить. Возможно, ряды должны идти по ${seriesBy === "columns" ? "строкам" : "столбцам"}.`);
  }
  if (kind === "XYScatter" && !labelColumn && seriesNames.length < 2) {
    warnings.push("Точечной диаграмме нужны два числовых столбца: X и Y. Здесь один.");
  }
  // Частая ошибка — ряды не в ту сторону: точек меньше, чем рядов.
  if (seriesNames.length > 1 && pointCount === 1) {
    warnings.push(`Получится ${seriesNames.length} рядов по одной точке. Похоже, ряды должны идти по ${seriesBy === "columns" ? "строкам" : "столбцам"}.`);
  }

  return {
    seriesBy,
    headerRow,
    labelColumn,
    seriesNames,
    pointCount,
    categories: categories.slice(0, 8),
    warnings
  };
}

/**
 * Где положить диаграмму, чтобы она не закрыла данные.
 *
 * Диаграмма лежит над ячейками и закрывает их, ничего не сообщая. Место —
 * через один столбец правее занятой области листа, на уровне начала данных.
 */
export function placementCell(used: { rowIndex: number; columnIndex: number; columnCount: number } | null, sourceRowIndex: number): string {
  const column = used ? used.columnIndex + used.columnCount + 1 : 0;
  return `${columnLetters(column + 1)}${sourceRowIndex + 1}`;
}

/** Положение диаграммы на листе в пунктах, как его отдаёт Excel. */
export interface ChartBox {
  name?: string;
  left: number;
  top: number;
  width: number;
  height: number;
}

/** Зазор между диаграммами, поставленными одна под другую, в пунктах. */
export const CHART_GAP = 12;

export function chartsOverlap(a: ChartBox, b: ChartBox): boolean {
  return a.left < b.left + b.width && b.left < a.left + a.width && a.top < b.top + b.height && b.top < a.top + a.height;
}

/**
 * Верх, на котором диаграмма не заденет ни одну из стоящих на листе.
 *
 * Диаграмма опускается под те, с которыми пересекается, и новое место снова
 * проверяется по всем — прежде сдвиг был один, по первоначальным
 * пересечениям, и третья диаграмма ложилась на вторую (план стабилизации,
 * S6). Движение только вниз, поэтому поиск конечен; предел шагов — страховка.
 * `passed` — под какими диаграммами пришлось опуститься; null — места
 * за предел шагов не нашлось.
 */
export function freeChartTop(box: ChartBox, others: readonly ChartBox[], maxSteps = 200): { top: number; passed: string[] } | null {
  let top = box.top;
  const passed: string[] = [];
  for (let step = 0; step < maxSteps; step++) {
    const hits = others.filter((item) => chartsOverlap({ ...box, top }, item));
    if (!hits.length) return { top, passed };
    for (const item of hits) if (item.name && !passed.includes(item.name)) passed.push(item.name);
    top = Math.max(...hits.map((item) => item.top + item.height)) + CHART_GAP;
  }
  return null;
}

/** Совпадают ли ряды построенной диаграммы с ожиданием. */
export function seriesMismatches(
  expected: ChartExpectation,
  actual: { names: readonly string[]; pointCounts: readonly number[] }
): string[] {
  const problems: string[] = [];
  if (actual.names.length !== expected.seriesNames.length) {
    problems.push(`рядов ${actual.names.length} вместо ${expected.seriesNames.length}`);
  } else if (expected.headerRow) {
    // Без шапки Excel называет ряды сам («Ряд1», «Series1» — зависит от языка),
    // и сверять имена не с чем: сверяется только число рядов и точек.
    actual.names.forEach((name, index) => {
      if (String(name).trim() !== expected.seriesNames[index].trim()) {
        problems.push(`ряд ${index + 1} называется «${name}» вместо «${expected.seriesNames[index]}»`);
      }
    });
  }
  const wrongPoints = actual.pointCounts.filter((count) => count !== expected.pointCount);
  if (wrongPoints.length) problems.push(`точек в ряду ${actual.pointCounts.join("/")} вместо ${expected.pointCount}`);
  return problems;
}
