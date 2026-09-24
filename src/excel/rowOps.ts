/**
 * Разбор последствий вставки и удаления строк.
 *
 * Это самые опасные операции этапа 6: они меняют адреса, у них нет отката,
 * а Excel выполняет их молча. Формула `=СУММ(A2:A10)` на другом листе после
 * удаления строк превращается в `#ССЫЛКА!` или тихо считает не то, и целевые
 * ячейки при этом выглядят безупречно. Поэтому последствия нужно назвать
 * до операции — здесь, на чистых данных, и сверить после неё.
 *
 * Разбор намеренно приблизительный в одну сторону: он не притворяется, что
 * знает про формулу всё. Ссылки в кавычках не трогаются, структурированные
 * ссылки таблиц (`Таблица[Столбец]`) не разбираются вовсе — про такие формулы
 * честнее сказать «не разобрана», чем уверенно промолчать.
 */

export interface RowBand {
  /** Первая строка полосы, нумерация с 1, как в интерфейсе Excel. */
  readonly startRow: number;
  /** Последняя строка полосы включительно. */
  readonly endRow: number;
}

export interface FormulaReference {
  /** Лист ссылки; null — тот же лист, где стоит формула. */
  readonly sheet: string | null;
  readonly rowStart: number;
  readonly rowEnd: number;
  /** Текст ссылки, как он записан в формуле. */
  readonly text: string;
  /** Строки закреплены знаком доллара с обеих сторон. */
  readonly rowsPinned: boolean;
  /** Столбцы ссылки, нумерация с 1: A — 1. */
  readonly columnStart: number;
  readonly columnEnd: number;
  /** Целые столбцы (`C:C`): строки ссылки — весь лист, операции со строками её не задевают. */
  readonly wholeColumns?: true;
  /** Целые строки (`2:2`): столбцы — весь лист, операции со столбцами её не задевают. */
  readonly wholeRows?: true;
}

/** Размеры листа Excel: целые столбцы и строки ссылаются на всю длину. */
export const SHEET_ROWS = 1_048_576;
export const SHEET_COLUMNS = 16_384;

export function rowBand(startRow: number, count: number): RowBand {
  return { startRow, endRow: startRow + count - 1 };
}

const SHEET = "(?:'((?:[^']|'')+)'|([A-Za-z_\u0400-\u04FF][A-Za-z0-9_.\u0400-\u04FF ]*))!";
const CELL = String.raw`(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})`;
const REFERENCE = new RegExp(`(?:${SHEET})?${CELL}(?::${CELL})?`, "g");
// Замер 24 сентября 2026 года: `SUM(C:C)` при удалении столбца C становится
// `#ССЫЛКА!`, `SUM(B:D)` сужается до `B:C`, `SUM(2:2)` — то же со строками.
const COLUMNS = new RegExp(`(?:${SHEET})?(\\$?)([A-Za-z]{1,3}):(\\$?)([A-Za-z]{1,3})`, "g");
const ROWS = new RegExp(`(?:${SHEET})?(\\$?)(\\d{1,7}):(\\$?)(\\d{1,7})`, "g");

/** Имя листа без кавычек и удвоенных апострофов, для сравнения. */
function sheetName(quoted: string | undefined, plain: string | undefined): string | null {
  if (quoted !== undefined) return quoted.replace(/''/g, "'");
  return plain ?? null;
}

function sameSheet(a: string | null, b: string): boolean {
  return a === null ? true : a.trim().toLowerCase() === b.trim().toLowerCase();
}

/**
 * Ссылки на ячейки внутри формулы.
 *
 * Текст в кавычках пропускается: в нём может стоять что угодно похожее.
 * Совпадение, за которым сразу идёт открывающая скобка, — имя функции
 * (`LOG10(`), а не ссылка.
 */
export function formulaReferences(formula: unknown): FormulaReference[] {
  if (typeof formula !== "string" || !formula.startsWith("=")) return [];
  const found: FormulaReference[] = [];
  let index = 0;
  let inQuotes = false;

  while (index < formula.length) {
    const char = formula[index];
    if (char === '"') { inQuotes = !inQuotes; index += 1; continue; }
    if (inQuotes) { index += 1; continue; }

    REFERENCE.lastIndex = index;
    const match = REFERENCE.exec(formula);
    if (!match || match.index !== index) {
      const whole = wholeReference(formula, index);
      if (whole) { found.push(whole.reference); index += whole.length; continue; }
      index += 1;
      continue;
    }

    const whole = match[0];
    const next = formula[index + whole.length] ?? "";
    const previous = index > 0 ? formula[index - 1] : "";
    // `LOG10(` — имя функции; буква или цифра перед совпадением — часть имени.
    if (next === "(" || /[A-Za-z0-9_.]/.test(previous)) { index += whole.length; continue; }

    const sheet = sheetName(match[1], match[2]);
    const firstRowPinned = match[5] === "$";
    const firstRow = Number(match[6]);
    const secondRow = match[10] === undefined ? firstRow : Number(match[10]);
    const secondRowPinned = match[9] === undefined ? firstRowPinned : match[9] === "$";
    const letters = (text: string) => [...text.toUpperCase()].reduce((total, ch) => total * 26 + ch.charCodeAt(0) - 64, 0);
    const firstColumn = letters(match[4]);
    const secondColumn = match[8] === undefined ? firstColumn : letters(match[8]);
    found.push({
      columnStart: Math.min(firstColumn, secondColumn),
      columnEnd: Math.max(firstColumn, secondColumn),
      sheet,
      rowStart: Math.min(firstRow, secondRow),
      rowEnd: Math.max(firstRow, secondRow),
      text: whole,
      rowsPinned: firstRowPinned && secondRowPinned
    });
    index += whole.length;
  }
  return found;
}

const letterNumber = (text: string) => [...text.toUpperCase()].reduce((total, ch) => total * 26 + ch.charCodeAt(0) - 64, 0);

/** Ссылка на целые столбцы или строки с этой позиции — или null. */
function wholeReference(formula: string, index: number): { reference: FormulaReference; length: number } | null {
  const previous = index > 0 ? formula[index - 1] : "";
  if (/[A-Za-z0-9_.]/.test(previous)) return null;
  for (const [pattern, kind] of [[COLUMNS, "columns"], [ROWS, "rows"]] as const) {
    pattern.lastIndex = index;
    const match = pattern.exec(formula);
    if (!match || match.index !== index) continue;
    const next = formula[index + match[0].length] ?? "";
    if (/[A-Za-z0-9_(!$]/.test(next)) continue;
    const sheet = sheetName(match[1], match[2]);
    if (kind === "columns") {
      const [a, b] = [letterNumber(match[4]), letterNumber(match[6])];
      return {
        length: match[0].length,
        reference: { sheet, rowStart: 1, rowEnd: SHEET_ROWS, columnStart: Math.min(a, b), columnEnd: Math.max(a, b), text: match[0], rowsPinned: true, wholeColumns: true }
      };
    }
    const [a, b] = [Number(match[4]), Number(match[6])];
    return {
      length: match[0].length,
      reference: { sheet, rowStart: Math.min(a, b), rowEnd: Math.max(a, b), columnStart: 1, columnEnd: SHEET_COLUMNS, text: match[0], rowsPinned: match[3] === "$" && match[5] === "$", wholeRows: true }
    };
  }
  return null;
}

/** Формулы со структурированными ссылками таблиц разобрать нельзя. */
export function usesTableReference(formula: unknown): boolean {
  return typeof formula === "string" && /[A-Za-z_Ѐ-ӿ][\w.Ѐ-ӿ]*\[/.test(formula);
}

/**
 * Ссылки, которые после удаления дубликатов увидят другие данные.
 *
 * Замер 24 сентября 2026 года: `removeDuplicates` сдвигает значения вверх
 * внутри области и не подстраивает ссылки — `=C8` остаётся `=C8`, а в C8
 * уже данные другой строки, без всякой ошибки. Задеты ссылки на ячейки
 * области начиная с первой удаляемой строки. Ссылка, охватывающая все
 * строки данных области по высоте, — это итог по таблице: он изменится
 * ровно на удалённые строки, как и задумано, и в список не входит.
 */
export function shiftedReferences(
  formula: unknown,
  formulaSheet: string,
  targetSheet: string,
  area: { rowStart: number; rowEnd: number; columnStart: number; columnEnd: number; dataRowStart: number },
  firstRemovedRow: number
): FormulaReference[] {
  const hit: FormulaReference[] = [];
  for (const reference of formulaReferences(formula)) {
    const onTarget = reference.sheet === null
      ? sameSheet(formulaSheet, targetSheet)
      : sameSheet(reference.sheet, targetSheet);
    if (!onTarget) continue;
    const columns = reference.columnStart <= area.columnEnd && reference.columnEnd >= area.columnStart;
    const rows = reference.rowEnd >= firstRemovedRow && reference.rowStart <= area.rowEnd;
    if (!columns || !rows) continue;
    const wholeHeight = reference.rowStart <= area.dataRowStart && reference.rowEnd >= area.rowEnd;
    if (!wholeHeight) hit.push(reference);
  }
  return hit;
}

export interface DeleteImpact {
  /** Ссылка целиком лежит в удаляемых строках: станет `#ССЫЛКА!`. */
  readonly broken: FormulaReference[];
  /** Ссылка задета частично: диапазон уменьшится, итог изменится молча. */
  readonly shrunk: FormulaReference[];
}

/**
 * Что станет со ссылками формулы после удаления строк.
 *
 * `formulaSheet` — лист, на котором стоит сама формула: ссылка без имени листа
 * указывает именно туда. Закреплённые долларом строки от удаления не спасают:
 * Excel всё равно ломает ссылку на исчезнувшую ячейку.
 */
/** Ось операции: строки или столбцы. Полоса столбцов задаётся теми же числами — номерами столбцов. */
export type Axis = "rows" | "columns";

/**
 * Протяжённость ссылки вдоль оси. Целые столбцы операции со строками
 * не задевают (ссылка по-прежнему на весь столбец), целые строки —
 * операции со столбцами.
 */
function extent(reference: FormulaReference, axis: Axis): { start: number; end: number } | null {
  if (axis === "rows") return reference.wholeColumns ? null : { start: reference.rowStart, end: reference.rowEnd };
  return reference.wholeRows ? null : { start: reference.columnStart, end: reference.columnEnd };
}

export function deleteImpact(
  formula: unknown,
  formulaSheet: string,
  targetSheet: string,
  band: RowBand,
  axis: Axis = "rows"
): DeleteImpact {
  const broken: FormulaReference[] = [];
  const shrunk: FormulaReference[] = [];
  for (const reference of formulaReferences(formula)) {
    const onTarget = reference.sheet === null
      ? sameSheet(null, formulaSheet) && sameSheet(formulaSheet, targetSheet)
      : sameSheet(reference.sheet, targetSheet);
    if (!onTarget) continue;
    const span = extent(reference, axis);
    if (!span) continue;
    if (span.start >= band.startRow && span.end <= band.endRow) broken.push(reference);
    else if (span.start <= band.endRow && span.end >= band.startRow) shrunk.push(reference);
  }
  return { broken, shrunk };
}

/**
 * Ссылки, которые после вставки не охватят новые строки.
 *
 * Excel расширяет диапазон, когда строки вставлены внутрь него, но не когда
 * они вставлены вплотную сверху или снизу. Так `=СУММ(A2:A10)` после вставки
 * строки 11 продолжает считать по старой границе, и новая строка молча
 * не попадает в итог.
 */
export function insertBlindSpots(
  formula: unknown,
  formulaSheet: string,
  targetSheet: string,
  band: RowBand,
  axis: Axis = "rows"
): FormulaReference[] {
  const missed: FormulaReference[] = [];
  for (const reference of formulaReferences(formula)) {
    const onTarget = reference.sheet === null
      ? sameSheet(formulaSheet, targetSheet)
      : sameSheet(reference.sheet, targetSheet);
    if (!onTarget) continue;
    const span = extent(reference, axis);
    if (!span) continue;
    if (span.start === span.end) continue; // одна ячейка просто едет дальше
    const touchesBefore = span.end === band.startRow - 1;
    const touchesAfter = span.start === band.startRow;
    if (touchesBefore || touchesAfter) missed.push(reference);
  }
  return missed;
}

/** Сколько ячеек содержат ошибку ссылки — считается до и после операции. */
export function countRefErrors(values: readonly (readonly unknown[])[]): number {
  let total = 0;
  for (const row of values) {
    for (const value of row) {
      if (typeof value === "string" && /^#(REF|ССЫЛКА)!$/i.test(value.trim())) total += 1;
    }
  }
  return total;
}

/** Непустые ячейки полосы — то, что удаление уничтожит без возврата. */
export function countFilled(values: readonly (readonly unknown[])[]): number {
  let total = 0;
  for (const row of values) {
    for (const value of row) {
      if (value !== "" && value !== null && value !== undefined) total += 1;
    }
  }
  return total;
}
