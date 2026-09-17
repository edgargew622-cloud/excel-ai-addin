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
}

export function rowBand(startRow: number, count: number): RowBand {
  return { startRow, endRow: startRow + count - 1 };
}

const SHEET = "(?:'((?:[^']|'')+)'|([A-Za-z_\u0400-\u04FF][A-Za-z0-9_.\u0400-\u04FF ]*))!";
const CELL = String.raw`(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})`;
const REFERENCE = new RegExp(`(?:${SHEET})?${CELL}(?::${CELL})?`, "g");

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
    if (!match || match.index !== index) { index += 1; continue; }

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
    found.push({
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

/** Формулы со структурированными ссылками таблиц разобрать нельзя. */
export function usesTableReference(formula: unknown): boolean {
  return typeof formula === "string" && /[A-Za-z_Ѐ-ӿ][\w.Ѐ-ӿ]*\[/.test(formula);
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
export function deleteImpact(
  formula: unknown,
  formulaSheet: string,
  targetSheet: string,
  band: RowBand
): DeleteImpact {
  const broken: FormulaReference[] = [];
  const shrunk: FormulaReference[] = [];
  for (const reference of formulaReferences(formula)) {
    const onTarget = reference.sheet === null
      ? sameSheet(null, formulaSheet) && sameSheet(formulaSheet, targetSheet)
      : sameSheet(reference.sheet, targetSheet);
    if (!onTarget) continue;
    if (reference.rowStart >= band.startRow && reference.rowEnd <= band.endRow) broken.push(reference);
    else if (reference.rowStart <= band.endRow && reference.rowEnd >= band.startRow) shrunk.push(reference);
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
  band: RowBand
): FormulaReference[] {
  const missed: FormulaReference[] = [];
  for (const reference of formulaReferences(formula)) {
    const onTarget = reference.sheet === null
      ? sameSheet(formulaSheet, targetSheet)
      : sameSheet(reference.sheet, targetSheet);
    if (!onTarget) continue;
    if (reference.rowStart === reference.rowEnd) continue; // одна ячейка просто едет вниз
    const touchesBelow = reference.rowEnd === band.startRow - 1;
    const touchesAbove = reference.rowStart === band.startRow;
    if (touchesBelow || touchesAbove) missed.push(reference);
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
