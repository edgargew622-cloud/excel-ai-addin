/**
 * Собственный стек отмены для изменений надстройки.
 * В стек попадают только действия, для которых можно выполнить безопасный и
 * достаточно точный обратный ход. Структурные операции строк и фильтры сюда
 * намеренно не добавляются: псевдо-undo опаснее отсутствия undo.
 */

import { FORMAT_PROPERTY, loadFormat, readFormat, type FormatKey, type FormatSnapshot } from "./formatProps";

export interface UndoAction {
  id: string;
  label: string;
  at: number;
  undo: () => Promise<void>;
}

export interface ContentSnapshot {
  sheet: string;
  address: string;
  formulas: unknown[][];
  /** Значения и типы нужны, чтобы вернуть литеральный текст текстом. Без них
   * отмена записывала бы «2026-09-04» обратно как ввод, и Excel превращал бы
   * его в дату. Необязательны для совместимости со старыми снимками. */
  values?: unknown[][];
  valueTypes?: unknown[][];
}

/**
 * Готовит содержимое снимка к обратной записи через range.formulas.
 *
 * Проверка в Excel 17 сентября 2026 года: отмена сортировки вернула строки
 * на место, но текстовые даты «2026-09-04» стали настоящими датами. Запись
 * через formulas — это ввод, и Excel распознаёт в тексте даты и числа.
 *
 * Литеральный текст узнаётся по двум признакам: тип значения — строка, и текст
 * формулы совпадает со значением. У формулы, возвращающей текст, они
 * различаются: формула «="x"», значение «x». Литерал записывается с апострофом:
 * так Excel хранит его как текст, а сам апостроф в ячейке не показывает.
 */
export function restorableFormulas(snapshot: ContentSnapshot): unknown[][] {
  const { formulas, values, valueTypes } = snapshot;
  if (!values || !valueTypes) return formulas;
  return formulas.map((row, r) => row.map((formula, c) => {
    const value = values[r]?.[c];
    const type = String(valueTypes[r]?.[c] ?? "");
    if (type === "String" && typeof value === "string" && value !== "" && String(formula) === value) {
      return `'${value}`;
    }
    return formula;
  }));
}

export interface ExactFormatSnapshot {
  sheet: string;
  address: string;
  rowCount: number;
  columnCount: number;
  /** Какие свойства ячеек сняты: отмена возвращает ровно их. */
  keys: FormatKey[];
  /** Свойства каждой ячейки; у одной ячейки значение всегда однородно. */
  cells: FormatSnapshot[][];
  /** Ширина каждого столбца и высота каждой строки, если они менялись. */
  columns?: Array<number | null>;
  rows?: Array<number | null>;
}

const stack: UndoAction[] = [];
const MAX_DEPTH = 25;
let structuralRevision = 0;
let undoSafetyRevision = 0;
let undoMonitorReady = false;
const undoAvailabilityListeners = new Set<(ready: boolean) => void>();

function id() {
  return Math.random().toString(36).slice(2, 10);
}

function sameMatrix(a: unknown[][], b: unknown[][]): boolean {
  return JSON.stringify(a) === JSON.stringify(b);
}

export function push(item: UndoAction): boolean {
  if (!undoMonitorReady) return false;
  stack.push(item);
  if (stack.length > MAX_DEPTH) stack.shift();
  return true;
}

export function isCustomUndoAvailable(): boolean {
  return undoMonitorReady;
}

export function getUndoSafetyRevision(): number {
  return undoSafetyRevision;
}

export function subscribeUndoAvailability(listener: (ready: boolean) => void): () => void {
  undoAvailabilityListeners.add(listener);
  return () => undoAvailabilityListeners.delete(listener);
}

/**
 * Custom undo is safe only while the workbook structural-change monitor is active.
 * Disabling the monitor invalidates queued and in-flight undo work. Re-enabling
 * starts with an empty stack; stale entries are never restored.
 */
export function setUndoMonitorReady(ready: boolean): number {
  const removed = stack.length;
  undoMonitorReady = ready;
  undoSafetyRevision += 1;
  if (!ready) stack.length = 0;
  for (const listener of undoAvailabilityListeners) listener(ready);
  return removed;
}

export function action(label: string, undo: () => Promise<void>): UndoAction {
  return { id: id(), label, at: Date.now(), undo };
}

export async function captureContent(
  ctx: Excel.RequestContext,
  sheetName: string,
  address: string
): Promise<ContentSnapshot> {
  const sheet = ctx.workbook.worksheets.getItem(sheetName);
  const range = sheet.getRange(address);
  range.load(["formulas", "values", "valueTypes", "address"]);
  await ctx.sync();
  return {
    sheet: sheetName,
    address,
    formulas: range.formulas as unknown[][],
    values: range.values as unknown[][],
    valueTypes: range.valueTypes as unknown[][]
  };
}

export async function restoreContent(snapshot: ContentSnapshot): Promise<void> {
  await Excel.run(async (ctx) => {
    const range = ctx.workbook.worksheets.getItem(snapshot.sheet).getRange(snapshot.address);
    range.formulas = restorableFormulas(snapshot) as any[][];
    await ctx.sync();
  });
}

/** Больше этого поячеечную дозапись не делаем: это уже не исправление, а новая запись. */
export const MAX_REPAIR_CELLS = 500;

const sameCell = (a: unknown, b: unknown) => JSON.stringify(a ?? "") === JSON.stringify(b ?? "");

/**
 * Сверяет область с ожидаемым содержимым и один раз дописывает расхождения.
 *
 * Проверка в Excel 17 сентября 2026 года: отмена записала область таблицы
 * целиком, и формула `=1/0`, стоявшая в одной ячейке столбца, появилась во всех
 * строках. У таблиц Excel есть вычисляемые столбцы: формула в столбце, где
 * остальные ячейки пусты, протягивается сама. Запись всей области эту
 * протяжку вызывает, а поячеечная запись пустоты — нет, поэтому расхождения
 * дописываются по одной ячейке.
 *
 * Возвращает адреса ячеек, которые не сошлись и после дозаписи. Пустой
 * список означает, что область совпадает с ожидаемой.
 */
export async function repairMismatchedCells(
  ctx: Excel.RequestContext,
  range: Excel.Range,
  options: { property: "formulas" | "values"; expected: unknown[][]; toWrite: unknown[][] }
): Promise<string[]> {
  const { property, expected, toWrite } = options;
  const mismatches = async () => {
    range.load(property);
    await ctx.sync();
    const actual = (range as any)[property] as unknown[][];
    const found: [number, number][] = [];
    expected.forEach((row, r) => row.forEach((cell, c) => {
      if (!sameCell(actual?.[r]?.[c], cell)) found.push([r, c]);
    }));
    return found;
  };

  let found = await mismatches();
  if (found.length === 0) return [];
  if (found.length <= MAX_REPAIR_CELLS && typeof (range as any).getCell === "function") {
    for (const [r, c] of found) {
      const cell = range.getCell(r, c);
      (cell as any)[property] = [[toWrite[r]?.[c] ?? ""]];
    }
    await ctx.sync();
    found = await mismatches();
    if (found.length === 0) return [];
  }
  if (typeof (range as any).getCell !== "function") return found.map(([r, c]) => `R${r + 1}C${c + 1}`);
  const cells = found.slice(0, 20).map(([r, c]) => {
    const cell = range.getCell(r, c);
    cell.load("address");
    return cell;
  });
  await ctx.sync();
  const names = cells.map((cell) => String(cell.address));
  return found.length > names.length ? [...names, `и ещё ${found.length - names.length}`] : names;
}

/**
 * Отмена записи/сортировки выполняется только если диапазон всё ещё совпадает
 * с состоянием сразу после действия агента. Это не даёт затереть более свежую
 * ручную правку пользователя.
 */
export function guardedContentUndo(
  label: string,
  before: ContentSnapshot,
  after: ContentSnapshot
): UndoAction {
  return action(label, async () => {
    const expectedRevision = getStructuralRevision();
    const expectedSafetyRevision = getUndoSafetyRevision();
    if (!isCustomUndoAvailable()) throw new Error("Custom undo недоступен: монитор структуры книги не активен.");
    await Excel.run(async (ctx) => {
      const range = ctx.workbook.worksheets.getItem(after.sheet).getRange(after.address);
      range.load("formulas");
      await ctx.sync();
      if (getStructuralRevision() !== expectedRevision) {
        throw new Error("Структура книги изменилась во время отмены. Операция отмены остановлена.");
      }
      if (!isCustomUndoAvailable() || getUndoSafetyRevision() !== expectedSafetyRevision) {
        throw new Error("Монитор структуры книги отключился во время отмены. Операция отмены остановлена.");
      }
      if (!sameMatrix(range.formulas as unknown[][], after.formulas)) {
        throw new Error(
          `Диапазон ${after.sheet}!${after.address} изменён после операции агента. ` +
            "Автоматическая отмена остановлена, чтобы не затереть более свежие изменения."
        );
      }
      const toWrite = restorableFormulas(before);
      range.formulas = toWrite as any[][];
      await ctx.sync();
      // Отмена прежде ничего не проверяла после записи и молча считала дело
      // сделанным — так размноженная формула и прошла незамеченной.
      const remaining = await repairMismatchedCells(ctx, range, {
        property: "formulas",
        expected: before.formulas,
        toWrite
      });
      if (remaining.length) {
        throw new Error(
          `Отмена записала ${after.sheet}!${after.address}, но ячейки ${remaining.join(", ")} не совпадают с исходным состоянием. ` +
            "Проверьте их вручную: Excel мог протянуть формулы по столбцу таблицы."
        );
      }
    });
  });
}

/**
 * Точный снимок тех свойств оформления, которые действительно меняются.
 *
 * Свойства ячеек снимаются поячеечно: у области Excel сообщает только
 * однородное значение или null, а вернуть по null нечего. Ширина и высота
 * живут у столбцов и строк и снимаются по ним.
 * Формулы и значения этот снимок не хранит и не восстанавливает.
 */
export async function captureExactFormat(
  ctx: Excel.RequestContext,
  sheetName: string,
  address: string,
  fields: { keys: readonly FormatKey[]; columns?: boolean; rows?: boolean }
): Promise<ExactFormatSnapshot> {
  const range = ctx.workbook.worksheets.getItem(sheetName).getRange(address);
  range.load(["rowCount", "columnCount"]);
  await ctx.sync();

  const one = { rowCount: 1, columnCount: 1 };
  const cellKeys = fields.keys.filter((key) => FORMAT_PROPERTY.get(key)?.scope === "cell");
  const cells: any[][] = [];
  for (let r = 0; r < range.rowCount; r++) {
    const row: any[] = [];
    for (let c = 0; c < range.columnCount; c++) {
      const cell = range.getCell(r, c);
      loadFormat(cell, cellKeys, one);
      row.push(cell);
    }
    cells.push(row);
  }
  const columns = fields.columns
    ? Array.from({ length: range.columnCount }, (_, index) => {
        const column = range.getColumn(index);
        column.format.load("columnWidth");
        return column;
      })
    : null;
  const rows = fields.rows
    ? Array.from({ length: range.rowCount }, (_, index) => {
        const row = range.getRow(index);
        row.format.load("rowHeight");
        return row;
      })
    : null;
  await ctx.sync();

  return {
    sheet: sheetName,
    address,
    rowCount: range.rowCount,
    columnCount: range.columnCount,
    keys: cellKeys,
    cells: cells.map((row) => row.map((cell) => readFormat(cell, cellKeys, one))),
    ...(columns ? { columns: columns.map((column) => (column.format.columnWidth ?? null) as number | null) } : {}),
    ...(rows ? { rows: rows.map((row) => (row.format.rowHeight ?? null) as number | null) } : {})
  };
}

function sameFormat(a: ExactFormatSnapshot, b: ExactFormatSnapshot): boolean {
  return (
    JSON.stringify(a.cells) === JSON.stringify(b.cells) &&
    JSON.stringify(a.columns ?? null) === JSON.stringify(b.columns ?? null) &&
    JSON.stringify(a.rows ?? null) === JSON.stringify(b.rows ?? null)
  );
}

async function applyExactFormat(ctx: Excel.RequestContext, snapshot: ExactFormatSnapshot): Promise<void> {
  const range = ctx.workbook.worksheets.getItem(snapshot.sheet).getRange(snapshot.address);
  const one = { rowCount: 1, columnCount: 1 };

  for (let r = 0; r < snapshot.rowCount; r++) {
    for (let c = 0; c < snapshot.columnCount; c++) {
      const cell = range.getCell(r, c);
      const saved = snapshot.cells[r]?.[c] ?? {};
      for (const key of snapshot.keys) {
        const property = FORMAT_PROPERTY.get(key);
        const value = saved[key];
        if (!property) continue;
        if (property.restore) { property.restore(cell, value, one); continue; }
        // У одной ячейки null значит «Excel не сообщил» — писать по нему нечего.
        if (value === null || value === undefined) continue;
        property.write(cell, value, one);
      }
    }
  }
  snapshot.columns?.forEach((width, index) => {
    if (typeof width === "number") range.getColumn(index).format.columnWidth = width;
  });
  snapshot.rows?.forEach((height, index) => {
    if (typeof height === "number") range.getRow(index).format.rowHeight = height;
  });
  await ctx.sync();
}

export function exactFormatUndo(
  label: string,
  before: ExactFormatSnapshot,
  after: ExactFormatSnapshot
): UndoAction {
  return action(label, async () => {
    const expectedRevision = getStructuralRevision();
    const expectedSafetyRevision = getUndoSafetyRevision();
    if (!isCustomUndoAvailable()) throw new Error("Custom undo недоступен: монитор структуры книги не активен.");
    await Excel.run(async (ctx) => {
      const current = await captureExactFormat(ctx, after.sheet, after.address, {
        keys: after.keys,
        columns: Boolean(after.columns),
        rows: Boolean(after.rows)
      });
      if (getStructuralRevision() !== expectedRevision) {
        throw new Error("Структура книги изменилась во время отмены. Операция отмены остановлена.");
      }
      if (!isCustomUndoAvailable() || getUndoSafetyRevision() !== expectedSafetyRevision) {
        throw new Error("Монитор структуры книги отключился во время отмены. Операция отмены остановлена.");
      }
      if (!sameFormat(current, after)) {
        throw new Error(
          `Формат ${after.sheet}!${after.address} изменён после операции агента. ` +
            "Автоматическая отмена остановлена, чтобы не затереть более свежие изменения."
        );
      }
      await applyExactFormat(ctx, before);
    });
  });
}

export function peek(): UndoAction | undefined {
  return stack[stack.length - 1];
}

export function depth() {
  return stack.length;
}

export function clear() {
  stack.length = 0;
}

/**
 * Структурные изменения строк могут сдвинуть адреса любых ранее сохранённых
 * диапазонов и изменить ссылки формул. Безопасно пересчитать такие адреса
 * невозможно без полноценной модели зависимостей Excel, поэтому после
 * insert/delete инвалидируем весь custom-undo стек. Это консервативно, но
 * исключает отмену в уже другой ячейке.
 */
export function invalidateAfterStructuralChange(): number {
  structuralRevision += 1;
  const removed = stack.length;
  stack.length = 0;
  return removed;
}

/** Monotonic epoch used to invalidate an undo that is already in flight. */
export function getStructuralRevision(): number {
  return structuralRevision;
}

export function describeNextUndo(): string | null {
  return peek()?.label ?? null;
}

export async function undoLast(): Promise<string> {
  if (!isCustomUndoAvailable()) {
    throw new Error("Custom undo недоступен: монитор структурных изменений Excel не активен.");
  }
  const item = stack.pop();
  if (!item) return "Отменять нечего.";
  const revisionAtStart = getStructuralRevision();
  const safetyRevisionAtStart = getUndoSafetyRevision();

  try {
    await item.undo();
    if (getStructuralRevision() !== revisionAtStart) {
      throw new Error("Структура книги изменилась во время отмены. Результат требует проверки пользователем.");
    }
    if (!isCustomUndoAvailable() || getUndoSafetyRevision() !== safetyRevisionAtStart) {
      throw new Error("Монитор структуры книги отключился во время отмены. Операция отмены остановлена.");
    }
    return `Отменено: ${item.label}.`;
  } catch (error) {
    // Возвращать действие в стек можно только если и структура, и монитор
    // остались теми же. Иначе A1-адрес или сама гарантия безопасности устарели.
    if (
      isCustomUndoAvailable() &&
      getStructuralRevision() === revisionAtStart &&
      getUndoSafetyRevision() === safetyRevisionAtStart
    ) {
      stack.push(item);
    }
    throw error;
  }
}
