/**
 * Собственный стек отмены для изменений надстройки.
 * В стек попадают только действия, для которых можно выполнить безопасный и
 * достаточно точный обратный ход. Структурные операции строк и фильтры сюда
 * намеренно не добавляются: псевдо-undo опаснее отсутствия undo.
 */

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
  numberFormat?: string[][];
  bold?: Array<Array<boolean | null>>;
  fillColor?: Array<Array<string | null>>;
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
      range.formulas = restorableFormulas(before) as any[][];
      await ctx.sync();
    });
  });
}

/**
 * Точный снимок тех свойств форматирования, которые действительно меняются.
 * Формулы/значения НЕ сохраняются и НЕ восстанавливаются этим undo.
 */
export async function captureExactFormat(
  ctx: Excel.RequestContext,
  sheetName: string,
  address: string,
  fields: { numberFormat: boolean; bold: boolean; fillColor: boolean }
): Promise<ExactFormatSnapshot> {
  const range = ctx.workbook.worksheets.getItem(sheetName).getRange(address);
  range.load(["rowCount", "columnCount"]);
  if (fields.numberFormat) range.load("numberFormat");
  await ctx.sync();

  const cells: Excel.Range[][] = [];
  for (let r = 0; r < range.rowCount; r++) {
    const row: Excel.Range[] = [];
    for (let c = 0; c < range.columnCount; c++) {
      const cell = range.getCell(r, c);
      if (fields.bold) cell.format.font.load("bold");
      if (fields.fillColor) cell.format.fill.load("color");
      row.push(cell);
    }
    cells.push(row);
  }
  if (fields.bold || fields.fillColor) await ctx.sync();

  return {
    sheet: sheetName,
    address,
    rowCount: range.rowCount,
    columnCount: range.columnCount,
    ...(fields.numberFormat ? { numberFormat: range.numberFormat as string[][] } : {}),
    ...(fields.bold
      ? { bold: cells.map((row) => row.map((cell) => (cell.format.font.bold ?? null) as boolean | null)) }
      : {}),
    ...(fields.fillColor
      ? { fillColor: cells.map((row) => row.map((cell) => (cell.format.fill.color || null) as string | null)) }
      : {})
  };
}

function sameFormat(a: ExactFormatSnapshot, b: ExactFormatSnapshot): boolean {
  return (
    JSON.stringify(a.numberFormat ?? null) === JSON.stringify(b.numberFormat ?? null) &&
    JSON.stringify(a.bold ?? null) === JSON.stringify(b.bold ?? null) &&
    JSON.stringify(a.fillColor ?? null) === JSON.stringify(b.fillColor ?? null)
  );
}

async function applyExactFormat(ctx: Excel.RequestContext, snapshot: ExactFormatSnapshot): Promise<void> {
  const range = ctx.workbook.worksheets.getItem(snapshot.sheet).getRange(snapshot.address);
  if (snapshot.numberFormat) range.numberFormat = snapshot.numberFormat as any[][];

  for (let r = 0; r < snapshot.rowCount; r++) {
    for (let c = 0; c < snapshot.columnCount; c++) {
      const cell = range.getCell(r, c);
      const bold = snapshot.bold?.[r]?.[c];
      if (typeof bold === "boolean") cell.format.font.bold = bold;

      if (snapshot.fillColor) {
        const color = snapshot.fillColor[r]?.[c];
        if (color) cell.format.fill.color = color;
        else cell.format.fill.clear();
      }
    }
  }
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
        numberFormat: Boolean(after.numberFormat),
        bold: Boolean(after.bold),
        fillColor: Boolean(after.fillColor)
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
