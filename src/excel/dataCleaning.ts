/**
 * Очистка данных (этап 7, срез 7.2): профиль, а затем операции очистки.
 *
 * Профиль только читает. Большая таблица читается порциями по строкам,
 * и ответ называет проверенную область: если прочитано не всё, это сказано
 * прямо, а вывод по части не выдаётся за вывод по всей таблице.
 */

import { intersects, parseA1Rect } from "./a1";
import { cleanText, dateFormatCode, dateFromText, excelSerial, formatDate, numberFromText, planDuplicates, SKIP_TEXT, type SkipReason } from "./cleanModel";
import { cultureDateOrder, profileData, type ColumnProfile, type DataProfile, type DateOrder, type NumberCulture } from "./dataProfile";
import {
  MAX_IO_CELLS,
  assertPlanWorkbook,
  assertTargetWritable,
  checkAddress,
  cloneMatrix,
  deepFreeze,
  preflightToolArgs,
  rangeOf,
  readTableRanges,
  scanWorkbookFormulas,
  ToolError,
  ToolExecutionError,
  valuesForLiteralWrite
} from "./excelTools";
import { columnLetters } from "./formulaFill";
import { shiftedReferences } from "./rowOps";
import { lastWorkbookBackup } from "./workbookBackup";
import { action, captureContent, guardedContentUndo, isCustomUndoAvailable, push } from "./undo";
import { captureTarget, officeCapabilities, type WorkbookTarget } from "./workbookContext";

/** Больше этого профиль не читает: ответ был бы огромным, а время — долгим. */
export const MAX_PROFILE_CELLS = 100_000;

export type WorkbookCulture = NumberCulture & { dateOrder: DateOrder | null; dateSeparator: string; shortDatePattern: string; name: string };

/** Разделители и порядок даты книги — как их сообщает Excel (ExcelApi 1.12). */
export async function readCulture(ctx: Excel.RequestContext): Promise<WorkbookCulture> {
  const culture = ctx.workbook.application.cultureInfo;
  culture.load("name");
  culture.numberFormat.load(["numberDecimalSeparator", "numberGroupSeparator"]);
  culture.datetimeFormat.load(["dateSeparator", "shortDatePattern"]);
  await ctx.sync();
  return {
    name: String(culture.name),
    decimal: String(culture.numberFormat.numberDecimalSeparator),
    group: String(culture.numberFormat.numberGroupSeparator),
    dateSeparator: String(culture.datetimeFormat.dateSeparator),
    shortDatePattern: String(culture.datetimeFormat.shortDatePattern),
    dateOrder: cultureDateOrder(String(culture.datetimeFormat.shortDatePattern))
  };
}

const FINDINGS: (keyof ColumnProfile)[] = [
  "edgeSpaces", "innerSpaces", "nonBreakingSpaces",
  "numbersAsText", "foreignNumbersAsText", "ambiguousNumbersAsText", "codesWithLeadingZeros",
  "datesAsText", "ambiguousDatesAsText"
];

/** Ответ модели: пустые находки опускаются, чтобы не тратить её контекст. */
function compact(profile: DataProfile) {
  return profile.columns.map((column) => {
    const findings = Object.fromEntries(FINDINGS
      .map((key) => [key, column[key]] as const)
      .filter(([, value]) => (value as { count: number }).count > 0));
    const { edgeSpaces, innerSpaces, nonBreakingSpaces, numbersAsText, foreignNumbersAsText, ambiguousNumbersAsText,
      codesWithLeadingZeros, datesAsText, ambiguousDatesAsText, ...counts } = column;
    return { ...counts, ...(Object.keys(findings).length ? { findings } : {}) };
  });
}

export async function profileRange(args: { sheet?: string; address?: string; hasHeaders?: boolean }) {
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
      if ((used as any).isNullObject) return { sheet: sheet.name, empty: true, note: "Лист пуст: профилировать нечего." };
      area = used;
    }
    area.load(["address", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
    await ctx.sync();
    const culture = await readCulture(ctx);

    // Порции по строкам, не больше MAX_IO_CELLS ячеек за чтение.
    const rowsPerChunk = Math.max(1, Math.floor(MAX_IO_CELLS / area.columnCount));
    const readableRows = Math.min(area.rowCount, Math.floor(MAX_PROFILE_CELLS / area.columnCount));
    if (readableRows < 1) throw new ToolError(`В ${area.address} слишком много столбцов для профиля: ${area.columnCount}.`);
    const values: unknown[][] = [];
    const formulas: unknown[][] = [];
    const valueTypes: unknown[][] = [];
    const numberFormat: unknown[][] = [];
    for (let start = 0; start < readableRows; start += rowsPerChunk) {
      const rows = Math.min(rowsPerChunk, readableRows - start);
      const chunk = sheet.getRangeByIndexes(area.rowIndex + start, area.columnIndex, rows, area.columnCount);
      chunk.load(["values", "formulas", "valueTypes", "numberFormat"]);
      await ctx.sync();
      values.push(...(chunk.values as unknown[][]));
      formulas.push(...(chunk.formulas as unknown[][]));
      valueTypes.push(...(chunk.valueTypes as unknown[][]));
      numberFormat.push(...(chunk.numberFormat as unknown[][]));
    }
    const lastColumn = columnLetters(area.columnIndex + area.columnCount);
    const checked = `${columnLetters(area.columnIndex + 1)}${area.rowIndex + 1}:${lastColumn}${area.rowIndex + readableRows}`;
    const incomplete = readableRows < area.rowCount;

    const profile = profileData({
      values, formulas, valueTypes, numberFormat,
      hasHeaders: args.hasHeaders !== false,
      origin: { rowIndex: area.rowIndex, columnIndex: area.columnIndex },
      address: checked,
      culture,
      columnName: columnLetters
    });

    return {
      sheet: sheet.name,
      address: area.address.slice(area.address.lastIndexOf("!") + 1),
      checkedAddress: checked,
      incomplete,
      ...(incomplete
        ? { incompleteNote: `Прочитаны первые ${readableRows} строк из ${area.rowCount}: профиль ограничен ${MAX_PROFILE_CELLS} ячейками. Выводы — только о проверенной области.` }
        : {}),
      hasHeaders: profile.hasHeaders,
      rows: profile.rows,
      emptyRows: profile.emptyRows,
      culture: { name: culture.name, decimal: culture.decimal, group: culture.group === " " ? "пробел" : culture.group, dateOrder: culture.dateOrder },
      columns: compact(profile),
      duplicateRows: profile.duplicateRows,
      duplicateRowsNormalized: profile.duplicateRowsNormalized,
      note:
        "Счётчики посчитаны панелью по прочитанным значениям. numbersAsText и datesAsText — однозначные числа и даты, " +
        "записанные текстом: в расчёт они не входят. ambiguous* и foreignNumbersAsText без ответа пользователя не преобразуются: " +
        "01.02.2026 может быть и 1 февраля, и 2 января, 1,500 — и полутора, и тысячей пятьсот. codesWithLeadingZeros — коды, " +
        "их нельзя превращать в числа. duplicateRowsNormalized — кандидаты после обрезки пробелов и без учёта регистра, не доказанные дубликаты."
    };
  });
}

/* =========================================================================
 * Операции очистки: лишние пробелы (7.2.2), числа и даты из текста (7.2.3)
 * ========================================================================= */

export type CleanToolName = "trim_text" | "convert_values";

export interface CleanChange {
  r: number;
  c: number;
  cell: string;
  before: string;
  after: string | number;
  /** Как дата должна выглядеть после записи — по ней сверяется система дат книги. */
  display?: string;
}

export interface CleanValuesPlan {
  readonly kind: CleanToolName;
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly resolvedAddress: string;
  readonly rows: number;
  readonly columns: number;
  readonly description: string;
  readonly changes: readonly CleanChange[];
  /** Формат даты, который получат изменённые ячейки: без него дата видна числом. */
  readonly numberFormat?: string;
  readonly skipped: Readonly<Record<string, { count: number; examples: readonly string[] }>>;
  readonly sample: readonly string[];
  readonly beforeFormulas: readonly (readonly unknown[])[];
  readonly undoAvailable: boolean;
  readonly undoNote?: string;
  readonly createdAt: string;
}

const PREVIEW_PAIRS = 8;
const MAX_SKIP_EXAMPLES = 3;
/** Сколько ячеек пишется за одну синхронизацию: так Excel не подвисает. */
const WRITE_CHUNK = 2_000;

type Decision = { after: string | number; display?: string } | { skip: SkipReason };

async function prepareCleanPlan(
  name: CleanToolName,
  args: { sheet?: string; address: string },
  description: (culture: WorkbookCulture | null) => string,
  decide: (text: string, culture: WorkbookCulture | null) => Decision,
  options: { needsCulture: boolean; numberFormat?: (culture: WorkbookCulture) => string }
): Promise<CleanValuesPlan> {
  preflightToolArgs(name, args);
  const address = checkAddress(args.address);
  const target = await captureTarget(args.sheet);
  const prepared = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    const range = await rangeOf(ctx, sheet, address);
    range.load(["address", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
    sheet.load(["id", "name"]);
    try {
      range.format?.protection?.load("locked");
      sheet.protection?.load("protected");
    } catch { /* среда без сведений о защите */ }
    await ctx.sync();
    const cells = range.rowCount * range.columnCount;
    if (cells > MAX_IO_CELLS) {
      throw new ToolError(`Очистка ограничена ${MAX_IO_CELLS} ячейками за операцию; ${range.address} содержит ${cells}. Разбейте область по строкам.`);
    }
    assertTargetWritable(sheet, range);
    range.load(["formulas", "values", "valueTypes"]);
    await ctx.sync();
    const culture = options.needsCulture ? await readCulture(ctx) : null;

    const changes: CleanChange[] = [];
    const skipped: Record<string, { count: number; examples: string[] }> = {};
    const skip = (reason: SkipReason, where: string) => {
      const item = (skipped[SKIP_TEXT[reason]] ??= { count: 0, examples: [] });
      item.count += 1;
      if (item.examples.length < MAX_SKIP_EXAMPLES) item.examples.push(where);
    };
    const formulas = range.formulas as unknown[][];
    const values = range.values as unknown[][];
    const types = range.valueTypes as unknown[][];
    for (let r = 0; r < range.rowCount; r++) {
      for (let c = 0; c < range.columnCount; c++) {
        const cell = `${columnLetters(range.columnIndex + c + 1)}${range.rowIndex + r + 1}`;
        const formula = formulas[r][c];
        const value = values[r][c];
        if (value === "" || types[r][c] === "Empty") continue;
        if (typeof formula === "string" && formula.startsWith("=")) { skip("formula", cell); continue; }
        if (types[r][c] !== "String" || typeof value !== "string") { skip("notText", cell); continue; }
        const decision = decide(value, culture);
        if ("skip" in decision) { skip(decision.skip, `${cell}: «${value}»`); continue; }
        if (decision.after === value) continue;
        changes.push({ r, c, cell, before: value, after: decision.after, ...(decision.display ? { display: decision.display } : {}) });
      }
    }
    const skippedText = Object.entries(skipped).map(([reason, item]) => `${reason} — ${item.count} (${item.examples.join("; ")})`).join("; ");
    if (!changes.length) {
      throw new ToolError(`В ${sheet.name}!${range.address.replace(/^.*!/, "")} менять нечего. ${skippedText ? `Пропущено: ${skippedText}.` : ""} Операция не выполнялась.`);
    }
    const undo = isCustomUndoAvailable();
    return {
      kind: name,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      target: { ...target, sheetName: sheet.name },
      resolvedAddress: range.address.replace(/^.*!/, ""),
      rows: range.rowCount,
      columns: range.columnCount,
      description: description(culture),
      changes,
      ...(options.numberFormat && culture ? { numberFormat: options.numberFormat(culture) } : {}),
      skipped,
      sample: changes.slice(0, PREVIEW_PAIRS).map((item) => `${item.cell}: «${item.before}» → ${typeof item.after === "number" ? (item.display ?? String(item.after)) : `«${item.after}»`}`),
      beforeFormulas: cloneMatrix(formulas),
      undoAvailable: undo,
      ...(undo ? {} : { undoNote: "Отмена недоступна: монитор изменений Excel не активен." }),
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export function prepareTrimTextPlan(args: unknown): Promise<CleanValuesPlan> {
  const a = (args ?? {}) as { sheet?: string; address: string; collapseInner?: boolean };
  const collapse = a.collapseInner !== false;
  return prepareCleanPlan(
    "trim_text",
    a,
    () => collapse
      ? "Удалить пробелы по краям и повторы внутри, неразрывные пробелы заменить обычными — как функция TRIM. Текст остаётся текстом."
      : "Удалить пробелы по краям, неразрывные пробелы заменить обычными; пробелы внутри не трогать. Текст остаётся текстом.",
    (text) => ({ after: cleanText(text, collapse) }),
    { needsCulture: false }
  );
}

export function prepareConvertValuesPlan(args: unknown): Promise<CleanValuesPlan> {
  const a = (args ?? {}) as { sheet?: string; address: string; to: "number" | "date"; decimalSeparator?: "," | "."; dateOrder?: DateOrder };
  if (a.to === "date") {
    return prepareCleanPlan(
      "convert_values",
      a,
      (culture) => `Текст → даты${a.dateOrder ? ` по порядку ${a.dateOrder}, названному пользователем` : " (только однозначные)"}. ` +
        `Ячейки получат числовое значение даты и формат ${dateFormatCode(culture ? shortPattern(culture) : "")}: меняется и значение, и отображение.`,
      (text, culture) => {
        const result = dateFromText(text, a.dateOrder);
        if ("skip" in result) return result;
        return { after: excelSerial(result.date), display: formatDate(result.date, dateFormatCode(shortPattern(culture!))) };
      },
      { needsCulture: true, numberFormat: (culture) => dateFormatCode(shortPattern(culture)) }
    );
  }
  return prepareCleanPlan(
    "convert_values",
    a,
    (culture) => `Текст → числа по разделителям ${a.decimalSeparator ? `названным пользователем: десятичный «${a.decimalSeparator}»` : `книги: десятичный «${culture?.decimal}», тысячи — ${culture?.group === " " ? "пробел" : `«${culture?.group}»`}`}. ` +
      "Меняется значение; формат ячеек не трогается.",
    (text, culture) => {
      const result = numberFromText(text, culture!, a.decimalSeparator);
      return "skip" in result ? result : { after: result.value };
    },
    { needsCulture: true }
  );
}

function shortPattern(culture: WorkbookCulture): string {
  return culture.shortDatePattern;
}


/** Исполнение общее: книга та же, область не менялась, пишутся только изменившиеся ячейки, каждая сверяется. */
export async function executeCleanPlan(plan: CleanValuesPlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(plan.target.sheetId);
    sheet.load(["id", "name"]);
    const range = sheet.getRange(plan.resolvedAddress);
    range.load(["formulas", "numberFormat", "rowIndex", "columnIndex"]);
    await ctx.sync();
    const where = `${sheet.name}!${plan.resolvedAddress}`;
    if (JSON.stringify(range.formulas) !== JSON.stringify(plan.beforeFormulas)) {
      throw new ToolExecutionError(`Данные ${where} изменились после предпросмотра. Очистка не выполнялась — сделайте новый предпросмотр.`, "failed_before_write");
    }
    const formatsBefore = range.numberFormat as unknown[][];
    const before = plan.undoAvailable ? await captureContent(ctx, sheet.name, plan.resolvedAddress) : null;

    try {
      for (let start = 0; start < plan.changes.length; start += WRITE_CHUNK) {
        for (const change of plan.changes.slice(start, start + WRITE_CHUNK)) {
          const cell = range.getCell(change.r, change.c);
          // Текст пишется с апострофом: «007» и «1 200» после обрезки
          // пробелов должны остаться текстом, а не стать числами.
          cell.values = [[typeof change.after === "string" ? valuesForLiteralWrite([[change.after]])[0][0] : change.after]] as any[][];
          if (plan.numberFormat) cell.numberFormat = [[plan.numberFormat]] as any[][];
        }
        await ctx.sync();
      }
    } catch (error: any) {
      throw new ToolExecutionError(`Не удалось определить итог очистки ${where}: ${error?.message ?? error}. Перечитайте область.`, "unknown");
    }

    range.load(["values", "valueTypes", "formulas", "text"]);
    await ctx.sync();
    const values = range.values as unknown[][];
    const types = range.valueTypes as unknown[][];
    const texts = range.text as unknown[][];
    const formulas = range.formulas as unknown[][];
    const mismatches: string[] = [];
    const changed = new Set(plan.changes.map((item) => `${item.r}:${item.c}`));
    for (const change of plan.changes) {
      const value = values[change.r][change.c];
      const ok = typeof change.after === "string"
        ? (change.after === "" ? value === "" : value === change.after && types[change.r][change.c] === "String")
        : typeof value === "number" && Math.abs(value - change.after) < 1e-9 &&
          (change.display === undefined || texts[change.r][change.c] === change.display);
      if (!ok) mismatches.push(`${change.cell}: ${JSON.stringify(value)}${change.display ? ` («${texts[change.r][change.c]}»)` : ""} вместо ${change.display ?? JSON.stringify(change.after)}`);
    }
    // Ячейки вне плана трогаться не должны.
    plan.beforeFormulas.forEach((row, r) => row.forEach((formula, c) => {
      if (!changed.has(`${r}:${c}`) && String(formulas[r][c]) !== String(formula)) {
        mismatches.push(`${columnLetters(range.columnIndex + c + 1)}${range.rowIndex + r + 1}: ячейка вне плана изменилась`);
      }
    }));

    let undoRecorded = false;
    if (before) {
      const after = await captureContent(ctx, sheet.name, plan.resolvedAddress);
      const content = guardedContentUndo(plan.kind === "trim_text" ? "удаление лишних пробелов" : "преобразование значений", before, after);
      undoRecorded = push(plan.numberFormat
        ? action(content.label, async () => {
            await content.undo();
            // Формат дат вернуть отдельно: содержимое его не несёт.
            await Excel.run(async (undoCtx) => {
              const undoRange = undoCtx.workbook.worksheets.getItem(plan.target.sheetId).getRange(plan.resolvedAddress);
              for (const change of plan.changes) undoRange.getCell(change.r, change.c).numberFormat = [[formatsBefore[change.r][change.c]]] as any[][];
              await undoCtx.sync();
            });
          })
        : content);
    }

    if (mismatches.length) {
      throw new ToolExecutionError(
        `Очистка ${where} выполнена не полностью: ${mismatches.length} расхождений (${mismatches.slice(0, 5).join("; ")}). ` +
        (undoRecorded ? "Её можно отменить кнопкой «Отменить». " : "") + "Перечитайте область; повтор без проверки не поможет.",
        "applied"
      );
    }
    return {
      ok: true,
      executionState: "verified",
      sheet: sheet.name,
      address: plan.resolvedAddress,
      operation: plan.description,
      changedCells: plan.changes.length,
      checkedCells: plan.changes.length,
      sample: plan.sample,
      ...(Object.keys(plan.skipped).length ? { skipped: plan.skipped } : {}),
      ...(plan.numberFormat ? { numberFormat: plan.numberFormat } : {}),
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: plan.undoNote ?? "Автоматическая отмена недоступна." })
    };
  });
}

/* =========================================================================
 * Удаление дубликатов (7.2.4)
 * ========================================================================= */

export interface ShiftRisk {
  sheet: string;
  cell: string;
  formula: string;
  references: string[];
}

export interface RemoveDuplicatesPlan {
  readonly kind: "remove_duplicates";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly resolvedAddress: string;
  readonly hasHeaders: boolean;
  /** Номера столбцов ключа внутри области (от 0) и их названия. */
  readonly keyColumns: readonly number[];
  readonly keyNames: readonly string[];
  readonly dataRows: number;
  /** Удаляемые строки листа и строка, повтором которой каждая является. */
  readonly removed: readonly { row: number; duplicateOf: number }[];
  readonly sampleRemoved: readonly string[];
  /** Какими станут строки данных области — по ним сверяется итог. */
  readonly expected: readonly (readonly unknown[])[];
  readonly beforeFormulas: readonly (readonly unknown[])[];
  /** Формулы книги, которые после сдвига увидят другие данные. */
  readonly risks: readonly ShiftRisk[];
  readonly riskOverflow: number;
  readonly unscannedSheets: readonly string[];
  readonly backup: { name: string; at: string } | null;
  readonly undoAvailable: false;
  readonly undoNote: string;
  readonly createdAt: string;
}

const MAX_SHIFT_RISKS = 20;
const SAMPLE_REMOVED = 8;

/** Данные вплотную слева и справа от области: после сдвига строки разъехались бы с ними. */
async function neighbourData(ctx: Excel.RequestContext, sheet: Excel.Worksheet, area: { rowIndex: number; columnIndex: number; rowCount: number; columnCount: number }): Promise<string[]> {
  const sides: { name: string; range: Excel.Range }[] = [];
  if (area.columnIndex > 0) sides.push({ name: columnLetters(area.columnIndex), range: sheet.getRangeByIndexes(area.rowIndex, area.columnIndex - 1, area.rowCount, 1) });
  if (area.columnIndex + area.columnCount < 16_384) {
    sides.push({ name: columnLetters(area.columnIndex + area.columnCount + 1), range: sheet.getRangeByIndexes(area.rowIndex, area.columnIndex + area.columnCount, area.rowCount, 1) });
  }
  for (const side of sides) side.range.load("values");
  await ctx.sync();
  return sides.filter((side) => (side.range.values as unknown[][]).some((row) => row[0] !== "" && row[0] !== null)).map((side) => side.name);
}

export async function prepareRemoveDuplicatesPlan(args: unknown): Promise<RemoveDuplicatesPlan> {
  preflightToolArgs("remove_duplicates", args);
  const a = args as { sheet?: string; address: string; columns?: string[]; hasHeaders?: boolean };
  const address = checkAddress(a.address);
  const hasHeaders = a.hasHeaders !== false;
  const target = await captureTarget(a.sheet);
  const prepared = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    const range = await rangeOf(ctx, sheet, address);
    range.load(["address", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
    sheet.load(["id", "name"]);
    try {
      range.format?.protection?.load("locked");
      sheet.protection?.load("protected");
    } catch { /* среда без сведений о защите */ }
    await ctx.sync();
    const where = `${sheet.name}!${range.address.replace(/^.*!/, "")}`;
    const cells = range.rowCount * range.columnCount;
    if (cells > MAX_IO_CELLS) throw new ToolError(`Удаление дубликатов ограничено ${MAX_IO_CELLS} ячейками; ${where} содержит ${cells}.`);
    assertTargetWritable(sheet, range);
    range.load(["formulas", "values"]);
    await ctx.sync();
    const formulas = range.formulas as unknown[][];
    const values = range.values as unknown[][];

    // Замер: removeDuplicates сдвигает значения внутри области и ссылок
    // не подстраивает. Формулы в самой области после сдвига считали бы
    // по чужим строкам — такие области не берём.
    const formulaCells = formulas.flat().filter((item) => typeof item === "string" && item.startsWith("=")).length;
    if (formulaCells) throw new ToolError(`В ${where} есть формулы (${formulaCells}): удаление дубликатов сдвигает значения, не подстраивая ссылки. Операция не выполнялась.`);
    const tables = await readTableRanges(ctx, sheet);
    const areaRect = parseA1Rect(range.address.replace(/^.*!/, ""))!;
    const table = tables.find((item) => { const rect = parseA1Rect(item.address.replace(/^.*!/, "")); return rect ? intersects(rect, areaRect) : false; });
    if (table) throw new ToolError(`${where} задевает таблицу Excel «${table.name}»: у таблицы свои правила строк, и здесь она не поддержана. Операция не выполнялась.`);
    const neighbours = await neighbourData(ctx, sheet, range);
    if (neighbours.length) {
      throw new ToolError(
        `Рядом с ${where} есть данные в столбцах ${neighbours.join(", ")}: удаление дубликатов сдвигает строки только внутри области, ` +
        "и строки разъехались бы с соседними столбцами. Расширьте область на всю таблицу. Операция не выполнялась."
      );
    }

    const headers = hasHeaders ? values[0].map((value) => String(value ?? "").trim()) : [];
    const letter = (index: number) => columnLetters(range.columnIndex + index + 1);
    const keyColumns = (a.columns?.length ? a.columns : null)?.map((name) => {
      const wanted = String(name).trim().toLowerCase();
      const byHeader = headers.findIndex((header) => header.toLowerCase() === wanted);
      if (byHeader >= 0) return byHeader;
      const byLetter = Array.from({ length: range.columnCount }, (_, index) => letter(index)).findIndex((item) => item.toLowerCase() === wanted);
      if (byLetter >= 0) return byLetter;
      throw new ToolError(`Столбца «${name}» в ${where} нет. ${hasHeaders ? `Заголовки: ${headers.map((item) => `«${item}»`).join(", ")}.` : ""}`);
    }) ?? Array.from({ length: range.columnCount }, (_, index) => index);

    const body = values.slice(hasHeaders ? 1 : 0);
    const firstDataRow = range.rowIndex + (hasHeaders ? 2 : 1);
    const duplicates = planDuplicates(body, keyColumns);
    if (!duplicates.removed.length) {
      const normalized = planDuplicates(body.map((row) => row.map((value) => (typeof value === "string" ? cleanText(value, true) : value))), keyColumns);
      throw new ToolError(
        `В ${where} дубликатов по ключу нет — удалять нечего.` +
        (normalized.removed.length ? ` Но после удаления лишних пробелов их было бы ${normalized.removed.length}: Excel считает «Москва » и «Москва» разными. Предложите сначала trim_text.` : "")
      );
    }

    // Какие формулы книги после сдвига увидят другие данные.
    const scan = await scanWorkbookFormulas(ctx);
    const firstRemovedRow = firstDataRow + Math.min(...duplicates.removed.map((item) => item.row));
    const zone = {
      rowStart: range.rowIndex + 1,
      rowEnd: range.rowIndex + range.rowCount,
      columnStart: range.columnIndex + 1,
      columnEnd: range.columnIndex + range.columnCount,
      dataRowStart: firstDataRow
    };
    const risks: ShiftRisk[] = [];
    let overflow = 0;
    for (const scanned of scan.sheets) {
      scanned.formulas.forEach((row, r) => row.forEach((formula, c) => {
        const hit = shiftedReferences(formula, scanned.name, sheet.name, zone, firstRemovedRow);
        if (!hit.length) return;
        if (risks.length >= MAX_SHIFT_RISKS) { overflow += 1; return; }
        risks.push({ sheet: scanned.name, cell: `${columnLetters(scanned.columnIndex + c + 1)}${scanned.rowIndex + r + 1}`, formula: String(formula), references: hit.map((item) => item.text) });
      }));
    }

    const backup = lastWorkbookBackup();
    return {
      kind: "remove_duplicates" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      target: { ...target, sheetName: sheet.name },
      resolvedAddress: range.address.replace(/^.*!/, ""),
      hasHeaders,
      keyColumns,
      keyNames: keyColumns.map((index) => (hasHeaders && headers[index] ? `«${headers[index]}»` : letter(index))),
      dataRows: body.length,
      removed: duplicates.removed.map((item) => ({ row: firstDataRow + item.row, duplicateOf: firstDataRow + item.duplicateOf })),
      sampleRemoved: duplicates.removed.slice(0, SAMPLE_REMOVED).map((item) =>
        `строка ${firstDataRow + item.row} = строка ${firstDataRow + item.duplicateOf}: ${body[item.row].map((value) => (value === "" ? "∅" : String(value))).join(" · ")}`),
      expected: duplicates.expected,
      beforeFormulas: cloneMatrix(formulas),
      risks,
      riskOverflow: overflow,
      unscannedSheets: scan.unscanned,
      backup: backup ? { name: backup.name, at: backup.at } : null,
      undoAvailable: false as const,
      undoNote:
        "Отмены нет: удаление дубликатов сдвигает значения внутри области, и точно вернуть прежнее панель не берётся. " +
        "Перед подтверждением имеет смысл сделать резервную копию книги.",
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executeRemoveDuplicatesPlan(plan: RemoveDuplicatesPlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(plan.target.sheetId);
    sheet.load(["id", "name"]);
    const range = sheet.getRange(plan.resolvedAddress);
    range.load(["formulas", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
    await ctx.sync();
    const where = `${sheet.name}!${plan.resolvedAddress}`;
    if (JSON.stringify(range.formulas) !== JSON.stringify(plan.beforeFormulas)) {
      throw new ToolExecutionError(`Данные ${where} изменились после предпросмотра. Дубликаты не удалялись — сделайте новый предпросмотр.`, "failed_before_write");
    }
    const neighbours = await neighbourData(ctx, sheet, range);
    if (neighbours.length) {
      throw new ToolExecutionError(`Рядом с ${where} появились данные в столбцах ${neighbours.join(", ")} после предпросмотра. Дубликаты не удалялись.`, "failed_before_write");
    }

    let removedByExcel: number;
    try {
      const result = range.removeDuplicates(plan.keyColumns as number[], plan.hasHeaders);
      result.load(["removed", "uniqueRemaining"]);
      await ctx.sync();
      removedByExcel = result.removed;
    } catch (error: any) {
      throw new ToolExecutionError(`Не удалось определить итог удаления дубликатов в ${where}: ${error?.message ?? error}. Перечитайте область.`, "unknown");
    }

    range.load("values");
    await ctx.sync();
    const after = (range.values as unknown[][]).slice(plan.hasHeaders ? 1 : 0);
    const firstDataRow = range.rowIndex + (plan.hasHeaders ? 2 : 1);
    const wrong: string[] = [];
    plan.expected.forEach((row, index) => {
      if (JSON.stringify(after[index]) !== JSON.stringify(row)) wrong.push(`строка ${firstDataRow + index}`);
    });
    if (removedByExcel !== plan.removed.length || wrong.length) {
      throw new ToolExecutionError(
        `Excel удалил дубликатов: ${removedByExcel}, а по расчёту панели — ${plan.removed.length}` +
        (wrong.length ? `; расходятся ${wrong.length} строк (${wrong.slice(0, 5).join(", ")})` : "") +
        `. Проверьте ${where}; отмены нет${plan.backup ? `, есть резервная копия ${plan.backup.name}` : ""}.`,
        "applied"
      );
    }
    return {
      ok: true,
      executionState: "verified",
      sheet: sheet.name,
      address: plan.resolvedAddress,
      key: plan.keyNames,
      removedRows: plan.removed.length,
      remainingRows: plan.dataRows - plan.removed.length,
      removed: plan.sampleRemoved,
      checkedRows: plan.expected.length,
      note: "Остались первые вхождения, строки ниже поднялись; итог сверен с расчётом панели построчно. Регистр Excel не различает, пробел в конце — различает.",
      ...(plan.risks.length
        ? {
            affectedFormulas: plan.risks,
            ...(plan.riskOverflow ? { affectedFormulasOmitted: plan.riskOverflow } : {}),
            affectedFormulasNote: "Ссылки этих формул не подстроились: теперь они смотрят на данные других строк. Перечисли их пользователю."
          }
        : {}),
      ...(plan.unscannedSheets.length ? { unscannedSheets: plan.unscannedSheets, scanNote: "Эти листы слишком велики для обхода формул: про них ничего не проверено." } : {}),
      undoable: false,
      undoNote: plan.undoNote
    };
  });
}
