/**
 * Таблица Excel → обычный диапазон (этап 7, 7.4.1).
 *
 * Замер в Excel 24 сентября 2026 года (ExcelApi 1.2 `convertToRange`):
 * - оформление стиля остаётся на ячейках обычным оформлением: шапка
 *   TableStyleMedium2 стала заливкой #4F81BD с белым полужирным шрифтом,
 *   чётные строки — заливкой #DCE6F1;
 * - структурные ссылки Excel переписывает в абсолютные адреса: на этом листе,
 *   на других листах и в именах (`=SUM(ТПроба[Сумма])` → `=SUM(Лист!$B$2:$B$5)`,
 *   строка итогов `SUBTOTAL(109,[Сумма])` → `SUBTOTAL(109,Лист!$B$2:$B$5)`);
 *   такие ссылки больше не растут вместе с данными;
 * - имя таблицы внутри текста — `INDIRECT("ТПроба[Сумма]")` — не переписывается
 *   и становится `#ССЫЛКА!`;
 * - фильтр таблицы снимается: скрытые им строки снова видны, и SUBTOTAL
 *   в строке итогов начинает считать все строки.
 *
 * Отмены нет: новая таблица на том же месте не вернёт структурных ссылок.
 */

import { parseA1Rect, type A1Rect } from "./a1";
import {
  assertPlanWorkbook,
  checkAddress,
  deepFreeze,
  MAX_IO_CELLS,
  preflightToolArgs,
  rangeOf,
  scanWorkbookFormulas,
  ToolError,
  ToolExecutionError
} from "./excelTools";
import { columnLetters } from "./formulaFill";
import type { CellMention, ScannedSheetLike } from "./sheetImpact";
import { invalidateAfterStructuralChange } from "./undo";
import { captureTarget, type WorkbookTarget } from "./workbookContext";

const MAX_LISTED = 20;
const MAX_ROW_SCAN = 2_000;

const withoutStrings = (formula: string) => formula.replace(/"(?:[^"]|"")*"/g, '""');
const escape = (text: string) => text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

/** Формула ссылается на таблицу структурной ссылкой: `ТПроба[Сумма]`, `ТПроба[#All]`. */
export function usesTable(formula: unknown, table: string): boolean {
  if (typeof formula !== "string" || !formula.startsWith("=")) return false;
  return new RegExp(`(?<![A-Za-z0-9_.\\u0400-\\u04FF])${escape(table)}\\[`, "i").test(withoutStrings(formula));
}

/** Внутри таблицы имя можно опустить: `=[@Сумма]*2`, `SUBTOTAL(109,[Сумма])`. */
export function usesOwnColumns(formula: unknown): boolean {
  if (typeof formula !== "string" || !formula.startsWith("=")) return false;
  // Внешняя книга тоже пишется в скобках — `[Книга.xlsx]Лист!A1` — её не считаем.
  return /(?<!['A-Za-z0-9_.Ѐ-ӿ])\[[^\]]*\](?![^\s!()+\-*\/,;&=<>^]*!)/.test(withoutStrings(formula));
}

/** Имя таблицы внутри текста формулы — Excel его не перепишет. */
export function literalMentionsTable(formula: unknown, table: string): boolean {
  if (typeof formula !== "string" || !formula.startsWith("=")) return false;
  const wanted = table.toLowerCase();
  return (formula.match(/"(?:[^"]|"")*"/g) ?? []).some((literal) => literal.toLowerCase().includes(wanted));
}

export interface TableImpact {
  /** Формулы со структурными ссылками: Excel перепишет их в обычные адреса. */
  structured: CellMention[];
  /** Формулы с именем таблицы в тексте: станут #ССЫЛКА!. */
  literal: CellMention[];
  overflow: number;
}

export function tableImpact(
  sheets: readonly ScannedSheetLike[],
  table: string,
  tableSheet: string,
  tableRect: A1Rect
): TableImpact {
  const impact: TableImpact = { structured: [], literal: [], overflow: 0 };
  const add = (list: CellMention[], item: CellMention) => {
    if (list.length < MAX_LISTED) list.push(item);
    else impact.overflow += 1;
  };
  for (const sheet of sheets) {
    const own = sheet.name.toLowerCase() === tableSheet.toLowerCase();
    sheet.formulas.forEach((row, r) => row.forEach((formula, c) => {
      if (typeof formula !== "string" || !formula.startsWith("=")) return;
      const rowNumber = sheet.rowIndex + r + 1;
      const columnNumber = sheet.columnIndex + c + 1;
      const inside = own &&
        rowNumber >= tableRect.rowStart && rowNumber <= tableRect.rowEnd &&
        columnNumber >= tableRect.columnStart && columnNumber <= tableRect.columnEnd;
      const item = { sheet: sheet.name, cell: `${columnLetters(columnNumber)}${rowNumber}`, formula };
      if (usesTable(formula, table) || (inside && usesOwnColumns(formula))) add(impact.structured, item);
      if (literalMentionsTable(formula, table)) add(impact.literal, item);
    }));
  }
  return impact;
}

/** Значение зависит от фильтра: SUBTOTAL и AGGREGATE считают только видимые строки. */
const filterSensitive = (formula: string) => /\b(SUBTOTAL|AGGREGATE)\s*\(/i.test(formula);

/* --- план ------------------------------------------------------------------------ */

export interface ConvertTablePlan {
  readonly kind: "convert_table_to_range";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly tableId: string;
  readonly tableName: string;
  readonly resolvedAddress: string;
  readonly style: string;
  readonly rows: number;
  readonly columns: number;
  readonly showTotals: boolean;
  readonly filtered: boolean;
  readonly hiddenRows: number;
  readonly structured: readonly CellMention[];
  readonly literal: readonly CellMention[];
  readonly overflow: number;
  readonly names: readonly { name: string; formula: string }[];
  readonly unscannedSheets: readonly string[];
  readonly signature: string;
  readonly styleNote: string;
  readonly warnings: readonly string[];
  readonly undoAvailable: false;
  readonly undoNote: string;
  readonly createdAt: string;
}

/** Таблица по имени или по ячейке внутри неё. */
async function findTable(ctx: Excel.RequestContext, a: { table?: string; sheet?: string; address?: string }, target: WorkbookTarget) {
  if (a.table) {
    const table = ctx.workbook.tables.getItemOrNullObject(a.table);
    table.load(["isNullObject", "id", "name"]);
    await ctx.sync();
    if (!table.isNullObject) return table;
    const all = ctx.workbook.tables;
    all.load("items/name");
    await ctx.sync();
    const names = all.items.map((item) => item.name);
    throw new ToolError(`Таблицы «${a.table}» в книге нет. ${names.length ? `Есть: ${names.join(", ")}.` : "Таблиц в книге нет."}`);
  }
  if (!a.address) throw new ToolError("Укажите table — имя таблицы — или address ячейки внутри неё.");
  const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
  const range = await rangeOf(ctx, sheet, checkAddress(a.address));
  range.load("address");
  const tables = sheet.tables;
  tables.load("items/name,items/id");
  await ctx.sync();
  const rect = parseA1Rect(String(range.address).replace(/^.*!/, ""));
  const ranges = tables.items.map((table) => { const r = table.getRange(); r.load("address"); return { table, r }; });
  await ctx.sync();
  const hit = ranges.filter(({ r }) => {
    const other = parseA1Rect(String(r.address).replace(/^.*!/, ""));
    return rect && other && rect.rowStart <= other.rowEnd && other.rowStart <= rect.rowEnd && rect.columnStart <= other.columnEnd && other.columnStart <= rect.columnEnd;
  });
  if (!hit.length) {
    throw new ToolError(`В ${range.address} нет таблицы Excel. ${tables.items.length ? `На листе есть: ${tables.items.map((item) => item.name).join(", ")}.` : "На этом листе таблиц нет."}`);
  }
  if (hit.length > 1) throw new ToolError(`Область задевает несколько таблиц: ${hit.map(({ table }) => table.name).join(", ")}. Назовите одну в table.`);
  return hit[0].table;
}

export async function prepareConvertTablePlan(args: unknown): Promise<ConvertTablePlan> {
  preflightToolArgs("convert_table_to_range", args);
  const a = args as { table?: string; sheet?: string; address?: string };
  const target = await captureTarget(a.sheet);
  const prepared = await Excel.run(async (ctx) => {
    const table = await findTable(ctx, a, target);
    const sheet = table.worksheet;
    sheet.load(["id", "name"]);
    try { sheet.protection?.load("protected"); } catch { /* среда без сведений о защите */ }
    const range = table.getRange();
    range.load(["address", "rowCount", "columnCount"]);
    table.load(["style", "showTotals", "showHeaders"]);
    await ctx.sync();
    if (sheet.protection?.protected) throw new ToolError(`Лист ${sheet.name} защищён: таблицу на нём преобразовать нельзя. Операция не выполнялась.`);
    const cells = range.rowCount * range.columnCount;
    if (cells > MAX_IO_CELLS) throw new ToolError(`Преобразуется таблица до ${MAX_IO_CELLS} ячеек; ${table.name} содержит ${cells}.`);
    range.load("formulas");

    let filtered = false;
    try {
      table.autoFilter.load("isDataFiltered");
      await ctx.sync();
      filtered = Boolean(table.autoFilter.isDataFiltered);
    } catch {
      await ctx.sync();
    }
    let hiddenRows = 0;
    if (range.rowCount <= MAX_ROW_SCAN) {
      const rows = Array.from({ length: range.rowCount }, (_, index) => { const row = range.getRow(index); row.load("rowHidden"); return row; });
      await ctx.sync();
      hiddenRows = rows.filter((row) => row.rowHidden).length;
    }

    const resolvedAddress = String(range.address).replace(/^.*!/, "");
    const rect = parseA1Rect(resolvedAddress)!;
    const scan = await scanWorkbookFormulas(ctx);
    const impact = tableImpact(scan.sheets, table.name, sheet.name, rect);

    const names: { name: string; formula: string }[] = [];
    try {
      const collection = ctx.workbook.names;
      collection.load("items/name,items/formula");
      await ctx.sync();
      for (const item of collection.items) if (usesTable(item.formula, table.name)) names.push({ name: item.name, formula: String(item.formula) });
    } catch { /* среда без списка имён */ }

    const warnings: string[] = [];
    if (filtered) {
      warnings.push(
        `У таблицы включён фильтр${hiddenRows ? `, скрыто строк: ${hiddenRows}` : ""}. Excel снимет его: скрытые строки снова станут видны` +
        (table.showTotals ? ", а итоги SUBTOTAL начнут считать все строки." : ".")
      );
    }
    if (impact.structured.length || names.length) {
      warnings.push(
        `Ссылки на таблицу (${impact.structured.length + (impact.overflow ? `+${impact.overflow}` : "") } формул${names.length ? `, имён: ${names.length}` : ""}) Excel перепишет в обычные адреса. ` +
        "Значения останутся прежними, но при добавлении строк такие ссылки расти больше не будут."
      );
    }
    if (impact.literal.length) {
      warnings.push(`Имя таблицы стоит внутри текста формулы (${impact.literal.map((item) => `${item.sheet}!${item.cell}`).join(", ")}) — такие формулы станут #ССЫЛКА!.`);
    }
    if (table.showTotals) warnings.push("Строка итогов останется обычной строкой с формулами.");
    if (scan.unscanned.length) warnings.push(`Листы ${scan.unscanned.join(", ")} слишком большие для разбора: ссылки на таблицу там не проверены.`);

    return {
      kind: "convert_table_to_range" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      target: { ...target, sheetId: sheet.id, sheetName: sheet.name },
      tableId: table.id,
      tableName: table.name,
      resolvedAddress,
      style: String(table.style ?? ""),
      rows: range.rowCount,
      columns: range.columnCount,
      showTotals: Boolean(table.showTotals),
      filtered,
      hiddenRows,
      structured: impact.structured,
      literal: impact.literal,
      overflow: impact.overflow,
      names,
      unscannedSheets: scan.unscanned,
      signature: JSON.stringify(range.formulas),
      styleNote: table.style
        ? `Оформление стиля ${table.style} — заливка шапки и полос, шрифт — останется на ячейках обычным оформлением: Excel его не снимает. Если оно не нужно, его снимают отдельно.`
        : "У таблицы нет стиля: оформление ячеек не изменится.",
      warnings,
      undoAvailable: false as const,
      undoNote:
        "Отмены в панели нет: Excel переписывает ссылки на таблицу в обычные адреса, и новая таблица на том же месте их не вернёт. " +
        "Таблицу можно создать заново через create_table, но ссылки останутся обычными.",
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executeConvertTablePlan(plan: ConvertTablePlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const table = ctx.workbook.tables.getItemOrNullObject(plan.tableName);
    table.load(["isNullObject", "id"]);
    await ctx.sync();
    if (table.isNullObject || table.id !== plan.tableId) {
      throw new ToolExecutionError(`Таблицы «${plan.tableName}» больше нет или это другая таблица. Операция не выполнялась — перечитайте лист.`, "failed_before_write");
    }
    const sheet = ctx.workbook.worksheets.getItem(plan.target.sheetId);
    sheet.load("name");
    const tableRange = table.getRange();
    tableRange.load(["address", "formulas", "values"]);
    const mentioned = [...plan.structured, ...plan.literal].filter((item, index, all) =>
      all.findIndex((other) => other.sheet === item.sheet && other.cell === item.cell) === index);
    const cells = mentioned.map((item) => {
      const cell = ctx.workbook.worksheets.getItem(item.sheet).getRange(item.cell);
      cell.load(["values", "formulas"]);
      return { item, cell };
    });
    await ctx.sync();
    if (String(tableRange.address).replace(/^.*!/, "") !== plan.resolvedAddress || JSON.stringify(tableRange.formulas) !== plan.signature) {
      throw new ToolExecutionError(`Таблица «${plan.tableName}» изменилась после предпросмотра. Операция не выполнялась — сделайте новый предпросмотр.`, "failed_before_write");
    }
    const valuesBefore = tableRange.values as unknown[][];
    const mentionedBefore = cells.map(({ cell }) => (cell.values as unknown[][])[0][0]);

    try {
      table.convertToRange();
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(`Excel отказал в преобразовании таблицы «${plan.tableName}»: ${error?.message ?? error}. Перечитайте лист.`, "unknown");
    }
    // Запись вплотную к области больше не расширяет таблицу: прежние снимки
    // отмены по этим ячейкам сделаны при другом поведении.
    const invalidatedUndo = invalidateAfterStructuralChange();

    const gone = ctx.workbook.tables.getItemOrNullObject(plan.tableName);
    gone.load("isNullObject");
    const after = sheet.getRange(plan.resolvedAddress);
    after.load(["values", "formulas"]);
    for (const { cell } of cells) cell.load(["values", "formulas"]);
    await ctx.sync();

    const where = `${sheet.name}!${plan.resolvedAddress}`;
    const problems: string[] = [];
    if (!gone.isNullObject) problems.push(`таблица «${plan.tableName}» по-прежнему в книге`);

    // Данные: константы обязаны остаться, формулы — дать то же значение,
    // кроме SUBTOTAL/AGGREGATE при снятом фильтре: там перемена ожидаема.
    const formulasAfter = after.formulas as unknown[][];
    const valuesAfter = after.values as unknown[][];
    const origin = parseA1Rect(plan.resolvedAddress)!;
    const changed: { cell: string; before: unknown; after: unknown; expected: boolean }[] = [];
    valuesBefore.forEach((row, r) => row.forEach((value, c) => {
      if (valuesAfter[r]?.[c] === value) return;
      const formula = String(formulasAfter[r]?.[c] ?? "");
      changed.push({ cell: `${columnLetters(origin.columnStart + c)}${origin.rowStart + r}`, before: value, after: valuesAfter[r]?.[c], expected: plan.filtered && filterSensitive(formula) });
    }));

    const literalKeys = new Set(plan.literal.map((item) => `${item.sheet}!${item.cell}`));
    const stillStructured: string[] = [];
    const brokenLiteral: string[] = [];
    cells.forEach(({ item, cell }, index) => {
      const key = `${item.sheet}!${item.cell}`;
      const formula = String((cell.formulas as unknown[][])[0][0] ?? "");
      const value = (cell.values as unknown[][])[0][0];
      if (usesTable(formula, plan.tableName)) stillStructured.push(key);
      if (literalKeys.has(key)) { if (value !== mentionedBefore[index]) brokenLiteral.push(key); return; }
      if (value !== mentionedBefore[index] && !(item.sheet === sheet.name && changed.some((entry) => entry.cell === item.cell))) {
        changed.push({ cell: key, before: mentionedBefore[index], after: value, expected: plan.filtered && filterSensitive(formula) });
      }
    });
    if (stillStructured.length) problems.push(`структурные ссылки остались в ${stillStructured.join(", ")}`);
    const unexpected = changed.filter((entry) => !entry.expected);
    if (unexpected.length) problems.push(`изменились значения ${unexpected.slice(0, 8).map((entry) => `${entry.cell}: ${String(entry.before)} → ${String(entry.after)}`).join(", ")}`);

    const result = {
      ok: true,
      executionState: "verified",
      convertedTable: plan.tableName,
      address: where,
      styleNote: plan.styleNote,
      rewrittenReferences: plan.structured.length + plan.overflow + plan.names.length,
      ...(brokenLiteral.length || plan.literal.length
        ? { brokenFormulas: plan.literal.map((item) => `${item.sheet}!${item.cell}`), brokenNote: "В этих формулах имя таблицы стояло внутри текста; Excel его не переписал, и они показывают #ССЫЛКА!. Назови их пользователю." }
        : {}),
      ...(changed.some((entry) => entry.expected)
        ? { filterNote: `Фильтр таблицы снят: скрытые строки видны, итоги пересчитались по всем строкам (${changed.filter((entry) => entry.expected).map((entry) => `${entry.cell}: ${String(entry.before)} → ${String(entry.after)}`).join(", ")}).` }
        : plan.filtered ? { filterNote: "Фильтр таблицы снят: скрытые им строки снова видны." } : {}),
      ...(invalidatedUndo ? { undoHistoryCleared: invalidatedUndo } : {}),
      undoable: false,
      undoNote: plan.undoNote
    };
    if (problems.length) {
      throw new ToolExecutionError(`Таблица «${plan.tableName}» преобразована, но сверка расходится: ${problems.join("; ")}. Проверьте ${where}; отмены нет.`, "applied");
    }
    return result;
  });
}
