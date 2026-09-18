/**
 * Вторая партия оформления: закрепление областей, условное форматирование
 * и превращение области в таблицу Excel.
 *
 * Все три идут тем же путём, что и остальные изменения: подготовка ничего
 * не меняет и показывается человеку, исполнение сперва проверяет, что книга
 * не изменилась с предпросмотра, потом сверяет результат обратным чтением.
 * Отличие в том, что именно сверяется — см. `sheetRules.ts`.
 */

import { intersects, parseA1Rect } from "./a1";
import {
  assertTargetWritable,
  checkAddress,
  deepFreeze,
  MAX_IO_CELLS,
  preflightToolArgs,
  probeMergedAreas,
  rangeOf,
  readTableRanges,
  ToolError,
  ToolExecutionError,
  type TableRange
} from "./excelTools";
import { columnLetters } from "./formulaFill";
import {
  CELL_VALUE_OPERATOR,
  checkTableName,
  checkTableStyle,
  describeConditionalRule,
  describeFreeze,
  headerProblems,
  officeRuleType,
  parseConditionalRequest,
  parseFreezeLocation,
  parseFreezeRequest,
  ruleFormula,
  ruleMatches,
  sameFreeze,
  type ComparisonRule,
  type ConditionalRequest,
  type FreezeState
} from "./sheetRules";
import {
  action,
  getStructuralRevision,
  invalidateAfterStructuralChange,
  isCustomUndoAvailable,
  push
} from "./undo";
import { captureTarget, type WorkbookTarget } from "./workbookContext";

function planId(): string {
  return globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

function withoutSheet(address: string): string {
  return address.slice(address.lastIndexOf("!") + 1);
}

function toolError(error: any): ToolError {
  return error instanceof ToolError ? error : new ToolError(error?.message ?? String(error));
}

/* =========================================================================
 * Закрепление областей
 * ========================================================================= */

export interface FreezePanesPlan {
  readonly kind: "freeze_panes";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly before: FreezeState;
  readonly expected: FreezeState;
  readonly beforeText: string;
  readonly expectedText: string;
  readonly undoAvailable: boolean;
  readonly undoNote?: string;
  readonly createdAt: string;
}

async function readFreeze(ctx: Excel.RequestContext, sheet: Excel.Worksheet): Promise<FreezeState | null> {
  const location = sheet.freezePanes.getLocationOrNullObject();
  location.load(["isNullObject", "address"]);
  await ctx.sync();
  return parseFreezeLocation(location.isNullObject ? null : location.address);
}

/** Применяет закрепление с чистого листа: прежнее снимается, чтобы не смешаться. */
function applyFreeze(sheet: Excel.Worksheet, state: FreezeState) {
  sheet.freezePanes.unfreeze();
  if (state.rows && state.columns) sheet.freezePanes.freezeAt(sheet.getRangeByIndexes(0, 0, state.rows, state.columns));
  else if (state.rows) sheet.freezePanes.freezeRows(state.rows);
  else if (state.columns) sheet.freezePanes.freezeColumns(state.columns);
}

export async function prepareFreezePanesPlan(args: unknown): Promise<FreezePanesPlan> {
  preflightToolArgs("freeze_panes", args);
  const a = args as { sheet?: string } & Record<string, unknown>;
  let expected: FreezeState;
  try { expected = parseFreezeRequest(a); } catch (error) { throw toolError(error); }
  const target = await captureTarget(a.sheet);

  const prepared = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    sheet.load(["id", "name"]);
    const before = await readFreeze(ctx, sheet);
    if (!before) throw new ToolError("Не удалось прочитать текущее закрепление листа: Excel вернул непонятный адрес. Операция не выполнялась.");
    if (sameFreeze(before, expected)) {
      throw new ToolError(`На листе ${sheet.name} уже закреплено ровно так: ${describeFreeze(before)}. Менять нечего.`);
    }
    const undo = isCustomUndoAvailable();
    return {
      kind: "freeze_panes" as const,
      id: planId(),
      target: { ...target, sheetName: sheet.name },
      before,
      expected,
      beforeText: describeFreeze(before),
      expectedText: describeFreeze(expected),
      undoAvailable: undo,
      ...(undo ? {} : { undoNote: "Отмена недоступна: монитор изменений Excel не активен." }),
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executeFreezePanesPlan(plan: FreezePanesPlan) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(plan.target.sheetId);
    sheet.load(["id", "name"]);
    const current = await readFreeze(ctx, sheet);
    if (!sameFreeze(current, plan.before)) {
      throw new ToolExecutionError(
        `Закрепление на листе ${sheet.name} изменилось после предпросмотра: сейчас ${describeFreeze(current)}. Операция не выполнялась — сделайте новый предпросмотр.`,
        "failed_before_write"
      );
    }

    try {
      applyFreeze(sheet, plan.expected);
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(
        `Не удалось определить итог закрепления на листе ${sheet.name}: ${error?.message ?? error}. Проверьте вид листа.`,
        "unknown"
      );
    }

    const after = await readFreeze(ctx, sheet);
    if (!sameFreeze(after, plan.expected)) {
      throw new ToolExecutionError(
        sameFreeze(after, plan.before)
          ? `Закрепление на листе ${sheet.name} не изменилось: по-прежнему ${describeFreeze(after)}. Повтор ничего не даст.`
          : `Закрепление на листе ${sheet.name} применено, но стало ${describeFreeze(after)} вместо ${plan.expectedText}.`,
        "applied"
      );
    }

    let undoRecorded = false;
    if (plan.undoAvailable) {
      const sheetId = plan.target.sheetId;
      const before = plan.before;
      const expected = plan.expected;
      undoRecorded = push(action(`закрепление на листе ${sheet.name}`, async () => {
        const revision = getStructuralRevision();
        await Excel.run(async (undoCtx) => {
          const undoSheet = undoCtx.workbook.worksheets.getItem(sheetId);
          const now = await readFreeze(undoCtx, undoSheet);
          if (!sameFreeze(now, expected)) {
            throw new Error(`Закрепление изменено после операции агента: сейчас ${describeFreeze(now)}. Отмена остановлена, чтобы не затереть более свежую настройку.`);
          }
          if (getStructuralRevision() !== revision) throw new Error("Структура книги изменилась во время отмены. Отмена остановлена.");
          applyFreeze(undoSheet, before);
          await undoCtx.sync();
        });
      }));
    }

    return {
      ok: true,
      executionState: "verified",
      sheet: sheet.name,
      before: plan.beforeText,
      after: describeFreeze(after),
      frozen: after,
      note: "Закрепление — настройка вида листа: данные и оформление ячеек не менялись.",
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: plan.undoNote ?? "Автоматическая отмена этой операции недоступна." })
    };
  });
}

/* =========================================================================
 * Условное форматирование
 * ========================================================================= */

const PREDICTION_SAMPLE = 10;

export interface ConditionalFormatPlan {
  readonly kind: "add_conditional_format";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly address: string;
  readonly resolvedAddress: string;
  readonly cellCount: number;
  readonly request: ConditionalRequest;
  readonly ruleText: string;
  /** Правила, уже задевающие область: новое добавится к ним, а не заменит. */
  readonly existingRules: readonly { id: string; type: string }[];
  readonly existingNote?: string;
  /** Оценка панели: какие ячейки правило подсветит. */
  readonly prediction?: { matches: number; total: number; sample: readonly string[]; note: string };
  readonly undoAvailable: boolean;
  readonly undoNote?: string;
  readonly createdAt: string;
}

async function readRules(ctx: Excel.RequestContext, range: Excel.Range) {
  const collection = range.conditionalFormats;
  collection.load("items/id,items/type");
  await ctx.sync();
  return collection.items.map((item) => ({ id: String(item.id), type: String(item.type) }));
}

export async function prepareConditionalFormatPlan(args: unknown): Promise<ConditionalFormatPlan> {
  preflightToolArgs("add_conditional_format", args);
  const a = args as { sheet?: string; address: string } & Record<string, unknown>;
  let request: ConditionalRequest;
  try { request = parseConditionalRequest(a); } catch (error) { throw toolError(error); }
  const address = checkAddress(a.address);
  const target = await captureTarget(a.sheet);

  const prepared = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    const range = await rangeOf(ctx, sheet, address);
    range.load(["address", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
    sheet.load(["id", "name"]);
    try {
      range.format?.protection?.load("locked");
      sheet.protection?.load(["protected", "options"]);
    } catch { /* среда без сведений о защите */ }
    await ctx.sync();
    assertTargetWritable(sheet, range, "format");

    const cells = range.rowCount * range.columnCount;
    const existingRules = await readRules(ctx, range);

    // Оценка совпадений: только там, где значения можно прочитать разом.
    let prediction: ConditionalFormatPlan["prediction"];
    if (cells <= MAX_IO_CELLS && request.rule !== "colorScale" && request.rule !== "dataBar") {
      range.load("values");
      await ctx.sync();
      const sample: string[] = [];
      let matches = 0;
      (range.values as unknown[][]).forEach((row, r) => row.forEach((value, c) => {
        if (ruleMatches(request, value) !== true) return;
        matches += 1;
        if (sample.length < PREDICTION_SAMPLE) sample.push(`${columnLetters(range.columnIndex + c + 1)}${range.rowIndex + r + 1}`);
      }));
      prediction = {
        matches,
        total: cells,
        sample,
        note: "Это оценка панели: Excel применяет правило сам и через API не сообщает, какие ячейки подсвечены."
      };
    }

    const undo = isCustomUndoAvailable();
    return {
      kind: "add_conditional_format" as const,
      id: planId(),
      target: { ...target, sheetName: sheet.name },
      address,
      resolvedAddress: withoutSheet(range.address),
      cellCount: cells,
      request,
      ruleText: describeConditionalRule(request),
      existingRules,
      ...(existingRules.length
        ? {
            existingNote: `На области уже есть правил: ${existingRules.length}. Новое добавится к ним и не заменит их; при конфликте Excel применит правило, добавленное позже.`
          }
        : {}),
      ...(prediction ? { prediction } : {}),
      undoAvailable: undo,
      ...(undo ? {} : { undoNote: "Отмена недоступна: монитор изменений Excel не активен." }),
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

function applyHighlight(format: any, request: ConditionalRequest) {
  const highlight = request.highlight ?? {};
  if (highlight.fillColor) format.fill.color = highlight.fillColor;
  if (highlight.fontColor) format.font.color = highlight.fontColor;
  if (typeof highlight.bold === "boolean") format.font.bold = highlight.bold;
}

/** Формула в правиле Excel может прийти со знаком равенства или без. */
function sameRuleFormula(actual: unknown, expected: string): boolean {
  const clean = (value: unknown) => String(value ?? "").trim().replace(/^=/, "").toLowerCase();
  return clean(actual) === clean(expected);
}

export async function executeConditionalFormatPlan(plan: ConditionalFormatPlan) {
  const { request } = plan;
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(plan.target.sheetId);
    sheet.load(["id", "name"]);
    const range = sheet.getRange(plan.resolvedAddress);
    const current = await readRules(ctx, range);
    const sameRules = current.length === plan.existingRules.length &&
      current.every((rule) => plan.existingRules.some((known) => known.id === rule.id));
    if (!sameRules) {
      throw new ToolExecutionError(
        `Правила условного форматирования на ${sheet.name}!${plan.resolvedAddress} изменились после предпросмотра. Операция не выполнялась — сделайте новый предпросмотр.`,
        "failed_before_write"
      );
    }

    let added: Excel.ConditionalFormat;
    try {
      added = range.conditionalFormats.add(officeRuleType(request.rule) as any);
      if (request.rule === "textContains") {
        applyHighlight(added.textComparison.format, request);
        added.textComparison.rule = { operator: "Contains", text: String(request.text) } as any;
      } else if (request.rule === "colorScale") {
        const scale = request.scale!;
        added.colorScale.criteria = {
          minimum: { formula: null, type: "LowestValue", color: scale.minColor },
          ...(scale.midColor ? { midpoint: { formula: "50", type: "Percentile", color: scale.midColor } } : {}),
          maximum: { formula: null, type: "HighestValue", color: scale.maxColor }
        } as any;
      } else if (request.rule === "dataBar") {
        added.dataBar.positiveFormat.fillColor = String(request.barColor);
      } else {
        applyHighlight(added.cellValue.format, request);
        added.cellValue.rule = {
          formula1: ruleFormula(request.value as number | string),
          ...(request.rule === "between" ? { formula2: ruleFormula(request.value2 as number) } : {}),
          operator: CELL_VALUE_OPERATOR[request.rule as ComparisonRule]
        } as any;
      }
      added.load(["id", "type"]);
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(
        `Не удалось определить итог добавления правила на ${sheet.name}!${plan.resolvedAddress}: ${error?.message ?? error}. Перечитайте правила области.`,
        "unknown"
      );
    }

    const newId = String(added.id);
    const after = await readRules(ctx, range);
    const created = after.find((rule) => rule.id === newId);
    if (!created || after.length !== plan.existingRules.length + 1) {
      throw new ToolExecutionError(
        `Правило на ${sheet.name}!${plan.resolvedAddress} добавлено, но набор правил после операции не такой, как ожидалось: было ${plan.existingRules.length}, стало ${after.length}. Перечитайте правила.`,
        "applied"
      );
    }
    if (created.type !== officeRuleType(request.rule)) {
      throw new ToolExecutionError(`Добавлено правило типа ${created.type} вместо ${officeRuleType(request.rule)}.`, "applied");
    }

    // Сверка содержимого правила: условие и подсветка, как их прочитал Excel.
    const mismatches: string[] = [];
    try {
      if (request.rule === "textContains") {
        added.textComparison.load("rule");
        added.textComparison.format.fill.load("color");
        await ctx.sync();
        if (String((added.textComparison.rule as any)?.text ?? "").toLowerCase() !== String(request.text).toLowerCase()) mismatches.push("текст условия");
      } else if (request.rule === "colorScale") {
        added.colorScale.load("criteria");
        await ctx.sync();
        const criteria = added.colorScale.criteria as any;
        if (String(criteria?.minimum?.color ?? "").toUpperCase() !== request.scale!.minColor) mismatches.push("цвет минимума");
        if (String(criteria?.maximum?.color ?? "").toUpperCase() !== request.scale!.maxColor) mismatches.push("цвет максимума");
      } else if (request.rule === "dataBar") {
        added.dataBar.positiveFormat.load("fillColor");
        await ctx.sync();
        if (String(added.dataBar.positiveFormat.fillColor ?? "").toUpperCase() !== request.barColor) mismatches.push("цвет полосы");
      } else {
        added.cellValue.load("rule");
        added.cellValue.format.fill.load("color");
        await ctx.sync();
        const rule = added.cellValue.rule as any;
        if (rule?.operator !== CELL_VALUE_OPERATOR[request.rule as ComparisonRule]) mismatches.push("оператор");
        if (!sameRuleFormula(rule?.formula1, ruleFormula(request.value as number | string))) mismatches.push("значение");
        if (request.rule === "between" && !sameRuleFormula(rule?.formula2, ruleFormula(request.value2 as number))) mismatches.push("верхняя граница");
        if (request.highlight?.fillColor && String(added.cellValue.format.fill.color ?? "").toUpperCase() !== request.highlight.fillColor) {
          mismatches.push("цвет заливки");
        }
      }
    } catch (error: any) {
      mismatches.push(`правило не прочиталось обратно: ${error?.message ?? error}`);
    }
    if (mismatches.length) {
      throw new ToolExecutionError(
        `Правило на ${sheet.name}!${plan.resolvedAddress} добавлено, но обратное чтение расходится с планом: ${mismatches.join(", ")}. Перечитайте правила области.`,
        "applied"
      );
    }

    let undoRecorded = false;
    if (plan.undoAvailable) {
      const sheetId = plan.target.sheetId;
      const address = plan.resolvedAddress;
      undoRecorded = push(action(`условное форматирование ${sheet.name}!${address}`, async () => {
        const revision = getStructuralRevision();
        await Excel.run(async (undoCtx) => {
          const undoRange = undoCtx.workbook.worksheets.getItem(sheetId).getRange(address);
          const rules = await readRules(undoCtx, undoRange);
          if (!rules.some((rule) => rule.id === newId)) {
            throw new Error("Добавленного правила уже нет: его удалили после операции агента. Отменять нечего.");
          }
          if (getStructuralRevision() !== revision) throw new Error("Структура книги изменилась во время отмены. Отмена остановлена.");
          undoRange.conditionalFormats.getItem(newId).delete();
          await undoCtx.sync();
        });
      }));
    }

    return {
      ok: true,
      executionState: "verified",
      sheet: sheet.name,
      address: plan.resolvedAddress,
      rule: plan.ruleText,
      ruleId: newId,
      rulesBefore: plan.existingRules.length,
      rulesAfter: after.length,
      ...(plan.prediction ? { prediction: plan.prediction } : {}),
      note: "Правило проверено обратным чтением. Какие ячейки оно подсветило, Excel через API не сообщает — в prediction оценка панели.",
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: plan.undoNote ?? "Автоматическая отмена этой операции недоступна." })
    };
  });
}

/* =========================================================================
 * Таблица Excel
 * ========================================================================= */

export const DEFAULT_TABLE_STYLE = "TableStyleMedium2";

export interface CreateTablePlan {
  readonly kind: "create_table";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly address: string;
  readonly resolvedAddress: string;
  readonly rows: number;
  readonly columns: number;
  readonly headers: readonly unknown[];
  readonly headerProblems: readonly string[];
  readonly style: string;
  readonly name?: string;
  /** Слепок формул области: по нему ловится ручная правка перед созданием. */
  readonly signature: string;
  readonly tablesBefore: readonly TableRange[];
  readonly autoFilterWarning?: string;
  readonly mergeWarning?: string;
  readonly behaviourNote: string;
  readonly undoAvailable: false;
  readonly undoNote: string;
  readonly createdAt: string;
}

export async function prepareCreateTablePlan(args: unknown): Promise<CreateTablePlan> {
  preflightToolArgs("create_table", args);
  const a = args as { sheet?: string; address: string; style?: string; name?: string };
  let style: string;
  let name: string | undefined;
  try {
    style = checkTableStyle(a.style ?? DEFAULT_TABLE_STYLE);
    name = a.name === undefined ? undefined : checkTableName(a.name);
  } catch (error) { throw toolError(error); }
  const address = checkAddress(a.address);
  const target = await captureTarget(a.sheet);

  const prepared = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    const range = await rangeOf(ctx, sheet, address);
    range.load(["address", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
    sheet.load(["id", "name"]);
    try { sheet.protection?.load("protected"); } catch { /* среда без сведений о защите */ }
    await ctx.sync();
    if (sheet.protection?.protected) {
      throw new ToolError(`Лист ${sheet.name} защищён: таблицу на нём создать нельзя. Операция не выполнялась.`);
    }
    const cells = range.rowCount * range.columnCount;
    if (cells > MAX_IO_CELLS) {
      throw new ToolError(`Таблица создаётся на области до ${MAX_IO_CELLS} ячеек; ${range.address} содержит ${cells}.`);
    }
    if (range.rowCount < 2) {
      throw new ToolError("В таблице нужна строка заголовков и хотя бы одна строка данных: в области одна строка.");
    }

    const rect = parseA1Rect(withoutSheet(range.address));
    const tables = await readTableRanges(ctx, sheet);
    const overlapping = tables.filter((table) => {
      const other = parseA1Rect(withoutSheet(table.address));
      return rect && other && intersects(rect, other);
    });
    if (overlapping.length) {
      throw new ToolError(
        `Область пересекается с таблицей ${overlapping.map((table) => `${table.name} (${table.address})`).join(", ")}. ` +
        "Таблицы Excel не могут пересекаться. Операция не выполнялась."
      );
    }
    if (name) {
      const existing = ctx.workbook.tables.getItemOrNullObject(name);
      existing.load("isNullObject");
      await ctx.sync();
      if (!existing.isNullObject) throw new ToolError(`Таблица с именем «${name}» в книге уже есть. Выберите другое имя.`);
    }

    const merged = await probeMergedAreas(ctx, sheet, range);
    if (merged.areas.length) {
      throw new ToolError(
        `В области есть объединённые ячейки: ${merged.areas.join(", ")}. Excel не создаёт таблицу поверх объединений. Операция не выполнялась.`
      );
    }

    range.load(["values", "formulas"]);
    let autoFilterEnabled = false;
    try {
      sheet.autoFilter.load("enabled");
      await ctx.sync();
      autoFilterEnabled = Boolean(sheet.autoFilter.enabled);
    } catch {
      await ctx.sync();
    }
    const values = range.values as unknown[][];
    const formulas = range.formulas as unknown[][];

    return {
      kind: "create_table" as const,
      id: planId(),
      target: { ...target, sheetName: sheet.name },
      address,
      resolvedAddress: withoutSheet(range.address),
      rows: range.rowCount,
      columns: range.columnCount,
      headers: [...values[0]],
      headerProblems: headerProblems(values[0], formulas[0]),
      style,
      ...(name ? { name } : {}),
      signature: JSON.stringify(formulas),
      tablesBefore: tables,
      ...(autoFilterEnabled
        ? { autoFilterWarning: "На листе стоит автофильтр. У таблицы свой фильтр, и фильтр листа Excel при этом снимет — скрытые им строки снова станут видны." }
        : {}),
      ...(merged.unresolvedAnchors.length
        ? { mergeWarning: `Рядом найдены углы объединений ${merged.unresolvedAnchors.join(", ")}, границы которых эта сборка Excel не сообщает. Если объединение заходит в область, Excel откажет в создании таблицы.` }
        : {}),
      behaviourNote:
        "Таблица меняет поведение области: запись вплотную к ней расширяет таблицу, формула в столбце протягивается на весь столбец, " +
        "у неё свой фильтр в шапке. Первая строка области станет заголовками.",
      undoAvailable: false as const,
      undoNote:
        "Отмены в панели нет. Вернуть обычную область можно в Excel: вкладка «Конструктор таблиц» → «Преобразовать в диапазон»; " +
        "оформление стиля при этом останется на ячейках.",
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executeCreateTablePlan(plan: CreateTablePlan) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(plan.target.sheetId);
    sheet.load(["id", "name"]);
    const range = sheet.getRange(plan.resolvedAddress);
    range.load("formulas");
    await ctx.sync();
    if (JSON.stringify(range.formulas) !== plan.signature) {
      throw new ToolExecutionError(
        `Данные ${sheet.name}!${plan.resolvedAddress} изменились после предпросмотра. Таблица не создавалась — сделайте новый предпросмотр.`,
        "failed_before_write"
      );
    }

    let table: Excel.Table;
    try {
      table = sheet.tables.add(plan.resolvedAddress, true);
      if (plan.name) table.name = plan.name;
      table.style = plan.style;
      table.load(["name", "style", "showHeaders"]);
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(
        `Excel отказал в создании таблицы на ${sheet.name}!${plan.resolvedAddress}: ${error?.message ?? error}. ` +
        "Неизвестно, успела ли таблица появиться — перечитайте лист.",
        "unknown"
      );
    }

    const tableRange = table.getRange();
    const header = table.getHeaderRowRange();
    tableRange.load("address");
    header.load("values");
    await ctx.sync();
    // Таблица меняет поведение записей в своей области: прежние снимки отмены
    // по этим ячейкам могли бы при возврате протянуть формулу на весь столбец.
    const invalidatedUndo = invalidateAfterStructuralChange();

    const problems: string[] = [];
    if (withoutSheet(String(tableRange.address)) !== plan.resolvedAddress) problems.push(`границы ${tableRange.address} вместо ${plan.resolvedAddress}`);
    if (table.style !== plan.style) problems.push(`стиль ${table.style} вместо ${plan.style}`);
    if (plan.name && table.name !== plan.name) problems.push(`имя ${table.name} вместо ${plan.name}`);
    if (!table.showHeaders) problems.push("строка заголовков не показана");
    if (problems.length) {
      throw new ToolExecutionError(
        `Таблица ${table.name} создана, но расходится с планом: ${problems.join("; ")}. Проверьте её в Excel.`,
        "applied"
      );
    }

    const headersAfter = (header.values as unknown[][])[0] ?? [];
    const renamed = headersAfter
      .map((value, index) => ({ column: index + 1, before: plan.headers[index], after: value }))
      .filter((item) => String(item.before ?? "") !== String(item.after ?? ""));

    return {
      ok: true,
      executionState: "verified",
      sheet: sheet.name,
      table: table.name,
      address: withoutSheet(String(tableRange.address)),
      style: table.style,
      headers: headersAfter,
      ...(renamed.length
        ? {
            renamedHeaders: renamed,
            renamedNote: "Excel изменил эти заголовки при создании таблицы: пустые получили имена, повторы — номера, формулы стали текстом. Назови это пользователю."
          }
        : {}),
      behaviourNote: plan.behaviourNote,
      undoable: false,
      undoNote: plan.undoNote,
      invalidatedUndo
    };
  });
}
