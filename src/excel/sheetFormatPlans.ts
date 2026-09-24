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
import type { FunctionCheck } from "./functionProbe";
import {
  assertPlanWorkbook,
  assertTargetWritable,
  checkAddress,
  deepFreeze,
  MAX_IO_CELLS,
  planFunctionCheck,
  preflightToolArgs,
  probeMergedAreas,
  rangeOf,
  readTableRanges,
  runFunctionCheck,
  ToolError,
  ToolExecutionError,
  type TableRange
} from "./excelTools";
import { columnLetters } from "./formulaFill";
import {
  CELL_VALUE_OPERATOR,
  conditionalRuleMismatches,
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
  type FreezeState,
  type RuleSnapshot
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
  assertPlanWorkbook(plan);
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
  /**
   * Прежние правила целиком — тип, приоритет, область, условие, оформление.
   * По нему перед исполнением ловится и правка содержимого с тем же ID.
   */
  readonly existingSignature: string;
  readonly existingNote?: string;
  /** Оценка панели: какие ячейки правило подсветит. */
  readonly prediction?: { matches: number; total: number; sample: readonly string[]; note: string };
  /** Куда встанет новое правило среди правил области: first — выше всех (как в Excel), last — ниже. */
  readonly order: "first" | "last";
  /** Функции формулы условия: Excel примет и неизвестную, а правило тихо не сработает. */
  readonly functionCheck?: FunctionCheck;
  readonly undoAvailable: boolean;
  readonly undoNote?: string;
  readonly createdAt: string;
}

/**
 * Правило без номера и приоритета — то, по чему его узнают.
 *
 * Замер 24 сентября 2026 года: ID правила у Excel — это номер в списке
 * правил области («0», «1», …), а приоритет — место в том же списке.
 * Правило, добавленное выше, сдвигает оба. Узнавать правило по ID нельзя:
 * отмена так удаляла чужое правило, оказавшееся на прежнем месте.
 */
export function ruleKey(rule: RuleSnapshot): string {
  const { id: _id, priority: _priority, ...content } = rule;
  return JSON.stringify(content);
}

/**
 * Правила области со всем, что панель умеет прочитать (план стабилизации,
 * S3.2). Формат ответа снят с Excel 24 сентября 2026 года: цвета заглавными,
 * незаданные цвет текста и жирность — null, формула условия со знаком «=».
 *
 * Свойства читаются по типу правила: чужое подсвойство Office.js грузить
 * отказывается. Если чтение содержимого сорвалось или тип панели незнаком,
 * правило помечается contentUnread — тогда о нём известны только тип,
 * приоритет и область, и это говорится, а не выдаётся за полную сверку.
 */
export async function readRuleSnapshots(ctx: Excel.RequestContext, range: Excel.Range): Promise<RuleSnapshot[]> {
  const collection = range.conditionalFormats;
  collection.load("items/id,items/type,items/priority");
  await ctx.sync();
  const pending = collection.items.map((item: any) => {
    const type = String(item.type);
    const where = item.getRangeOrNullObject();
    where.load(["isNullObject", "address"]);
    const highlight = type === "CellValue" ? item.cellValue : type === "ContainsText" ? item.textComparison : null;
    const custom = type === "Custom" ? item.custom : null;
    return { item, type, where, highlight, custom };
  });
  await ctx.sync();

  let contentRead = true;
  try {
    for (const { item, type, highlight, custom } of pending) {
      if (highlight) {
        highlight.load("rule");
        highlight.format.fill.load("color");
        highlight.format.font.load(["color", "bold"]);
      } else if (custom) {
        custom.rule.load("formula");
        custom.format.fill.load("color");
        custom.format.font.load(["color", "bold"]);
      } else if (type === "ColorScale") {
        item.colorScale.load("criteria");
      } else if (type === "DataBar") {
        item.dataBar.positiveFormat.load("fillColor");
      }
    }
    await ctx.sync();
  } catch {
    contentRead = false;
  }

  return pending.map(({ item, type, where, highlight, custom }) => {
    const base: RuleSnapshot = {
      id: String(item.id),
      type,
      priority: typeof item.priority === "number" ? item.priority : null,
      range: where.isNullObject ? null : String(where.address)
    };
    if (!contentRead) return { ...base, contentUnread: true };
    if (highlight) {
      const rule = highlight.rule ? { ...highlight.rule } : null;
      if (rule) delete (rule as any)["@odata.type"];
      return { ...base, rule, fill: highlight.format.fill.color ?? null, fontColor: highlight.format.font.color ?? null, bold: highlight.format.font.bold ?? null };
    }
    if (custom) {
      return {
        ...base,
        rule: { formula: custom.rule.formula ?? null },
        fill: custom.format.fill.color ?? null,
        fontColor: custom.format.font.color ?? null,
        bold: custom.format.font.bold ?? null
      };
    }
    if (type === "ColorScale") {
      const criteria = item.colorScale.criteria as any;
      const point = (value: any) => (value ? { color: value.color ?? null, type: value.type ?? null, formula: value.formula ?? null } : null);
      return { ...base, criteria: criteria ? { minimum: point(criteria.minimum), midpoint: point(criteria.midpoint), maximum: point(criteria.maximum) } : null };
    }
    if (type === "DataBar") return { ...base, barColor: item.dataBar.positiveFormat.fillColor ?? null };
    return { ...base, contentUnread: true };
  });
}

export async function prepareConditionalFormatPlan(args: unknown): Promise<ConditionalFormatPlan> {
  preflightToolArgs("add_conditional_format", args);
  const a = args as { sheet?: string; address: string; order?: string } & Record<string, unknown>;
  let request: ConditionalRequest;
  try { request = parseConditionalRequest(a); } catch (error) { throw toolError(error); }
  const order: "first" | "last" = a.order === "last" ? "last" : "first";
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
    const existing = await readRuleSnapshots(ctx, range);
    const existingRules = existing.map((rule) => ({ id: rule.id, type: rule.type }));
    const unread = existing.filter((rule) => rule.contentUnread).length;
    // Замер: неизвестную функцию в условии Excel принимает, и правило молча
    // не срабатывает никогда. Функции проверяются так же, как при записи формул.
    const planned = request.rule === "formula" ? await planFunctionCheck(ctx, sheet, range, [request.formula]) : undefined;
    const functionCheck = planned && !planned.probeCell
      ? { ...planned, note: planned.note.replace(/Если какой-то.*$/, "Если какой-то из них нет в этом Excel, правило молча не подсветит ничего.") }
      : planned;

    // Оценка совпадений: только там, где значения можно прочитать разом.
    let prediction: ConditionalFormatPlan["prediction"];
    if (cells <= MAX_IO_CELLS && ruleMatches(request, 0) !== null) {
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
      existingSignature: JSON.stringify(existing),
      ...(existingRules.length
        ? {
            existingNote:
              `На области уже есть правил: ${existingRules.length}. Новое добавится к ним и не заменит их. ` +
              (order === "first"
                ? "Оно встанет выше них: где правила задают одно и то же, например заливку, будет видно новое."
                : "Оно встанет ниже них: где правила задают одно и то же, например заливку, нового видно не будет.") +
              (unread
                ? ` Содержимое ${unread} из них панель не читает (тип или сбой чтения): если его изменят до подтверждения, это заметно не будет — только по типу, приоритету и области.`
                : "")
          }
        : {}),
      ...(prediction ? { prediction } : {}),
      order,
      ...(functionCheck ? { functionCheck } : {}),
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

/** Прежние правила по содержимому, в прежнем порядке. */
const keysOf = (rules: readonly RuleSnapshot[]) => rules.map(ruleKey);

/**
 * Убрать правила, которых не было до операции, и проверить, что осталось
 * прежнее. true — правила области снова как до операции.
 */
async function removeLeftovers(ctx: Excel.RequestContext, range: Excel.Range, before: readonly RuleSnapshot[]): Promise<boolean> {
  try {
    const now = await readRuleSnapshots(ctx, range);
    const wanted = keysOf(before);
    const extra = now.filter((rule) => {
      const index = wanted.indexOf(ruleKey(rule));
      if (index === -1) return true;
      wanted.splice(index, 1);
      return false;
    });
    // Удаление с конца: номера правил выше удаляемого не сдвигаются.
    for (const rule of [...extra].reverse()) range.conditionalFormats.getItem(rule.id).delete();
    if (extra.length) await ctx.sync();
    const left = await readRuleSnapshots(ctx, range);
    return JSON.stringify(keysOf(left)) === JSON.stringify(keysOf(before));
  } catch {
    return false;
  }
}

export async function executeConditionalFormatPlan(plan: ConditionalFormatPlan) {
  assertPlanWorkbook(plan);
  const { request } = plan;
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(plan.target.sheetId);
    sheet.load(["id", "name"]);
    const range = sheet.getRange(plan.resolvedAddress);
    // Сверяется содержимое, а не только число и ID: правило, которому
    // поменяли условие или цвет, остаётся с тем же ID (S3.2).
    const before = await readRuleSnapshots(ctx, range);
    const where = `${sheet.name}!${plan.resolvedAddress}`;
    if (JSON.stringify(before) !== plan.existingSignature) {
      throw new ToolExecutionError(
        `Правила условного форматирования на ${where} изменились после предпросмотра. Операция не выполнялась — сделайте новый предпросмотр.`,
        "failed_before_write"
      );
    }
    const functionsChecked = await runFunctionCheck(ctx, sheet, plan.functionCheck);

    let failure: unknown = null;
    try {
      const added = range.conditionalFormats.add(officeRuleType(request.rule) as any);
      if (request.rule === "textContains") {
        applyHighlight(added.textComparison.format, request);
        added.textComparison.rule = { operator: "Contains", text: String(request.text) } as any;
      } else if (request.rule === "formula") {
        applyHighlight(added.custom.format, request);
        added.custom.rule.formula = String(request.formula);
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
      await ctx.sync();
      // Замер: новое правило Excel ставит последним; «first» поднимает его
      // выше всех правил области — так же, как это делает интерфейс Excel.
      if (plan.order === "first" && before.length) {
        added.priority = 0;
        await ctx.sync();
      }
    } catch (error) {
      failure = error;
    }
    if (failure) {
      // Замер: на неверной формуле Excel отказывает, но пустое правило уже
      // добавлено. Его надо убрать, иначе отказ оставит след в книге.
      const message = String((failure as any)?.message ?? failure).replace(/\.+$/, "");
      const clean = await removeLeftovers(ctx, range, before);
      throw new ToolExecutionError(
        clean
          ? `Excel не принял правило на ${where}: ${message}. Добавленное им пустое правило панель убрала — правила области как до операции.`
          : `Не удалось определить итог добавления правила на ${where}: ${message}. Перечитайте правила области.`,
        clean ? "failed_before_write" : "unknown"
      );
    }

    let after: RuleSnapshot[];
    try {
      after = await readRuleSnapshots(ctx, range);
    } catch (error: any) {
      throw new ToolExecutionError(
        `Правило на ${where} добавлено, но правила области не прочитались обратно: ${error?.message ?? error}. Перечитайте правила.`,
        "applied"
      );
    }
    const position = plan.order === "first" ? 0 : after.length - 1;
    const created = after[position];
    const others = after.filter((_, index) => index !== position);
    if (after.length !== before.length + 1 || JSON.stringify(keysOf(others)) !== JSON.stringify(keysOf(before))) {
      throw new ToolExecutionError(
        `Правило на ${where} добавлено, но набор правил после операции не такой, как ожидалось: было ${before.length}, стало ${after.length}` +
          `${after.length === before.length + 1 ? ", и новое стоит не на своём месте или прежние изменились" : ""}. Перечитайте правила.`,
        "applied"
      );
    }

    // Сверка всего запрошенного: тип, область, условие, каждое свойство
    // оформления, точки шкалы — как их прочитал Excel.
    const mismatches = conditionalRuleMismatches(request, created, plan.resolvedAddress);
    if (mismatches.length) {
      throw new ToolExecutionError(
        `Правило на ${where} добавлено, но обратное чтение расходится с планом: ${mismatches.join("; ")}. Перечитайте правила области.`,
        "applied"
      );
    }

    let undoRecorded = false;
    if (plan.undoAvailable) {
      const sheetId = plan.target.sheetId;
      const address = plan.resolvedAddress;
      const createdKey = ruleKey(created);
      undoRecorded = push(action(`условное форматирование ${where}`, async () => {
        const revision = getStructuralRevision();
        await Excel.run(async (undoCtx) => {
          const undoRange = undoCtx.workbook.worksheets.getItem(sheetId).getRange(address);
          const current = await readRuleSnapshots(undoCtx, undoRange);
          // Правило узнаётся по содержимому: номер у него мог смениться.
          const matching = current.filter((rule) => ruleKey(rule) === createdKey);
          if (!matching.length) throw new Error("Добавленного правила уже нет или его изменили после операции агента. Отменять нечего.");
          if (matching.length > 1) throw new Error("На области несколько одинаковых правил, и какое из них добавил агент, не различить. Отмена остановлена — удалите лишнее в Excel.");
          if (getStructuralRevision() !== revision) throw new Error("Структура книги изменилась во время отмены. Отмена остановлена.");
          undoRange.conditionalFormats.getItem(matching[0].id).delete();
          await undoCtx.sync();
          const left = await readRuleSnapshots(undoCtx, undoRange);
          const expected = current.filter((rule) => rule !== matching[0]);
          if (JSON.stringify(keysOf(left)) !== JSON.stringify(keysOf(expected))) {
            throw new Error("Отмена удалила правило, но набор оставшихся правил не тот, что ожидался. Проверьте правила области в Excel.");
          }
        });
      }));
    }

    return {
      ok: true,
      executionState: "verified",
      sheet: sheet.name,
      address: plan.resolvedAddress,
      rule: plan.ruleText,
      rulesBefore: before.length,
      rulesAfter: after.length,
      ...(after.length > 1
        ? {
            priority: position + 1,
            priorityNote:
              `Новое правило стоит ${position + 1}-м из ${after.length} по приоритету (1 — самый высокий). ` +
              (position === 0
                ? "Где правила задают одно и то же оформление, действует новое."
                : "Где правила задают одно и то же оформление, например заливку, действуют прежние правила выше него — назови это пользователю.")
          }
        : {}),
      ...(plan.prediction ? { prediction: plan.prediction } : {}),
      ...(functionsChecked ? { functionsChecked } : {}),
      note: request.rule === "formula"
        ? "Правило проверено обратным чтением. Какие ячейки подсветит формула, считает только Excel: оценки у панели нет."
        : "Правило проверено обратным чтением. Какие ячейки оно подсветило, Excel через API не сообщает — в prediction оценка панели.",
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
  /** Ручное оформление, которое перекроет стиль таблицы. */
  readonly manualFormattingWarning?: string;
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
    // Ручное оформление Excel ставит поверх стиля таблицы. Проверка
    // 18 сентября 2026 года: таблица «с синим стилем» легла под тёмную
    // заливку шапки и сетку с прошлой проверки, и стиля почти не было видно.
    const headerRow = range.getRow(0);
    headerRow.format.fill.load("color");
    const body = range.getOffsetRange(1, 0).getResizedRange(-1, 0);
    body.format.fill.load("color");
    const insideBorder = range.format.borders.getItem("InsideHorizontal");
    insideBorder.load("style");
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
    const plainFill = (color: unknown) => typeof color === "string" && color.toUpperCase() === "#FFFFFF";
    const manualFormatting = [
      !plainFill(headerRow.format.fill.color) && "заливка шапки",
      !plainFill(body.format.fill.color) && "заливка данных",
      insideBorder.style !== null && insideBorder.style !== "None" && "границы ячеек"
    ].filter(Boolean) as string[];

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
      ...(manualFormatting.length
        ? {
            manualFormattingWarning:
              `В области есть ручное оформление: ${manualFormatting.join(", ")}. Excel оставит его поверх стиля таблицы, ` +
              "и цвета стиля там видны не будут. Если нужен чистый стиль, сначала снимите это оформление."
          }
        : {}),
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
  assertPlanWorkbook(plan);
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
