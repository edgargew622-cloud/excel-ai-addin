/**
 * Правила проверки ввода (этап 7, 7.4.2): список, целое число, число, дата.
 *
 * Замер в Excel 24 сентября 2026 года (ExcelApi 1.8–1.9):
 * - новое правило ставится поверх прежнего без очистки; область с разными
 *   правилами Excel называет `Inconsistent` — такое прежнее состояние одним
 *   правилом не вернуть, и отмены у операции тогда нет;
 * - уже введённые значения, не прошедшие правило, Excel называет сам
 *   (`getInvalidCellsOrNullObject`): для списка «Новая,В работе,Закрыта»
 *   он отбраковал «Отменена» и «в работе» — регистр в списке важен;
 *   пустые ячейки проходят;
 * - дату правило принимает в ISO (`2026-01-01`), а отдаёт в американском
 *   виде (`1/1/2026`): сверяется дата, а не текст.
 *
 * Правило не чистит уже введённое: неподходящие значения остаются, и это
 * называется, а не обещается их исправление.
 */

import { assertPlanWorkbook, assertTargetWritable, checkAddress, deepFreeze, MAX_IO_CELLS, preflightToolArgs, rangeOf, ToolError, ToolExecutionError } from "./excelTools";
import { columnLetters } from "./formulaFill";
import { action, isCustomUndoAvailable, push } from "./undo";
import { captureTarget, type WorkbookTarget } from "./workbookContext";

export type ValidationKind = "list" | "wholeNumber" | "decimal" | "date";
export type ValidationOperator = "between" | "notBetween" | "equalTo" | "notEqualTo" | "greaterThan" | "lessThan" | "greaterOrEqual" | "lessOrEqual";

const OFFICE_OPERATOR: Record<ValidationOperator, string> = {
  between: "Between", notBetween: "NotBetween", equalTo: "EqualTo", notEqualTo: "NotEqualTo",
  greaterThan: "GreaterThan", lessThan: "LessThan", greaterOrEqual: "GreaterThanOrEqualTo", lessOrEqual: "LessThanOrEqualTo"
};
const OPERATOR_TEXT: Record<ValidationOperator, string> = {
  between: "от … до", notBetween: "вне", equalTo: "равно", notEqualTo: "не равно",
  greaterThan: "больше", lessThan: "меньше", greaterOrEqual: "не меньше", lessOrEqual: "не больше"
};

export interface ValidationRequest {
  kind: ValidationKind;
  items?: string[];
  operator?: ValidationOperator;
  value?: number | string;
  value2?: number | string;
}

const isoDate = (value: unknown) => typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value);

/** Проверка запроса до Excel: неполное правило Excel примет молча или откажет на середине. */
export function parseValidationRequest(args: Record<string, unknown>): ValidationRequest {
  const kind = args.rule as ValidationKind;
  if (kind === "list") {
    const items = Array.isArray(args.items) ? args.items.map((item) => String(item).trim()).filter(Boolean) : [];
    if (!items.length) throw new ToolError("Для списка нужны items — допустимые значения.");
    const withComma = items.find((item) => item.includes(","));
    if (withComma) throw new ToolError(`Значение «${withComma}» содержит запятую: в правиле-списке запятая разделяет значения. Уберите её.`);
    if (items.join(",").length > 255) throw new ToolError("Список длиннее 255 знаков: Excel такой не примет. Сократите список.");
    return { kind, items };
  }
  if (kind !== "wholeNumber" && kind !== "decimal" && kind !== "date") throw new ToolError(`Неизвестный вид правила «${String(kind)}».`);
  const operator = (args.operator ?? "between") as ValidationOperator;
  if (!OFFICE_OPERATOR[operator]) throw new ToolError(`Неизвестный оператор «${String(operator)}».`);
  const needsTwo = operator === "between" || operator === "notBetween";
  const check = (value: unknown, name: string) => {
    if (kind === "date" ? !isoDate(value) : typeof value !== "number" || !Number.isFinite(value)) {
      throw new ToolError(`${name} — ${kind === "date" ? "дата в виде ГГГГ-ММ-ДД" : "число"}; получено «${String(value)}».`);
    }
    if (kind === "wholeNumber" && !Number.isInteger(value)) throw new ToolError(`${name} для целых чисел — целое; получено ${String(value)}.`);
  };
  check(args.value, "value");
  if (needsTwo) check(args.value2, "value2");
  return { kind, operator, value: args.value as number | string, ...(needsTwo ? { value2: args.value2 as number | string } : {}) };
}

export function describeValidation(request: ValidationRequest): string {
  if (request.kind === "list") return `список: ${request.items!.join(", ")}`;
  const what = request.kind === "wholeNumber" ? "целое число" : request.kind === "decimal" ? "число" : "дата";
  const range = request.value2 !== undefined ? `${request.value} … ${request.value2}` : String(request.value);
  return `${what} ${OPERATOR_TEXT[request.operator!]} ${range}`;
}

/** Правило в виде Office.js. */
export function officeRule(request: ValidationRequest): Record<string, unknown> {
  if (request.kind === "list") return { list: { inCellDropDown: true, source: request.items!.join(",") } };
  return {
    [request.kind]: {
      formula1: request.value,
      ...(request.value2 !== undefined ? { formula2: request.value2 } : {}),
      operator: OFFICE_OPERATOR[request.operator!]
    }
  };
}

/** Дата из того, что отдал Excel («1/1/2026») или что задано («2026-01-01»). */
function dateKey(value: unknown): string | null {
  const text = String(value ?? "");
  const iso = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text);
  if (iso) return `${Number(iso[1])}-${Number(iso[2])}-${Number(iso[3])}`;
  const us = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(text);
  if (us) return `${Number(us[3])}-${Number(us[1])}-${Number(us[2])}`;
  return null;
}

/** Совпадает ли правило, прочитанное из Excel, с запрошенным. */
export function sameValidation(request: ValidationRequest, type: unknown, rule: any): boolean {
  const expectedType = { list: "List", wholeNumber: "WholeNumber", decimal: "Decimal", date: "Date" }[request.kind];
  if (type !== expectedType) return false;
  if (request.kind === "list") return rule?.list?.source === request.items!.join(",");
  const part = rule?.[request.kind];
  if (!part || part.operator !== OFFICE_OPERATOR[request.operator!]) return false;
  const same = (actual: unknown, expected: unknown) =>
    request.kind === "date" ? dateKey(actual) === dateKey(expected) : Number(actual) === Number(expected);
  return same(part.formula1, request.value) && (request.value2 === undefined || same(part.formula2, request.value2));
}

/**
 * Какие уже введённые значения правило не пропустит — оценка панели по тем
 * же правилам, что показал замер: пусто проходит, список — с учётом регистра.
 */
export function predictInvalid(request: ValidationRequest, value: unknown): boolean {
  if (value === "" || value === null || value === undefined) return false;
  if (request.kind === "list") return !request.items!.includes(String(value));
  if (typeof value !== "number") return true;
  const toNumber = (item: unknown) => (request.kind === "date" ? dateSerial(String(item)) : Number(item));
  if (request.kind === "wholeNumber" && !Number.isInteger(value)) return true;
  const a = toNumber(request.value);
  const b = request.value2 !== undefined ? toNumber(request.value2) : NaN;
  switch (request.operator) {
    case "between": return !(value >= a && value <= b);
    case "notBetween": return value >= a && value <= b;
    case "equalTo": return value !== a;
    case "notEqualTo": return value === a;
    case "greaterThan": return !(value > a);
    case "lessThan": return !(value < a);
    case "greaterOrEqual": return !(value >= a);
    case "lessOrEqual": return !(value <= a);
  }
  return false;
}

function dateSerial(iso: string): number {
  const [y, m, d] = iso.split("-").map(Number);
  return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86_400_000);
}

/* --- план ------------------------------------------------------------------------ */

export interface ValidationPlan {
  readonly kind: "set_data_validation";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly resolvedAddress: string;
  readonly request: ValidationRequest;
  readonly ruleText: string;
  /** Прежнее правило области: тип и само правило — для предпросмотра и отмены. */
  readonly previousType: string;
  readonly previousRule: Record<string, unknown> | null;
  readonly predictedInvalid: readonly string[];
  readonly predictedInvalidCount: number;
  readonly undoAvailable: boolean;
  readonly undoNote: string;
  readonly createdAt: string;
}

/** Правило из ответа Excel, пригодное для обратной записи: без служебных полей и пустых видов. */
function cleanRule(rule: any): Record<string, unknown> | null {
  if (!rule) return null;
  const out: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(rule)) {
    if (key.startsWith("@") || value === null || value === undefined) continue;
    out[key] = Object.fromEntries(Object.entries(value as Record<string, unknown>).filter(([name, item]) => !name.startsWith("@") && item !== null));
  }
  return Object.keys(out).length ? out : null;
}

export async function prepareValidationPlan(args: unknown): Promise<ValidationPlan> {
  preflightToolArgs("set_data_validation", args);
  const a = args as { sheet?: string; address: string } & Record<string, unknown>;
  const request = parseValidationRequest(a);
  const address = checkAddress(a.address);
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
    const cells = range.rowCount * range.columnCount;
    if (cells > MAX_IO_CELLS) throw new ToolError(`Правило ставится на область до ${MAX_IO_CELLS} ячеек; ${range.address} содержит ${cells}.`);
    assertTargetWritable(sheet, range);
    const dv = range.dataValidation;
    dv.load(["type", "rule"]);
    range.load("values");
    await ctx.sync();
    const invalid: string[] = [];
    (range.values as unknown[][]).forEach((row, r) => row.forEach((value, c) => {
      if (predictInvalid(request, value)) invalid.push(`${columnLetters(range.columnIndex + c + 1)}${range.rowIndex + r + 1}`);
    }));
    const previousType = String(dv.type);
    const previousRule = previousType === "None" || previousType === "Inconsistent" ? null : cleanRule(dv.rule);
    const undo = isCustomUndoAvailable() && previousType !== "Inconsistent";
    return {
      kind: "set_data_validation" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      target: { ...target, sheetName: sheet.name },
      resolvedAddress: range.address.replace(/^.*!/, ""),
      request,
      ruleText: describeValidation(request),
      previousType,
      previousRule,
      predictedInvalid: invalid.slice(0, 20),
      predictedInvalidCount: invalid.length,
      undoAvailable: undo,
      undoNote: previousType === "Inconsistent"
        ? "Отмены нет: на области были разные правила, и одним правилом их не вернуть."
        : undo
          ? previousType === "None" ? "Отмена снимет правило." : "Отмена вернёт прежнее правило."
          : "Отмена недоступна: монитор изменений Excel не активен.",
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executeValidationPlan(plan: ValidationPlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(plan.target.sheetId);
    sheet.load("name");
    const range = sheet.getRange(plan.resolvedAddress);
    const dv = range.dataValidation;
    dv.load(["type", "rule"]);
    await ctx.sync();
    const where = `${sheet.name}!${plan.resolvedAddress}`;
    if (String(dv.type) !== plan.previousType || JSON.stringify(cleanRule(dv.rule)) !== JSON.stringify(plan.previousRule)) {
      throw new ToolExecutionError(`Правило проверки ввода на ${where} изменилось после предпросмотра. Операция не выполнялась — сделайте новый предпросмотр.`, "failed_before_write");
    }
    try {
      dv.rule = officeRule(plan.request) as any;
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(`Не удалось определить итог установки правила на ${where}: ${error?.message ?? error}. Перечитайте правило.`, "unknown");
    }

    dv.load(["type", "rule"]);
    const bad = dv.getInvalidCellsOrNullObject();
    bad.load(["isNullObject", "address", "cellCount"]);
    await ctx.sync();
    const invalidByExcel = bad.isNullObject ? [] : String(bad.address).split(",").map((item) => item.replace(/^.*!/, ""));

    let undoRecorded = false;
    if (plan.undoAvailable) {
      undoRecorded = push(action(`правило ввода ${where}`, async () => {
        await Excel.run(async (undoCtx) => {
          const undoRange = undoCtx.workbook.worksheets.getItem(plan.target.sheetId).getRange(plan.resolvedAddress);
          const current = undoRange.dataValidation;
          current.load(["type", "rule"]);
          await undoCtx.sync();
          if (!sameValidation(plan.request, current.type, current.rule)) throw new Error(`Правило на ${where} изменили после операции агента. Отмена остановлена.`);
          if (plan.previousRule) current.rule = plan.previousRule as any;
          else current.clear();
          await undoCtx.sync();
        });
      }));
    }

    if (!sameValidation(plan.request, dv.type, dv.rule)) {
      throw new ToolExecutionError(
        `Правило на ${where} установлено, но обратное чтение расходится с запрошенным (${String(dv.type)}). ${undoRecorded ? "Его можно отменить кнопкой «Отменить»." : "Проверьте правило."}`,
        "applied"
      );
    }
    return {
      ok: true,
      executionState: "verified",
      address: where,
      rule: plan.ruleText,
      replaced: plan.previousType === "None" ? null : plan.previousType,
      // Нарушителей называет сам Excel: это факт, а не оценка панели.
      invalidExistingCells: invalidByExcel.length,
      ...(invalidByExcel.length
        ? {
            invalidExamples: invalidByExcel.slice(0, 20),
            invalidNote: "Эти уже введённые значения правилу не соответствуют. Правило их не исправляет и не удаляет — назови их пользователю; Excel будет отклонять только новый неподходящий ввод."
          }
        : {}),
      undoable: undoRecorded,
      undoNote: plan.undoNote
    };
  });
}
