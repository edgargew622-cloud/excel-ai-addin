/**
 * Порядок правил условного форматирования (этап 7, 7.4.3).
 *
 * Замер в Excel 24 сентября 2026 года (ExcelApi 1.6):
 * - правила области Excel отдаёт в порядке приоритета; ID правила — его номер
 *   в этом списке, приоритет — место в нём же, и оба сдвигаются, когда
 *   правило встаёт выше;
 * - приоритет 0, заданный через список области, поднимает правило на место
 *   самого высокого правила этой области; приоритет n−1 опускает на место
 *   самого низкого; число вне списка Excel молча пропускает.
 *
 * Поэтому правило узнаётся по содержимому (`ruleKey`), а номер из списка
 * `get_conditional_formats` служит только для выбора.
 */

import {
  assertPlanWorkbook,
  assertTargetWritable,
  checkAddress,
  deepFreeze,
  preflightToolArgs,
  rangeOf,
  ToolError,
  ToolExecutionError
} from "./excelTools";
import { readRuleSnapshots, ruleKey } from "./sheetFormatPlans";
import type { RuleSnapshot } from "./sheetRules";
import { action, getStructuralRevision, isCustomUndoAvailable, push } from "./undo";
import { captureTarget, type WorkbookTarget } from "./workbookContext";

const OPERATOR_TEXT: Record<string, string> = {
  GreaterThan: "больше", LessThan: "меньше", GreaterThanOrEqual: "не меньше", LessThanOrEqual: "не больше",
  EqualTo: "равно", NotEqualTo: "не равно", Between: "от", NotBetween: "вне"
};

/** Одна строка о правиле, прочитанном из Excel. */
export function describeRule(rule: RuleSnapshot): string {
  const bare = (value: unknown) => String(value ?? "").replace(/^=/, "");
  const look = [
    rule.fill && `заливка ${rule.fill}`,
    rule.fontColor && `текст ${rule.fontColor}`,
    rule.bold === true && "жирный"
  ].filter(Boolean).join(", ");
  const tail = look ? ` → ${look}` : "";
  if (rule.contentUnread) return `${rule.type}: содержимое панель не читает`;
  switch (rule.type) {
    case "CellValue": {
      const op = String(rule.rule?.operator ?? "");
      const range = op === "Between" || op === "NotBetween" ? `${bare(rule.rule?.formula1)} … ${bare(rule.rule?.formula2)}` : bare(rule.rule?.formula1);
      return `значение ${OPERATOR_TEXT[op] ?? op} ${range}${tail}`;
    }
    case "ContainsText": return `текст содержит «${String(rule.rule?.text ?? "")}»${tail}`;
    case "Custom": return `формула ${String(rule.rule?.formula ?? "")}${tail}`;
    case "ColorScale": return "цветовая шкала";
    case "DataBar": return `гистограмма ${rule.barColor ?? ""}`.trim();
    default: return rule.type;
  }
}

const areaOf = (rule: RuleSnapshot) => (rule.range === null ? null : rule.range.slice(rule.range.lastIndexOf("!") + 1));

/** Список правил области в порядке приоритета — только чтение. */
export async function listConditionalFormats(args: { sheet?: string; address: string }) {
  const address = checkAddress(args.address);
  const target = await captureTarget(args.sheet);
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    const range = await rangeOf(ctx, sheet, address);
    range.load("address");
    sheet.load("name");
    await ctx.sync();
    const rules = await readRuleSnapshots(ctx, range);
    return {
      sheet: sheet.name,
      address: String(range.address).replace(/^.*!/, ""),
      rules: rules.map((rule, index) => ({ position: index + 1, type: rule.type, appliesTo: areaOf(rule), rule: describeRule(rule) })),
      note: rules.length > 1
        ? "Порядок — по приоритету: 1 — самый высокий. Где правила задают одно и то же оформление, видно правило выше. Номер position годится для move_conditional_format."
        : rules.length ? "На области одно правило." : "На области нет правил условного форматирования."
    };
  });
}

/* --- перенос правила ---------------------------------------------------------- */

export interface MoveRulePlan {
  readonly kind: "move_conditional_format";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly resolvedAddress: string;
  readonly from: number;
  readonly to: "first" | "last";
  readonly ruleText: string;
  readonly ruleArea: string | null;
  readonly orderBefore: readonly string[];
  readonly signature: string;
  readonly undoAvailable: boolean;
  readonly undoNote: string;
  readonly createdAt: string;
}

/** Порядок после переноса: правило с места from — первым или последним. */
export function movedOrder<T>(items: readonly T[], from: number, to: "first" | "last"): T[] {
  const rest = items.filter((_, index) => index !== from);
  return to === "first" ? [items[from], ...rest] : [...rest, items[from]];
}

export async function prepareMoveRulePlan(args: unknown): Promise<MoveRulePlan> {
  preflightToolArgs("move_conditional_format", args);
  const a = args as { sheet?: string; address: string; position: number; to: "first" | "last" };
  const address = checkAddress(a.address);
  const target = await captureTarget(a.sheet);
  const prepared = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    const range = await rangeOf(ctx, sheet, address);
    range.load(["address", "rowCount", "columnCount"]);
    sheet.load(["id", "name"]);
    try {
      range.format?.protection?.load("locked");
      sheet.protection?.load(["protected", "options"]);
    } catch { /* среда без сведений о защите */ }
    await ctx.sync();
    assertTargetWritable(sheet, range, "format");
    const rules = await readRuleSnapshots(ctx, range);
    if (rules.length < 2) throw new ToolError(`На ${sheet.name}!${String(range.address).replace(/^.*!/, "")} правил ${rules.length}: переставлять нечего.`);
    const from = Number(a.position) - 1;
    if (!Number.isInteger(from) || from < 0 || from >= rules.length) {
      throw new ToolError(`На области правил ${rules.length}; position — от 1 до ${rules.length}. Список даёт get_conditional_formats.`);
    }
    if ((a.to === "first" && from === 0) || (a.to === "last" && from === rules.length - 1)) {
      throw new ToolError(`Правило ${a.position} уже стоит ${a.to === "first" ? "первым" : "последним"}. Переставлять нечего.`);
    }
    const keys = rules.map(ruleKey);
    if (keys.filter((key) => key === keys[from]).length > 1) {
      throw new ToolError("На области несколько одинаковых правил: после переноса их не различить, и сверка с отменой невозможны. Удалите лишнее в Excel.");
    }
    const undo = isCustomUndoAvailable();
    return {
      kind: "move_conditional_format" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      target: { ...target, sheetName: sheet.name },
      resolvedAddress: String(range.address).replace(/^.*!/, ""),
      from,
      to: a.to,
      ruleText: describeRule(rules[from]),
      ruleArea: areaOf(rules[from]),
      orderBefore: rules.map(describeRule),
      signature: JSON.stringify(rules),
      undoAvailable: undo,
      undoNote: undo ? "Отмена вернёт правило на прежнее место." : "Отмена недоступна: монитор изменений Excel не активен.",
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executeMoveRulePlan(plan: MoveRulePlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(plan.target.sheetId);
    sheet.load("name");
    const range = sheet.getRange(plan.resolvedAddress);
    const before = await readRuleSnapshots(ctx, range);
    const where = `${sheet.name}!${plan.resolvedAddress}`;
    if (JSON.stringify(before) !== plan.signature) {
      throw new ToolExecutionError(`Правила на ${where} изменились после предпросмотра. Порядок не менялся — сделайте новый предпросмотр.`, "failed_before_write");
    }
    try {
      range.conditionalFormats.getItem(before[plan.from].id).priority = plan.to === "first" ? 0 : before.length - 1;
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(`Не удалось определить итог переноса правила на ${where}: ${error?.message ?? error}. Перечитайте правила.`, "unknown");
    }
    const after = await readRuleSnapshots(ctx, range);
    const expected = movedOrder(before, plan.from, plan.to).map(ruleKey);
    const movedKey = ruleKey(before[plan.from]);

    let undoRecorded = false;
    if (plan.undoAvailable) {
      const sheetId = plan.target.sheetId;
      const address = plan.resolvedAddress;
      undoRecorded = push(action(`порядок правил ${where}`, async () => {
        const revision = getStructuralRevision();
        await Excel.run(async (undoCtx) => {
          const undoRange = undoCtx.workbook.worksheets.getItem(sheetId).getRange(address);
          const current = await readRuleSnapshots(undoCtx, undoRange);
          if (JSON.stringify(current.map(ruleKey)) !== JSON.stringify(expected)) {
            throw new Error(`Правила на ${where} изменили после операции агента. Отмена остановлена, чтобы не переставить чужое.`);
          }
          if (getStructuralRevision() !== revision) throw new Error("Структура книги изменилась во время отмены. Отмена остановлена.");
          const index = current.findIndex((rule) => ruleKey(rule) === movedKey);
          undoRange.conditionalFormats.getItem(current[index].id).priority = plan.from;
          await undoCtx.sync();
          const back = await readRuleSnapshots(undoCtx, undoRange);
          if (JSON.stringify(back.map(ruleKey)) !== JSON.stringify(before.map(ruleKey))) {
            throw new Error(`Отмена переставила правило, но порядок на ${where} не вернулся к прежнему. Проверьте правила в Excel.`);
          }
        });
      }));
    }

    if (JSON.stringify(after.map(ruleKey)) !== JSON.stringify(expected)) {
      throw new ToolExecutionError(
        `Приоритет правила на ${where} задан, но порядок после операции не тот, что ожидался: ${after.map(describeRule).join(" | ")}. ${undoRecorded ? "Прежний порядок вернёт «Отменить»." : "Проверьте правила."}`,
        "applied"
      );
    }
    return {
      ok: true,
      executionState: "verified",
      address: where,
      moved: plan.ruleText,
      to: plan.to,
      order: after.map((rule, index) => `${index + 1}. ${describeRule(rule)}`),
      ...(plan.ruleArea && plan.ruleArea.replace(/\$/g, "") !== plan.resolvedAddress
        ? { areaNote: `Правило действует на ${plan.ruleArea}, не только на ${plan.resolvedAddress}: новый порядок виден на всей его области.` }
        : {}),
      undoable: undoRecorded,
      undoNote: plan.undoNote
    };
  });
}
