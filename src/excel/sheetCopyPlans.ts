/**
 * Копия листа (этап 8, 8.3.5).
 *
 * План предполагал, что `worksheet.copy` есть только в Microsoft 365. Замер
 * 28 сентября 2026 года на Office 2021 (ExcelApi 1.10+): лист копируется
 * вместе со значениями, формулами и условным форматированием, копия встаёт
 * за исходным и получает имя «Лист (2)». Копия сверяется с исходным по
 * занятой области, значениям, числу формул, таблиц, диаграмм, сводных
 * и правил условного форматирования.
 */

import {
  assertPlanWorkbook,
  deepFreeze,
  preflightToolArgs,
  ToolError,
  ToolExecutionError
} from "./excelTools";
import { checkSheetName, freeSheetName } from "./sheetRules";
import { action, getStructuralRevision, isCustomUndoAvailable, push } from "./undo";
import { captureTarget, type WorkbookTarget } from "./workbookContext";

export const SHEET_COPY_API = "1.10";

export interface SheetFacts {
  /** Занятая область без имени листа; null — лист пуст. */
  used: string | null;
  values: string;
  formulas: number;
  tables: number;
  charts: number;
  pivots: number;
  conditionalFormats: number;
}

export interface CopySheetPlan {
  readonly kind: "copy_sheet";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly sourceName: string;
  /** Имя копии; пусто — Excel назовёт сам («Лист (2)»). */
  readonly newName?: string;
  readonly position: "after" | "end";
  readonly facts: SheetFacts;
  readonly warnings: readonly string[];
  readonly undoAvailable: boolean;
  readonly undoNote?: string;
  readonly createdAt: string;
}

const MAX_COMPARE_CELLS = 200_000;

export async function readSheetFacts(ctx: Excel.RequestContext, sheet: Excel.Worksheet): Promise<SheetFacts> {
  const used = sheet.getUsedRangeOrNullObject(true);
  used.load(["isNullObject", "address", "rowCount", "columnCount"]);
  sheet.tables.load("items/name");
  sheet.charts.load("items/name");
  sheet.pivotTables.load("items/name");
  await ctx.sync();
  let values = "";
  let formulas = 0;
  let conditionalFormats = 0;
  if (!used.isNullObject) {
    if (used.rowCount * used.columnCount > MAX_COMPARE_CELLS) {
      throw new ToolError(`Лист слишком большой для сверки копии: ${used.rowCount * used.columnCount} ячеек, панель сверяет до ${MAX_COMPARE_CELLS}.`);
    }
    used.load(["values", "formulas"]);
    const rules = used.conditionalFormats;
    rules.load("items/type");
    await ctx.sync();
    values = JSON.stringify(used.values);
    formulas = (used.formulas as unknown[][]).flat().filter((cell) => typeof cell === "string" && cell.startsWith("=")).length;
    conditionalFormats = rules.items.length;
  }
  return {
    used: used.isNullObject ? null : String(used.address).slice(String(used.address).lastIndexOf("!") + 1),
    values,
    formulas,
    tables: sheet.tables.items.length,
    charts: sheet.charts.items.length,
    pivots: sheet.pivotTables.items.length,
    conditionalFormats
  };
}

export function factMismatches(copy: SheetFacts, source: SheetFacts): string[] {
  const problems: string[] = [];
  if (copy.used !== source.used) problems.push(`занятая область ${copy.used ?? "пусто"} вместо ${source.used ?? "пусто"}`);
  else if (copy.values !== source.values) problems.push("значения ячеек отличаются от исходного листа");
  const counts: [keyof SheetFacts, string][] = [["formulas", "формул"], ["tables", "таблиц"], ["charts", "диаграмм"], ["pivots", "сводных"], ["conditionalFormats", "правил условного форматирования"]];
  for (const [key, text] of counts) if (copy[key] !== source[key]) problems.push(`${text} ${copy[key]} вместо ${source[key]}`);
  return problems;
}

export async function prepareCopySheetPlan(args: unknown): Promise<CopySheetPlan> {
  preflightToolArgs("copy_sheet", args);
  const a = (args ?? {}) as { sheet?: string; newName?: string; position?: string };
  const target = await captureTarget(a.sheet);
  const prepared = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    sheet.load(["id", "name"]);
    const all = ctx.workbook.worksheets;
    all.load("items/name");
    await ctx.sync();
    const names = all.items.map((item) => item.name);
    let newName: string | undefined;
    if (a.newName?.trim()) {
      try {
        newName = checkSheetName(a.newName, names);
      } catch (error: any) {
        const message = String(error?.message ?? error);
        throw new ToolError(/уже есть/.test(message) ? `${message} Свободно, например, «${freeSheetName(String(a.newName), names)}».` : message);
      }
    }
    const facts = await readSheetFacts(ctx, sheet);
    const warnings: string[] = [];
    if (facts.tables) warnings.push(`На листе таблиц: ${facts.tables}. У копий Excel сменит имена — формулы с именами таблиц по-прежнему смотрят на исходные.`);
    if (facts.pivots) warnings.push(`На листе сводных: ${facts.pivots}. Копии сводных берут данные из того же источника.`);
    if (facts.formulas) warnings.push(`Формул: ${facts.formulas}. Ссылки без имени листа в копии смотрят на саму копию, ссылки на другие листы — туда же, куда и раньше.`);
    const undo = isCustomUndoAvailable();
    return {
      kind: "copy_sheet" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      target: { ...target, sheetName: sheet.name },
      sourceName: sheet.name,
      ...(newName ? { newName } : {}),
      position: a.position === "end" ? "end" as const : "after" as const,
      facts,
      warnings,
      undoAvailable: undo,
      ...(undo ? {} : { undoNote: "Отмена недоступна: монитор изменений Excel не активен." }),
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executeCopySheetPlan(plan: CopySheetPlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(plan.target.sheetId);
    sheet.load(["id", "name"]);
    const facts = await readSheetFacts(ctx, sheet);
    if (JSON.stringify(facts) !== JSON.stringify(plan.facts)) {
      throw new ToolExecutionError(`Лист ${sheet.name} изменился после предпросмотра. Копия не делалась — сделайте новый предпросмотр.`, "failed_before_write");
    }
    if (plan.newName) {
      const existing = ctx.workbook.worksheets.getItemOrNullObject(plan.newName);
      existing.load("isNullObject");
      await ctx.sync();
      if (!existing.isNullObject) throw new ToolExecutionError(`Лист «${plan.newName}» появился после предпросмотра. Копия не делалась — выберите другое имя.`, "failed_before_write");
    }

    let copy: Excel.Worksheet;
    try {
      copy = sheet.copy((plan.position === "end" ? "End" : "After") as any, plan.position === "end" ? undefined : sheet);
      copy.load(["id", "name", "position"]);
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(`Excel отказал в копировании листа ${sheet.name}: ${error?.message ?? error}. Посмотрите, не появилась ли копия.`, "unknown");
    }
    let renameProblem: string | null = null;
    if (plan.newName) {
      try {
        copy.name = plan.newName;
        await ctx.sync();
        copy.load("name");
        await ctx.sync();
      } catch (error: any) {
        renameProblem = `переименовать копию в «${plan.newName}» не удалось (${error?.message ?? error}), она называется «${copy.name}»`;
      }
    }
    const copyFacts = await readSheetFacts(ctx, copy);
    const problems = [...(renameProblem ? [renameProblem] : []), ...factMismatches(copyFacts, plan.facts)];

    let undoRecorded = false;
    if (plan.undoAvailable) {
      const copyId = copy.id;
      const copyName = copy.name;
      const signature = JSON.stringify(copyFacts);
      undoRecorded = push(action(`копия листа ${plan.sourceName} («${copyName}»)`, async () => {
        const revision = getStructuralRevision();
        await Excel.run(async (undoCtx) => {
          const created = undoCtx.workbook.worksheets.getItemOrNullObject(copyId);
          created.load("isNullObject");
          await undoCtx.sync();
          if (created.isNullObject) throw new Error("Копии листа уже нет. Отменять нечего.");
          const now = await readSheetFacts(undoCtx, created);
          if (JSON.stringify(now) !== signature) throw new Error(`В копию «${copyName}» внесли изменения после операции агента. Отмена остановлена, чтобы не удалить чужую работу.`);
          if (getStructuralRevision() !== revision) throw new Error("Структура книги изменилась во время отмены. Отмена остановлена.");
          created.delete();
          await undoCtx.sync();
        });
      }));
    }

    if (problems.length) {
      throw new ToolExecutionError(
        `Лист ${plan.sourceName} скопирован в «${copy.name}», но копия расходится с исходным: ${problems.join("; ")}. ` +
        (undoRecorded ? "Её можно убрать кнопкой «Отменить»." : "Проверьте копию."),
        "applied"
      );
    }
    return {
      ok: true,
      executionState: "verified",
      source: plan.sourceName,
      copy: copy.name,
      position: copy.position + 1,
      checked: `занятая область ${copyFacts.used ?? "пусто"}, значения, формул ${copyFacts.formulas}, таблиц ${copyFacts.tables}, диаграмм ${copyFacts.charts}, сводных ${copyFacts.pivots}, правил условного форматирования ${copyFacts.conditionalFormats}`,
      ...(plan.warnings.length ? { warnings: plan.warnings } : {}),
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: plan.undoNote ?? "Автоматическая отмена этой операции недоступна." })
    };
  });
}
