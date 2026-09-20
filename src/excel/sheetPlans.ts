/**
 * Создание листа.
 *
 * Самая безопасная из структурных операций: она ничего не рушит и откат
 * у неё честный — удалить созданный лист, пока он пуст. Пустоту приходится
 * проверять при отмене: если на листе уже успели поработать, удаление
 * унесло бы чужую работу, и такая отмена останавливается.
 *
 * Имя проверяется заранее (`sheetRules.ts`): правила Excel жёсткие
 * и молчаливые, и отказ на середине операции хуже отказа до неё.
 */

import { deepFreeze, preflightToolArgs, ToolError, ToolExecutionError } from "./excelTools";
import { checkSheetName, freeSheetName } from "./sheetRules";
import { action, getStructuralRevision, isCustomUndoAvailable, push } from "./undo";
import { currentWorkbookIdentity } from "./workbookContext";

export interface CreateSheetPlan {
  readonly kind: "create_sheet";
  readonly id: string;
  readonly workbook: ReturnType<typeof currentWorkbookIdentity>;
  readonly name: string;
  /** Имя листа, после которого встанет новый; пусто — в конец книги. */
  readonly after?: string;
  readonly positionText: string;
  /** Листы книги на момент предпросмотра: по ним ловится чужая правка. */
  readonly sheetsBefore: readonly string[];
  readonly undoAvailable: boolean;
  readonly undoNote: string;
  readonly createdAt: string;
}

async function listSheetNames(ctx: Excel.RequestContext): Promise<{ names: string[]; positions: Map<string, number> }> {
  const collection = ctx.workbook.worksheets;
  collection.load("items/name,items/position");
  await ctx.sync();
  const positions = new Map<string, number>();
  for (const sheet of collection.items) positions.set(sheet.name, sheet.position);
  return { names: collection.items.map((sheet) => sheet.name), positions };
}

export async function prepareCreateSheetPlan(args: unknown): Promise<CreateSheetPlan> {
  preflightToolArgs("create_sheet", args);
  const a = args as { name: string; after?: string };

  const prepared = await Excel.run(async (ctx) => {
    const { names } = await listSheetNames(ctx);
    let name: string;
    try {
      name = checkSheetName(a.name, names);
    } catch (error: any) {
      const message = String(error?.message ?? error);
      // Отказ по занятому имени бесполезен без свободного варианта:
      // по опыту сводной агент на таком просто бросает задачу.
      throw new ToolError(
        /уже есть/.test(message)
          ? `${message} Свободно, например, «${freeSheetName(String(a.name), names)}».`
          : message
      );
    }
    const after = a.after?.trim();
    if (after && !names.some((item) => item.trim().toLowerCase() === after.toLowerCase())) {
      throw new ToolError(`Листа «${after}» нет, после него вставить нельзя. Листы книги: ${names.map((item) => `«${item}»`).join(", ")}.`);
    }

    const undo = isCustomUndoAvailable();
    return {
      kind: "create_sheet" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      workbook: currentWorkbookIdentity(),
      name,
      ...(after ? { after } : {}),
      positionText: after ? `сразу после листа «${after}»` : "последним в книге",
      sheetsBefore: names,
      undoAvailable: undo,
      undoNote: undo
        ? "Отмена удалит созданный лист, но только пока он пуст: если на нём успеют поработать, отмена остановится."
        : "Отмена недоступна: монитор изменений Excel не активен. Лишний лист придётся удалить вручную.",
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executeCreateSheetPlan(plan: CreateSheetPlan) {
  const identity = currentWorkbookIdentity();
  if (identity.workbookSessionId !== plan.workbook.workbookSessionId || identity.documentUrl !== plan.workbook.documentUrl) {
    throw new ToolExecutionError("Открытая книга изменилась после предпросмотра. Лист не создавался — сделайте новый предпросмотр.", "failed_before_write");
  }

  return Excel.run(async (ctx) => {
    const { names, positions } = await listSheetNames(ctx);
    // Имя могли занять руками между предпросмотром и подтверждением.
    if (names.some((item) => item.trim().toLowerCase() === plan.name.toLowerCase())) {
      throw new ToolExecutionError(
        `Лист «${plan.name}» появился в книге после предпросмотра. Лист не создавался; свободно, например, «${freeSheetName(plan.name, names)}».`,
        "failed_before_write"
      );
    }
    if (plan.after && !positions.has(plan.after)) {
      throw new ToolExecutionError(`Листа «${plan.after}» больше нет: непонятно, куда вставлять. Лист не создавался.`, "failed_before_write");
    }

    let sheet: Excel.Worksheet;
    try {
      sheet = ctx.workbook.worksheets.add(plan.name);
      if (plan.after) sheet.position = (positions.get(plan.after) ?? 0) + 1;
      sheet.load(["id", "name", "position"]);
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(
        `Excel отказал в создании листа «${plan.name}»: ${error?.message ?? error}. Неизвестно, появился ли он — посмотрите на книгу.`,
        "unknown"
      );
    }

    // Excel молча исправляет неподходящее имя, поэтому имя сверяется.
    if (sheet.name !== plan.name) {
      throw new ToolExecutionError(
        `Лист создан, но Excel назвал его «${sheet.name}» вместо «${plan.name}». Проверьте книгу.`,
        "applied"
      );
    }

    const sheetId = sheet.id;
    const sheetName = sheet.name;
    let undoRecorded = false;
    if (plan.undoAvailable) {
      undoRecorded = push(action(`создание листа «${sheetName}»`, async () => {
        const revision = getStructuralRevision();
        await Excel.run(async (undoCtx) => {
          const existing = undoCtx.workbook.worksheets.getItemOrNullObject(sheetId);
          existing.load(["isNullObject", "name"]);
          await undoCtx.sync();
          if (existing.isNullObject) throw new Error("Листа уже нет: его удалили после операции агента. Отменять нечего.");
          const used = existing.getUsedRangeOrNullObject(true);
          used.load(["isNullObject", "address"]);
          await undoCtx.sync();
          if (!used.isNullObject) {
            throw new Error(
              `На листе «${existing.name}» уже есть данные в ${used.address}. Отмена остановлена: удаление листа унесло бы их. Удалите лист вручную, если он не нужен.`
            );
          }
          if (getStructuralRevision() !== revision) throw new Error("Структура книги изменилась во время отмены. Отмена остановлена.");
          existing.delete();
          await undoCtx.sync();
        });
      }));
    }

    const after = await listSheetNames(ctx);
    return {
      ok: true,
      executionState: "verified",
      sheet: sheetName,
      sheetId,
      position: sheet.position + 1,
      placedAfter: plan.after ?? null,
      sheetsAfter: after.names,
      note: "Лист создан пустым. Данные и другие листы не менялись.",
      undoable: undoRecorded,
      undoNote: plan.undoNote
    };
  });
}
