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

import { assertPlanWorkbook, deepFreeze, MAX_IO_CELLS, preflightToolArgs, scanWorkbookFormulas, ToolError, ToolExecutionError } from "./excelTools";
import { columnLetters } from "./formulaFill";
import { countRefErrors } from "./rowOps";
import { referencesSheet, sheetImpact, type CellMention } from "./sheetImpact";
import { lastWorkbookBackup } from "./workbookBackup";
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
  assertPlanWorkbook(plan);
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

/* =========================================================================
 * Переименование и удаление листа (этап 7, 7.3.1)
 * ========================================================================= */

/** Именованные диапазоны, указывающие на лист: при удалении они ломаются. */
async function namesOnSheet(ctx: Excel.RequestContext, sheetName: string): Promise<string[]> {
  try {
    const names = ctx.workbook.names;
    names.load("items/name,items/formula");
    await ctx.sync();
    return names.items.filter((item) => referencesSheet(String(item.formula ?? ""), sheetName)).map((item) => item.name);
  } catch {
    return [];
  }
}

async function findSheet(ctx: Excel.RequestContext, name: unknown): Promise<{ sheet: Excel.Worksheet; names: string[] }> {
  const { names } = await listSheetNames(ctx);
  const wanted = String(name ?? "").trim().toLowerCase();
  const exact = names.find((item) => item.trim().toLowerCase() === wanted);
  if (!exact) throw new ToolError(`Листа «${String(name ?? "")}» нет. Листы книги: ${names.map((item) => `«${item}»`).join(", ")}.`);
  const sheet = ctx.workbook.worksheets.getItem(exact);
  sheet.load(["id", "name", "visibility", "position"]);
  await ctx.sync();
  return { sheet, names };
}

/** Структура книги под защитой — листы не переименовать и не удалить. */
async function assertStructureEditable(ctx: Excel.RequestContext): Promise<void> {
  try {
    const protection = (ctx.workbook as any).protection;
    protection.load("protected");
    await ctx.sync();
    if (protection.protected) throw new ToolError("Структура книги защищена: листы нельзя переименовывать и удалять, пока защита не снята. Операция не выполнялась.");
  } catch (error) {
    if (error instanceof ToolError) throw error;
    /* среда без сведений о защите книги */
  }
}

export interface RenameSheetPlan {
  readonly kind: "rename_sheet";
  readonly id: string;
  readonly workbook: ReturnType<typeof currentWorkbookIdentity>;
  readonly sheetId: string;
  readonly oldName: string;
  readonly newName: string;
  /** Формулы других листов, ссылающиеся на лист: их значения после переименования не должны измениться. */
  readonly referencing: readonly CellMention[];
  readonly referencingValues: readonly unknown[];
  /** Формулы с именем листа внутри текста — Excel их не перепишет. */
  readonly literal: readonly CellMention[];
  readonly textMentions: readonly CellMention[];
  readonly overflow: number;
  readonly unscannedSheets: readonly string[];
  readonly undoAvailable: boolean;
  readonly undoNote: string;
  readonly createdAt: string;
}

/** Значения ячеек по адресам «Лист!A1» — для сверки до и после. */
async function readCells(ctx: Excel.RequestContext, cells: readonly CellMention[]): Promise<unknown[]> {
  const ranges = cells.map((item) => {
    const range = ctx.workbook.worksheets.getItem(item.sheet).getRange(item.cell);
    range.load("values");
    return range;
  });
  if (ranges.length) await ctx.sync();
  return ranges.map((range) => (range.values as unknown[][])[0][0]);
}

export async function prepareRenameSheetPlan(args: unknown): Promise<RenameSheetPlan> {
  preflightToolArgs("rename_sheet", args);
  const a = args as { sheet: string; newName: string };
  const prepared = await Excel.run(async (ctx) => {
    await assertStructureEditable(ctx);
    const { sheet, names } = await findSheet(ctx, a.sheet);
    let newName: string;
    try {
      newName = checkSheetName(a.newName, names.filter((item) => item !== sheet.name));
    } catch (error: any) {
      const message = String(error?.message ?? error);
      throw new ToolError(/уже есть/.test(message) ? `${message} Свободно, например, «${freeSheetName(String(a.newName), names)}».` : message);
    }
    if (newName === sheet.name) throw new ToolError(`Лист уже называется «${newName}» — менять нечего.`);
    const scan = await scanWorkbookFormulas(ctx);
    const impact = sheetImpact(scan.sheets, sheet.name, [], columnLetters);
    const referencingValues = await readCells(ctx, impact.referencing);
    const undo = isCustomUndoAvailable();
    return {
      kind: "rename_sheet" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      workbook: currentWorkbookIdentity(),
      sheetId: sheet.id,
      oldName: sheet.name,
      newName,
      referencing: impact.referencing,
      referencingValues,
      literal: impact.literal,
      textMentions: impact.textMentions,
      overflow: impact.overflow,
      unscannedSheets: scan.unscanned,
      undoAvailable: undo,
      undoNote: undo
        ? "Отмена вернёт прежнее имя, если его к тому времени не заняли."
        : "Отмена недоступна: монитор изменений Excel не активен. Прежнее имя можно вернуть переименованием.",
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executeRenameSheetPlan(plan: RenameSheetPlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItemOrNullObject(plan.sheetId);
    sheet.load(["isNullObject", "name"]);
    await ctx.sync();
    if (sheet.isNullObject) throw new ToolExecutionError(`Листа «${plan.oldName}» больше нет. Переименование не выполнялось.`, "failed_before_write");
    if (sheet.name !== plan.oldName) {
      throw new ToolExecutionError(`Лист уже переименован в «${sheet.name}» после предпросмотра. Переименование не выполнялось.`, "failed_before_write");
    }
    const { names } = await listSheetNames(ctx);
    if (names.some((item) => item !== plan.oldName && item.trim().toLowerCase() === plan.newName.toLowerCase())) {
      throw new ToolExecutionError(`Имя «${plan.newName}» заняли после предпросмотра. Переименование не выполнялось.`, "failed_before_write");
    }
    try {
      sheet.name = plan.newName;
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(`Excel отказал в переименовании: ${error?.message ?? error}. Перечитайте список листов.`, "unknown");
    }

    // Тот же лист (по ID) с новым именем, а формулы, ссылавшиеся на него, дают прежние значения.
    const renamed = ctx.workbook.worksheets.getItem(plan.sheetId);
    renamed.load("name");
    await ctx.sync();
    const valuesAfter = await readCells(ctx, plan.referencing);
    const changed = plan.referencing
      .map((item, index) => (JSON.stringify(valuesAfter[index]) === JSON.stringify(plan.referencingValues[index]) ? null : `${item.sheet}!${item.cell}`))
      .filter((item): item is string => item !== null);

    let undoRecorded = false;
    if (plan.undoAvailable) {
      undoRecorded = push(action(`переименование листа «${plan.oldName}» → «${plan.newName}»`, async () => {
        await Excel.run(async (undoCtx) => {
          const target = undoCtx.workbook.worksheets.getItemOrNullObject(plan.sheetId);
          target.load(["isNullObject", "name"]);
          await undoCtx.sync();
          if (target.isNullObject) throw new Error("Листа уже нет: его удалили после операции агента. Отменять нечего.");
          if (target.name !== plan.newName) throw new Error(`Лист уже называется «${target.name}», а не «${plan.newName}». Отмена остановлена.`);
          const { names: current } = await listSheetNames(undoCtx);
          if (current.some((item) => item.trim().toLowerCase() === plan.oldName.toLowerCase())) {
            throw new Error(`Имя «${plan.oldName}» уже занято другим листом. Отмена остановлена.`);
          }
          target.name = plan.oldName;
          await undoCtx.sync();
        });
      }));
    }

    if (renamed.name !== plan.newName || changed.length) {
      throw new ToolExecutionError(
        (renamed.name !== plan.newName ? `Лист называется «${renamed.name}», а не «${plan.newName}». ` : "") +
        (changed.length ? `После переименования изменились значения формул, ссылавшихся на лист: ${changed.join(", ")}. ` : "") +
        (undoRecorded ? "Переименование можно отменить кнопкой «Отменить»." : "Проверьте книгу."),
        "applied"
      );
    }
    return {
      ok: true,
      executionState: "verified",
      sheet: plan.newName,
      previousName: plan.oldName,
      sheetId: plan.sheetId,
      formulasRewritten: plan.referencing.length,
      note: `Excel сам переписал ссылки на лист; значения ${plan.referencing.length} формул, ссылавшихся на него, сверены — не изменились.`,
      ...(plan.literal.length
        ? { brokenLiteralFormulas: plan.literal, brokenLiteralNote: "В этих формулах прежнее имя листа стоит внутри текста (INDIRECT, HYPERLINK): Excel его не переписал, они показывают #ССЫЛКА! или ведут в никуда. Назови их пользователю." }
        : {}),
      ...(plan.textMentions.length ? { textMentions: plan.textMentions, textMentionsNote: "В этих ячейках прежнее имя упомянуто просто текстом — оно не изменилось." } : {}),
      ...(plan.unscannedSheets.length ? { unscannedSheets: plan.unscannedSheets } : {}),
      undoable: undoRecorded,
      undoNote: plan.undoNote
    };
  });
}

export interface DeleteSheetPlan {
  readonly kind: "delete_sheet";
  readonly id: string;
  readonly workbook: ReturnType<typeof currentWorkbookIdentity>;
  readonly sheetId: string;
  readonly sheetName: string;
  /** Что исчезнет вместе с листом. */
  readonly usedAddress: string | null;
  readonly filledCells: number | null;
  readonly charts: number;
  readonly pivots: number;
  readonly tables: readonly string[];
  readonly brokenNames: readonly string[];
  /** Формулы, которые станут #ССЫЛКА!. */
  readonly referencing: readonly CellMention[];
  readonly literal: readonly CellMention[];
  readonly viaNames: readonly CellMention[];
  readonly overflow: number;
  readonly unscannedSheets: readonly string[];
  readonly refErrorsBefore: number;
  /** Слепок содержимого: правка после предпросмотра останавливает удаление. */
  readonly signature: string;
  readonly backup: { name: string; at: string } | null;
  readonly undoAvailable: false;
  readonly undoNote: string;
  readonly createdAt: string;
}

async function sheetContents(ctx: Excel.RequestContext, sheet: Excel.Worksheet) {
  const used = sheet.getUsedRangeOrNullObject(true);
  used.load(["isNullObject", "address", "rowCount", "columnCount"]);
  const charts = sheet.charts;
  charts.load("items/name");
  const tables = sheet.tables;
  tables.load("items/name");
  let pivotCollection: Excel.PivotTableCollection | null = null;
  try {
    pivotCollection = sheet.pivotTables;
    pivotCollection.load("items/name");
  } catch { pivotCollection = null; }
  await ctx.sync();
  let filled: number | null = null;
  let signature = "пусто";
  if (!used.isNullObject) {
    if (used.rowCount * used.columnCount <= MAX_IO_CELLS) {
      used.load(["formulas"]);
      await ctx.sync();
      const formulas = used.formulas as unknown[][];
      filled = formulas.flat().filter((value) => value !== "" && value !== null).length;
      signature = JSON.stringify([used.address, formulas]);
    } else {
      signature = JSON.stringify([used.address]);
    }
  }
  return {
    usedAddress: used.isNullObject ? null : String(used.address).replace(/^.*!/, ""),
    filledCells: used.isNullObject ? 0 : filled,
    charts: charts.items.length,
    pivots: pivotCollection ? pivotCollection.items.length : 0,
    tables: tables.items.map((item) => item.name),
    signature: `${signature}|${charts.items.length}|${tables.items.length}`
  };
}

export async function prepareDeleteSheetPlan(args: unknown): Promise<DeleteSheetPlan> {
  preflightToolArgs("delete_sheet", args);
  const a = args as { sheet: string };
  const prepared = await Excel.run(async (ctx) => {
    await assertStructureEditable(ctx);
    const { sheet } = await findSheet(ctx, a.sheet);
    const collection = ctx.workbook.worksheets;
    collection.load("items/visibility,items/name");
    await ctx.sync();
    const visible = collection.items.filter((item) => String(item.visibility) === "Visible");
    if (String(sheet.visibility) === "Visible" && visible.length <= 1) {
      throw new ToolError(`«${sheet.name}» — последний видимый лист книги: Excel не удаляет его. Операция не выполнялась.`);
    }
    const contents = await sheetContents(ctx, sheet);
    const brokenNames = await namesOnSheet(ctx, sheet.name);
    const scan = await scanWorkbookFormulas(ctx);
    const impact = sheetImpact(scan.sheets.filter((item) => item.name !== sheet.name), sheet.name, brokenNames, columnLetters);
    const backup = lastWorkbookBackup();
    return {
      kind: "delete_sheet" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      workbook: currentWorkbookIdentity(),
      sheetId: sheet.id,
      sheetName: sheet.name,
      usedAddress: contents.usedAddress,
      filledCells: contents.filledCells,
      charts: contents.charts,
      pivots: contents.pivots,
      tables: contents.tables,
      brokenNames,
      referencing: impact.referencing,
      literal: impact.literal,
      viaNames: impact.viaNames,
      overflow: impact.overflow,
      unscannedSheets: scan.unscanned,
      refErrorsBefore: scan.sheets.filter((item) => item.name !== sheet.name).reduce((total, item) => total + countRefErrors(item.values), 0),
      signature: contents.signature,
      backup: backup ? { name: backup.name, at: backup.at } : null,
      undoAvailable: false as const,
      undoNote: "Отмены нет и не будет: лист со всеми данными и объектами удаляется, ссылки на него становятся #ССЫЛКА!. Восстановить — только из резервной копии.",
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executeDeleteSheetPlan(plan: DeleteSheetPlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItemOrNullObject(plan.sheetId);
    sheet.load(["isNullObject", "name"]);
    await ctx.sync();
    if (sheet.isNullObject) throw new ToolExecutionError(`Листа «${plan.sheetName}» уже нет. Удалять нечего.`, "failed_before_write");
    const contents = await sheetContents(ctx, sheet);
    if (contents.signature !== plan.signature || sheet.name !== plan.sheetName) {
      throw new ToolExecutionError(`Лист «${plan.sheetName}» изменился после предпросмотра. Удаление не выполнялось — сделайте новый предпросмотр.`, "failed_before_write");
    }
    try {
      sheet.delete();
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(`Excel отказал в удалении листа: ${error?.message ?? error}. Перечитайте список листов.`, "unknown");
    }
    const gone = ctx.workbook.worksheets.getItemOrNullObject(plan.sheetId);
    gone.load("isNullObject");
    await ctx.sync();
    if (!gone.isNullObject) throw new ToolExecutionError(`Лист «${plan.sheetName}» после удаления по-прежнему в книге. Проверьте список листов.`, "applied");

    // Ошибки ссылок до и после: новых должно быть не больше предсказанных.
    const scan = await scanWorkbookFormulas(ctx);
    const refErrorsAfter = scan.sheets.reduce((total, item) => total + countRefErrors(item.values), 0);
    const predicted = plan.referencing.length + plan.literal.length + plan.viaNames.length + plan.overflow;
    const newRefErrors = refErrorsAfter - plan.refErrorsBefore;
    const result = {
      ok: true,
      executionState: "verified",
      deletedSheet: plan.sheetName,
      removed: { usedAddress: plan.usedAddress, filledCells: plan.filledCells, charts: plan.charts, pivots: plan.pivots, tables: plan.tables },
      refErrorsBefore: plan.refErrorsBefore,
      refErrorsAfter,
      ...(plan.referencing.length || plan.literal.length || plan.viaNames.length
        ? { brokenFormulas: [...plan.referencing, ...plan.literal, ...plan.viaNames], brokenNote: "Эти формулы ссылались на удалённый лист и теперь показывают #ССЫЛКА!. Перечисли их пользователю." }
        : {}),
      ...(plan.brokenNames.length ? { brokenNames: plan.brokenNames } : {}),
      undoable: false,
      undoNote: plan.undoNote
    };
    if (newRefErrors > predicted) {
      throw new ToolExecutionError(
        `Лист «${plan.sheetName}» удалён, но новых ошибок ссылок в книге ${newRefErrors}, а предсказано ${predicted}: разбор формул нашёл не всё. Проверьте книгу; отмены нет${plan.backup ? `, есть резервная копия ${plan.backup.name}` : ""}.`,
        "applied"
      );
    }
    return result;
  });
}
