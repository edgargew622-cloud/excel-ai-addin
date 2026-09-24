/**
 * Сводная таблица на проверяемом пути.
 *
 * Подготовка считает сводную сама (`pivotModel.ts`): группы, итоги, общий
 * итог и размер. По размеру проверяется, что место под сводной пусто, —
 * иначе Excel молча затёр бы данные или отказал на середине. Исполнение
 * строит сводную и сверяет прочитанные из неё числа с расчётом панели:
 * расхождение значит, что Excel свёл не то, что ожидалось, и об этом
 * говорится, а не молчится.
 */

import { checkSheetName, freeSheetName } from "./sheetRules";
import {
  assertPlanWorkbook,
  checkAddress,
  deepFreeze,
  MAX_IO_CELLS,
  preflightToolArgs,
  probeMergedAreas,
  rangeOf,
  ToolError,
  ToolExecutionError
} from "./excelTools";
import {
  AGGREGATIONS,
  AGGREGATION_TEXT,
  expectPivot,
  fieldIndex,
  OFFICE_AGGREGATION,
  pivotHeaderProblems,
  pivotMismatches,
  type Aggregation,
  type PivotExpectation,
  type PivotValueField
} from "./pivotModel";
import { intersects, parseA1Rect, type A1Rect } from "./a1";
import { columnLetters } from "./formulaFill";
import { placementCell } from "./chartModel";
import { action, getStructuralRevision, isCustomUndoAvailable, push } from "./undo";
import { captureTarget, officeCapabilities, type WorkbookTarget } from "./workbookContext";

export interface CreatePivotPlan {
  readonly kind: "create_pivot_table";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly name: string;
  readonly sourceAddress: string;
  readonly sourceRows: number;
  readonly rowFields: readonly string[];
  readonly valueFields: readonly PivotValueField[];
  readonly destSheet: string;
  /** Пусто, если лист создаётся самой операцией (newSheet). */
  readonly destSheetId: string;
  /** Лист создаётся при исполнении (этап 7, 7.3.4); отмена уберёт его, если он останется пустым. */
  readonly newSheet?: true;
  /** Левый верхний угол и вся область, которую займёт сводная. */
  readonly destCell: string;
  readonly destArea: string;
  readonly expectation: PivotExpectation;
  readonly preview: readonly string[];
  /** Слепок источника — формулы и значения: по ним посчитан расчёт панели. */
  readonly signature: string;
  readonly sourceSameSheet: boolean;
  readonly undoAvailable: boolean;
  readonly undoNote?: string;
  readonly createdAt: string;
}

function withoutSheet(address: string): string {
  return address.slice(address.lastIndexOf("!") + 1);
}

function areaAt(cell: string, height: number, width: number): { rect: A1Rect; address: string } {
  const start = parseA1Rect(cell)!;
  const rect: A1Rect = {
    kind: "cells",
    rowStart: start.rowStart,
    columnStart: start.columnStart,
    rowEnd: start.rowStart + height - 1,
    columnEnd: start.columnStart + width - 1
  };
  const address = `${columnLetters(rect.columnStart)}${rect.rowStart}:${columnLetters(rect.columnEnd)}${rect.rowEnd}`;
  return { rect, address };
}

function parseValueFields(raw: unknown): PivotValueField[] {
  if (!Array.isArray(raw) || !raw.length) throw new ToolError("Нужно хотя бы одно поле в values.");
  return raw.map((item) => {
    if (typeof item === "string") return { field: item, aggregation: "sum" as Aggregation };
    const value = item as { field?: unknown; aggregation?: unknown };
    if (typeof value?.field !== "string" || !value.field.trim()) throw new ToolError("У каждого поля значений нужен field — заголовок столбца.");
    const aggregation = (value.aggregation ?? "sum") as Aggregation;
    if (!AGGREGATIONS.includes(aggregation)) throw new ToolError(`Неизвестная агрегация ${String(value.aggregation)}: доступны ${AGGREGATIONS.join(", ")}.`);
    return { field: value.field, aggregation };
  });
}

/**
 * Сколько ячеек места заняты.
 *
 * Занятость смотрится и по формулам, и по значениям. Формула `=""` даёт
 * пустое значение, и проверка по одним значениям принимала её за свободную
 * ячейку, хотя сводная её затёрла бы (план стабилизации, S2).
 */
function occupiedCells(formulas: unknown[][], values: unknown[][]): number {
  let count = 0;
  formulas.forEach((row, r) => row.forEach((formula, c) => {
    const value = values[r]?.[c];
    if ((formula !== "" && formula !== null && formula !== undefined) || (value !== "" && value !== null && value !== undefined)) count += 1;
  }));
  return count;
}

/**
 * Таблицы листа — без проглатывания ошибки.
 *
 * Общий `readTableRanges` при сбое отдаёт пустой список: для предупреждений
 * этого хватает. Но для места сводной «не смогли прочитать» — не «таблиц нет»,
 * и строить поверх непроверенного места нельзя.
 */
async function readTablesStrict(ctx: Excel.RequestContext, sheet: Excel.Worksheet): Promise<{ name: string; address: string }[]> {
  try {
    const tables = sheet.tables;
    tables.load("items/name");
    await ctx.sync();
    const ranges = tables.items.map((table) => {
      const range = table.getRange();
      range.load("address");
      return { name: table.name, range };
    });
    if (ranges.length) await ctx.sync();
    return ranges.map((item) => ({ name: item.name, address: String(item.range.address) }));
  } catch (error: any) {
    throw new ToolError(`Не удалось прочитать таблицы листа ${sheet.name}: ${error?.message ?? error}. Место под сводной не проверено, и строить на нём нельзя.`);
  }
}

interface DestinationCheck {
  /** Причина, по которой место не годится вовсе; null — годится, если пусто. */
  problem: string | null;
  occupied: number;
}

/**
 * Одно правило места для всех случаев: подготовки, повторной проверки перед
 * созданием и поиска свободного места. Иначе подсказанное «свободное» место
 * могло бы не пройти следующую же проверку.
 */
async function checkDestination(
  ctx: Excel.RequestContext,
  sheet: Excel.Worksheet,
  area: { rect: A1Rect; address: string },
  sourceRect: A1Rect | null
): Promise<DestinationCheck> {
  sheet.load("name");
  sheet.protection.load("protected");
  const place = sheet.getRange(area.address);
  place.load(["address", "rowIndex", "columnIndex", "rowCount", "columnCount", "formulas", "values"]);
  await ctx.sync();
  const occupied = occupiedCells(place.formulas as unknown[][], place.values as unknown[][]);
  const refuse = (problem: string) => ({ problem, occupied });

  if (sheet.protection.protected) return refuse(`Лист ${sheet.name} защищён: сводную на нём не построить. Снимите защиту или выберите другой лист.`);
  if (sourceRect && intersects(area.rect, sourceRect)) return refuse(`Сводная займёт ${area.address} и наложится на источник.`);

  const tables = await readTablesStrict(ctx, sheet);
  const hitTable = tables.find((table) => {
    const rect = parseA1Rect(withoutSheet(table.address));
    return rect && intersects(rect, area.rect);
  });
  if (hitTable) return refuse(`Сводная займёт ${area.address} и заденет таблицу ${hitTable.name} (${hitTable.address}).`);

  const pivots = sheet.pivotTables;
  pivots.load("items/name");
  await ctx.sync();
  const pivotRanges = pivots.items.map((item) => {
    const layoutRange = item.layout.getRange();
    layoutRange.load("address");
    return { name: item.name, range: layoutRange };
  });
  if (pivotRanges.length) await ctx.sync();
  const hitPivot = pivotRanges.find((item) => {
    const rect = parseA1Rect(withoutSheet(String(item.range.address)));
    return rect && intersects(rect, area.rect);
  });
  if (hitPivot) return refuse(`Сводная займёт ${area.address} и заденет сводную ${hitPivot.name}.`);

  // Excel не строит сводную поверх объединённых ячеек. Угол объединения
  // с неизвестными границами внутри места — тоже отказ: доказать, что
  // объединение не заходит в место, нельзя.
  const merged = await probeMergedAreas(ctx, sheet, place);
  const hitMerged = merged.areas.find((address) => {
    const rect = parseA1Rect(withoutSheet(address));
    return rect && intersects(rect, area.rect);
  }) ?? merged.unresolvedAnchors.find((address) => {
    const rect = parseA1Rect(withoutSheet(address));
    return rect && intersects(rect, area.rect);
  });
  if (hitMerged) return refuse(`Сводная займёт ${area.address}, а там объединённые ячейки (${hitMerged}). Excel не строит сводную поверх объединений.`);

  return { problem: null, occupied };
}

/**
 * Первое свободное место под сводную на листе назначения.
 *
 * Сначала правее занятой области, потом под ней: оба места привычны человеку
 * и не режут данные. Пустоту проверяет сам Excel — чтения дешёвые, а гадать
 * по занятой области нельзя, на ней могли остаться одиночные заметки.
 */
async function findFreeCell(
  ctx: Excel.RequestContext,
  sheet: Excel.Worksheet,
  used: Excel.Range,
  expectation: PivotExpectation,
  sourceRect: A1Rect | null
): Promise<string | null> {
  const empty = Boolean((used as any).isNullObject);
  if (empty) return "A1";
  const candidates = [
    { row: used.rowIndex + 1, column: used.columnIndex + used.columnCount + 2 },
    { row: used.rowIndex + used.rowCount + 3, column: used.columnIndex + 1 }
  ];
  for (const candidate of candidates) {
    const cell = `${columnLetters(candidate.column)}${candidate.row}`;
    const check = await checkDestination(ctx, sheet, areaAt(cell, expectation.height, expectation.width), sourceRect);
    if (!check.problem && check.occupied === 0) return cell;
  }
  return null;
}

export async function prepareCreatePivotPlan(args: unknown): Promise<CreatePivotPlan> {
  preflightToolArgs("create_pivot_table", args);
  const a = args as { sheet?: string; sourceAddress: string; destSheet?: string; destAddress?: string; newSheet?: string; rows: string[]; values: unknown[] };
  if (a.newSheet?.trim() && (a.destSheet?.trim() || a.destAddress?.trim())) {
    throw new ToolError("newSheet не сочетается с destSheet и destAddress: на новом листе сводная встаёт в A1.");
  }
  const source = checkAddress(a.sourceAddress);
  if (a.destAddress !== undefined) {
    const cell = parseA1Rect(a.destAddress);
    if (!cell || cell.kind !== "cells" || cell.rowStart !== cell.rowEnd || cell.columnStart !== cell.columnEnd) {
      throw new ToolError(`destAddress должен быть одной ячейкой, например H1; получено «${a.destAddress}».`);
    }
  }
  if (!Array.isArray(a.rows) || !a.rows.length) throw new ToolError("Нужно хотя бы одно поле в rows.");
  const valueFields = parseValueFields(a.values);
  const target = await captureTarget(a.sheet);

  const prepared = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    const range = await rangeOf(ctx, sheet, source);
    range.load(["address", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
    sheet.load(["id", "name"]);
    await ctx.sync();
    const cells = range.rowCount * range.columnCount;
    if (cells > MAX_IO_CELLS) {
      throw new ToolError(`Сводная строится по области до ${MAX_IO_CELLS} ячеек: панель должна посчитать её заранее. В ${range.address} ${cells}.`);
    }
    if (range.rowCount < 2) throw new ToolError("В источнике нужна строка заголовков и хотя бы одна строка данных.");
    range.load(["values", "formulas"]);
    await ctx.sync();
    const values = range.values as unknown[][];

    const headerProblems = pivotHeaderProblems(values[0]);
    if (headerProblems.length) {
      throw new ToolError(`Шапка ${range.address} не годится для сводной: ${headerProblems.join("; ")}. Операция не выполнялась.`);
    }
    const missing = [...a.rows, ...valueFields.map((item) => item.field)].filter((name) => fieldIndex(values[0], name) < 0);
    if (missing.length) {
      throw new ToolError(
        `Нет полей ${missing.map((name) => `«${name}»`).join(", ")}. Заголовки источника: ${values[0].map((value) => `«${String(value)}»`).join(", ")}.`
      );
    }

    const expectation = expectPivot(values, a.rows, valueFields);

    // Новый лист (этап 7, 7.3.4): создаётся при исполнении, сводная — в A1.
    // Место проверять незачем — лист будет пуст; проверяется только имя.
    if (a.newSheet?.trim()) {
      const all = ctx.workbook.worksheets;
      all.load("items/name");
      await ctx.sync();
      const names = all.items.map((item) => item.name);
      let newName: string;
      try {
        newName = checkSheetName(a.newSheet, names);
      } catch (error: any) {
        const message = String(error?.message ?? error);
        throw new ToolError(/уже есть/.test(message) ? `${message} Свободно, например, «${freeSheetName(String(a.newSheet), names)}».` : message);
      }
      const undoNew = isCustomUndoAvailable();
      return {
        kind: "create_pivot_table" as const,
        id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
        target: { ...target, sheetName: sheet.name },
        name: `Сводная_${Date.now().toString(36)}`,
        sourceAddress: withoutSheet(range.address),
        sourceRows: range.rowCount - 1,
        rowFields: [...a.rows],
        valueFields,
        destSheet: newName,
        destSheetId: "",
        newSheet: true as const,
        destCell: "A1",
        destArea: areaAt("A1", expectation.height, expectation.width).address,
        expectation,
        preview: [
          ...expectation.groups.slice(0, 8).map((group) => `${group.label}: ${group.totals.map((value) => Math.round(value * 100) / 100).join(" · ")}`),
          `Общий итог: ${expectation.grandTotals.map((value) => Math.round(value * 100) / 100).join(" · ")}`
        ],
        signature: JSON.stringify({ formulas: range.formulas, values }),
        sourceSameSheet: false,
        undoAvailable: undoNew,
        ...(undoNew ? {} : { undoNote: "Отмена недоступна: монитор изменений Excel не активен." }),
        createdAt: new Date().toISOString()
      };
    }

    // Лист назначения: указанный или тот же.
    const destSheet = a.destSheet?.trim()
      ? ctx.workbook.worksheets.getItemOrNullObject(a.destSheet.trim())
      : sheet;
    destSheet.load(["id", "name", "isNullObject"]);
    await ctx.sync();
    if ((destSheet as any).isNullObject) {
      throw new ToolError(`Листа «${a.destSheet}» нет. Создайте его через create_sheet или укажите существующий лист.`);
    }
    const sameSheet = destSheet.id === sheet.id;

    const used = officeCapabilities().usedRangeOrNull ? destSheet.getUsedRangeOrNullObject(true) : destSheet.getUsedRange(true);
    used.load(["isNullObject", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
    await ctx.sync();
    const emptySheet = Boolean((used as any).isNullObject);
    const destCell = (a.destAddress?.trim().toUpperCase())
      ?? (emptySheet ? "A1" : placementCell({ rowIndex: used.rowIndex, columnIndex: used.columnIndex, columnCount: used.columnCount }, sameSheet ? range.rowIndex : 0));
    const area = areaAt(destCell, expectation.height, expectation.width);

    // Место под сводной обязано быть пустым: Excel не спрашивает, а данные
    // под ней пропадают или операция рвётся на середине.
    const sourceRect = parseA1Rect(withoutSheet(range.address));
    const check = await checkDestination(ctx, destSheet, area, sameSheet ? sourceRect : null);
    if (check.problem) throw new ToolError(`${check.problem} Выберите другое место.`);
    if (check.occupied) {
      // Проверка 20 сентября 2026 года: отказ говорил «укажите свободное
      // место», и агент на этом сдавался, хотя рядом было пусто. Свободное
      // место ищет панель — она и так знает размер будущей сводной.
      const free = await findFreeCell(ctx, destSheet, used, expectation, sameSheet ? sourceRect : null);
      throw new ToolError(
        `Сводная займёт ${destSheet.name}!${area.address}, а там ${check.occupied} непустых ячеек — они были бы затёрты. Операция не выполнялась. ` +
        (free
          ? `Свободно, например, ${destSheet.name}!${free} — повторите с destAddress: "${free}".`
          : "Свободного места такого размера на листе не нашлось: укажите другой лист в destSheet.")
      );
    }

    const preview = [
      ...expectation.groups.slice(0, 8).map((group) => `${group.label}: ${group.totals.map((value) => Math.round(value * 100) / 100).join(" · ")}`),
      `Общий итог: ${expectation.grandTotals.map((value) => Math.round(value * 100) / 100).join(" · ")}`
    ];
    const undo = isCustomUndoAvailable();
    return {
      kind: "create_pivot_table" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      target: { ...target, sheetName: sheet.name },
      name: `Сводная_${Date.now().toString(36)}`,
      sourceAddress: withoutSheet(range.address),
      sourceRows: range.rowCount - 1,
      rowFields: [...a.rows],
      valueFields,
      destSheet: destSheet.name,
      destSheetId: destSheet.id,
      destCell,
      destArea: area.address,
      expectation,
      preview,
      // Подпись по формулам и значениям: формула источника может ссылаться
      // на другой лист, и тогда её текст прежний, а расчёт панели — уже нет.
      signature: JSON.stringify({ formulas: range.formulas, values }),
      sourceSameSheet: sameSheet,
      undoAvailable: undo,
      ...(undo ? {} : { undoNote: "Отмена недоступна: монитор изменений Excel не активен." }),
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

/** Убирает сводную, которую не удалось достроить, и лист, созданный под неё.
 * true — только если повторное чтение книги подтвердило, что их больше нет. */
async function removeUnfinishedPivot(ctx: Excel.RequestContext, name: string, createdSheetId: string | null): Promise<boolean> {
  try {
    const pivot = ctx.workbook.pivotTables.getItemOrNullObject(name);
    pivot.load("isNullObject");
    await ctx.sync();
    if (!pivot.isNullObject) {
      pivot.delete();
      await ctx.sync();
    }
    if (createdSheetId) {
      const created = ctx.workbook.worksheets.getItemOrNullObject(createdSheetId);
      created.load("isNullObject");
      await ctx.sync();
      if (!created.isNullObject) {
        // На листе уже что-то есть — значит, не только наша сводная. Не трогаем.
        const used = created.getUsedRangeOrNullObject(true);
        used.load("isNullObject");
        await ctx.sync();
        if (!used.isNullObject) return false;
        created.delete();
        await ctx.sync();
      }
    }
    const pivotLeft = ctx.workbook.pivotTables.getItemOrNullObject(name);
    pivotLeft.load("isNullObject");
    const sheetLeft = createdSheetId ? ctx.workbook.worksheets.getItemOrNullObject(createdSheetId) : null;
    sheetLeft?.load("isNullObject");
    await ctx.sync();
    return pivotLeft.isNullObject && (!sheetLeft || sheetLeft.isNullObject);
  } catch {
    return false;
  }
}

export async function executeCreatePivotPlan(plan: CreatePivotPlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(plan.target.sheetId);
    const source = sheet.getRange(plan.sourceAddress);
    source.load(["formulas", "values"]);
    await ctx.sync();
    if (JSON.stringify({ formulas: source.formulas, values: source.values }) !== plan.signature) {
      throw new ToolExecutionError(
        `Данные источника ${plan.sourceAddress} изменились после предпросмотра — формулы или их значения. ` +
        "Расчёт групп и итогов устарел. Сводная не строилась — сделайте новый предпросмотр.",
        "failed_before_write"
      );
    }
    let destSheet: Excel.Worksheet;
    let createdSheetId: string | null = null;
    if (plan.newSheet) {
      // Имя могли занять за время подтверждения.
      const all = ctx.workbook.worksheets;
      all.load("items/name");
      await ctx.sync();
      if (all.items.some((item) => item.name.trim().toLowerCase() === plan.destSheet.toLowerCase())) {
        throw new ToolExecutionError(`Лист «${plan.destSheet}» появился после предпросмотра. Сводная не строилась — выберите другое имя.`, "failed_before_write");
      }
      try {
        destSheet = ctx.workbook.worksheets.add(plan.destSheet);
        destSheet.load("id");
        await ctx.sync();
        createdSheetId = destSheet.id;
      } catch (error: any) {
        throw new ToolExecutionError(`Excel отказал в создании листа «${plan.destSheet}»: ${error?.message ?? error}. Неизвестно, появился ли он — посмотрите на книгу.`, "unknown");
      }
    } else {
      destSheet = ctx.workbook.worksheets.getItem(plan.destSheetId);
      // То же правило места, что и при подготовке: за время подтверждения
      // могли появиться данные, таблица, защита или объединение.
      const sourceRect = plan.sourceSameSheet ? parseA1Rect(plan.sourceAddress) : null;
      let check: DestinationCheck;
      try {
        check = await checkDestination(ctx, destSheet, areaAt(plan.destCell, plan.expectation.height, plan.expectation.width), sourceRect);
      } catch (error: any) {
        throw new ToolExecutionError(`${error?.message ?? error} Сводная не строилась.`, "failed_before_write");
      }
      if (check.problem) {
        throw new ToolExecutionError(`${check.problem} Это появилось после предпросмотра. Сводная не строилась.`, "failed_before_write");
      }
      if (check.occupied) {
        throw new ToolExecutionError(`Место ${plan.destSheet}!${plan.destArea} перестало быть пустым после предпросмотра. Сводная не строилась.`, "failed_before_write");
      }
    }

    let pivot: Excel.PivotTable;
    try {
      pivot = destSheet.pivotTables.add(plan.name, source, destSheet.getRange(plan.destCell));
      await ctx.sync();
    } catch (error: any) {
      // Созданный под сводную лист без сводной не нужен: убрать его, если пуст.
      if (createdSheetId) {
        try {
          const orphan = ctx.workbook.worksheets.getItem(createdSheetId);
          const pivots = orphan.pivotTables;
          pivots.load("items/name");
          await ctx.sync();
          if (!pivots.items.length) { orphan.delete(); await ctx.sync(); }
        } catch { /* лист останется — об этом говорит сообщение ниже */ }
      }
      throw new ToolExecutionError(
        `Excel отказал в построении сводной: ${error?.message ?? error}. Неизвестно, успела ли она появиться — посмотрите на лист ${plan.destSheet}.`,
        "unknown"
      );
    }

    try {
      // Макет по умолчанию задаётся в настройках Excel, а размер и сверку
      // панель рассчитывает для табличного с итогами внизу групп: только
      // в нём у каждого уровня свой столбец, и вложенные итоги проверяемы
      // (план стабилизации, S3.1). Выставляется до полей, чтобы сводная
      // ни на каком шаге не была шире рассчитанного места. Оба свойства —
      // ExcelApi 1.8, как и сами сводные.
      pivot.layout.layoutType = "Tabular" as any;
      pivot.layout.subtotalLocation = "AtBottom" as any;
      for (const field of plan.rowFields) pivot.rowHierarchies.add(pivot.hierarchies.getItem(field));
      for (const item of plan.valueFields) {
        const data = pivot.dataHierarchies.add(pivot.hierarchies.getItem(item.field));
        data.summarizeBy = OFFICE_AGGREGATION[item.aggregation] as any;
      }
      await ctx.sync();
    } catch (error: any) {
      const reason = error?.message ?? error;
      if (await removeUnfinishedPivot(ctx, plan.name, createdSheetId)) {
        throw new ToolExecutionError(
          `Excel отказал в добавлении полей сводной: ${reason}. Недостроенная сводная удалена` +
            `${createdSheetId ? ` вместе с созданным под неё листом «${plan.destSheet}»` : ""}; книга в прежнем виде.`,
          "failed_before_write"
        );
      }
      throw new ToolExecutionError(
        `Excel отказал в добавлении полей сводной: ${reason}. Недостроенную сводную ${plan.name} убрать не удалось — посмотрите на лист ${plan.destSheet}.`,
        "unknown"
      );
    }

    let undoRecorded = false;
    if (plan.undoAvailable) {
      const name = plan.name;
      undoRecorded = push(action(`сводная ${name} на листе ${plan.destSheet}`, async () => {
        const revision = getStructuralRevision();
        await Excel.run(async (undoCtx) => {
          const existing = undoCtx.workbook.pivotTables.getItemOrNullObject(name);
          existing.load("isNullObject");
          await undoCtx.sync();
          if (existing.isNullObject) throw new Error("Сводной уже нет: её удалили после операции агента. Отменять нечего.");
          if (getStructuralRevision() !== revision) throw new Error("Структура книги изменилась во время отмены. Отмена остановлена.");
          existing.delete();
          await undoCtx.sync();
          // Лист, созданный под сводную, уходит вместе с ней — если на нём
          // больше ничего нет. Иначе остаётся: там уже чужая работа.
          if (createdSheetId) {
            const created = undoCtx.workbook.worksheets.getItemOrNullObject(createdSheetId);
            created.load("isNullObject");
            await undoCtx.sync();
            if (!created.isNullObject) {
              const used = created.getUsedRangeOrNullObject(true);
              used.load("isNullObject");
              await undoCtx.sync();
              if (used.isNullObject) { created.delete(); await undoCtx.sync(); }
            }
          }
        });
      }));
    }

    const layout = pivot.layout.getRange();
    layout.load(["address", "values"]);
    await ctx.sync();
    const actualArea = withoutSheet(String(layout.address));
    const problems = pivotMismatches(plan.expectation, layout.values as unknown[][]);
    if (actualArea !== plan.destArea) problems.unshift(`заняла ${actualArea} вместо ${plan.destArea}`);
    if (problems.length) {
      throw new ToolExecutionError(
        `Сводная ${plan.name} построена, но расходится с расчётом панели: ${problems.join("; ")}. ` +
        (undoRecorded ? "Её можно убрать кнопкой «Отменить»." : "Проверьте её на листе."),
        "applied"
      );
    }

    return {
      ok: true,
      executionState: "verified",
      pivot: plan.name,
      sheet: plan.destSheet,
      address: actualArea,
      source: plan.sourceAddress,
      rows: plan.rowFields,
      values: plan.valueFields.map((item) => `${item.field} — ${AGGREGATION_TEXT[item.aggregation]}`),
      grandTotals: plan.expectation.grandTotals,
      groups: plan.expectation.groups.length,
      ...(plan.expectation.warnings.length ? { warnings: plan.expectation.warnings } : {}),
      note: plan.rowFields.length > 1
        ? "Итоги всех групп на всех уровнях и общий итог сверены с расчётом панели по исходным данным."
        : "Итоги каждой группы и общий итог сверены с расчётом панели по исходным данным.",
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: plan.undoNote ?? "Автоматическая отмена этой операции недоступна." })
    };
  });
}
