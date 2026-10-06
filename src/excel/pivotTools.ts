/**
 * Сводные «как у профессионала» и вид листа (срез 10.7, 07.10.2026):
 * изменить готовую сводную, обновить сводные, добавить срезы, убрать сетку
 * и скрыть лист. Всё проверено замером в Excel 07.10.2026 (ExcelApi 1.14).
 *
 * Чего Excel надстройкам не даёт (тот же замер): подключить один срез к
 * нескольким сводным («Подключение к отчётам») и закрепить срез от сдвига.
 * Срез отбирает только свою сводную — так и говорится в ответе.
 */

import { assertPlanWorkbook, deepFreeze, preflightToolArgs, ToolError, ToolExecutionError } from "./excelTools";
import { applyPivotFinish, finishText, GRAND_TOTALS, type GrandTotals, type PivotFinish } from "./pivotFinish";
import { AGGREGATIONS, OFFICE_AGGREGATION, type Aggregation } from "./pivotModel";
import { action, isCustomUndoAvailable, push } from "./undo";
import { captureTarget, type WorkbookTarget } from "./workbookContext";

const newId = () => globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`;
const sameName = (a: string, b: string) => a.trim().toLowerCase() === b.trim().toLowerCase();

/** Служебная «иерархия» значений в столбцах — не поле источника. */
const isValuesHierarchy = (name: string) => name === "Значения" || name === "Values" || name === "Σ Значения";

async function findPivot(ctx: Excel.RequestContext, name: string): Promise<Excel.PivotTable> {
  const pivots = ctx.workbook.pivotTables;
  pivots.load("items/name");
  await ctx.sync();
  const found = pivots.items.find((item) => sameName(item.name, name));
  if (!found) {
    const all = pivots.items.map((item) => `«${item.name}»`).join(", ") || "в книге сводных нет";
    throw new ToolError(`Сводной «${name}» нет. Есть: ${all}.`);
  }
  return found;
}

interface PivotShape {
  rows: string[];
  columns: string[];
  values: { name: string; source: string; aggregation: string }[];
  sourceFields: string[];
  showRowGrandTotals: boolean;
  showColumnGrandTotals: boolean;
}

async function readShape(ctx: Excel.RequestContext, pivot: Excel.PivotTable): Promise<PivotShape> {
  const rows = pivot.rowHierarchies;
  const columns = pivot.columnHierarchies;
  const data = pivot.dataHierarchies;
  const all = pivot.hierarchies;
  rows.load("items/name");
  columns.load("items/name");
  data.load("items/name,items/summarizeBy");
  all.load("items/name");
  const layout = pivot.layout;
  layout.load(["showRowGrandTotals", "showColumnGrandTotals"]);
  await ctx.sync();
  const sources = data.items.map((item) => {
    const field = (item as any).field;
    try { field?.load?.("name"); } catch { /* старый Excel */ }
    return field;
  });
  await ctx.sync();
  return {
    rows: rows.items.map((item) => item.name),
    columns: columns.items.map((item) => item.name).filter((name) => !isValuesHierarchy(name)),
    values: data.items.map((item, index) => ({ name: item.name, source: String(sources[index]?.name ?? ""), aggregation: String(item.summarizeBy) })),
    sourceFields: all.items.map((item) => item.name),
    showRowGrandTotals: Boolean(layout.showRowGrandTotals),
    showColumnGrandTotals: Boolean(layout.showColumnGrandTotals)
  };
}

/* --------------------------------------------------------- update_pivot */

export interface UpdatePivotPlan {
  readonly kind: "update_pivot";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly pivot: string;
  readonly sheet: string;
  readonly before: PivotShape;
  readonly rows?: readonly string[];
  readonly columns?: readonly string[];
  readonly values?: readonly { field: string; aggregation: Aggregation }[];
  /** Порядок элементов поля строк по итогу первого поля значений. */
  readonly sort?: { readonly field: string; readonly order: "asc" | "desc" };
  readonly finish: PivotFinish;
  readonly preview: readonly string[];
  readonly undoAvailable: boolean;
  readonly createdAt: string;
}

export async function prepareUpdatePivotPlan(args: unknown): Promise<UpdatePivotPlan> {
  preflightToolArgs("update_pivot", args);
  const a = args as {
    sheet?: string; pivot: string; rows?: string[]; columns?: string[];
    values?: { field: string; aggregation?: string; label?: string; numberFormat?: string }[];
    grandTotals?: GrandTotals; subtotals?: boolean; sort?: { field: string; order: "asc" | "desc" };
  };
  if (a.grandTotals !== undefined && !GRAND_TOTALS.includes(a.grandTotals)) throw new ToolError(`grandTotals — ${GRAND_TOTALS.join(", ")}.`);
  if (a.rows && !a.rows.length) throw new ToolError("rows не может быть пустым: у сводной должно быть хотя бы одно поле строк.");
  const target = await captureTarget(a.sheet);
  const prepared = await Excel.run(async (ctx) => {
    const pivot = await findPivot(ctx, a.pivot);
    const sheet = pivot.worksheet;
    sheet.load("name");
    pivot.load("name");
    const before = await readShape(ctx, pivot);
    const known = (name: string) => before.sourceFields.some((field) => sameName(field, name));
    const unknown = [...(a.rows ?? []), ...(a.columns ?? []), ...(a.values ?? []).map((item) => item.field)].filter((name) => !known(name));
    if (unknown.length) {
      throw new ToolError(`В сводной «${pivot.name}» нет полей ${unknown.map((name) => `«${name}»`).join(", ")}. Поля источника: ${before.sourceFields.map((name) => `«${name}»`).join(", ")}.`);
    }
    const asked = a.values?.map((item) => {
      const aggregation = (item.aggregation ?? "sum") as Aggregation;
      if (!AGGREGATIONS.includes(aggregation)) throw new ToolError(`Агрегация «${item.aggregation}» — одна из ${AGGREGATIONS.join(", ")}.`);
      return { field: item.field, aggregation };
    });
    // Те же поля и агрегации, что уже стоят, — значения не пересоздаются:
    // пересоздание сбрасывает сортировку (живая проверка 10.7, 07.10.2026),
    // а подпись и формат меняются и так.
    const sameValues = asked && asked.length === before.values.length && asked.every((item, index) =>
      sameName(item.field, before.values[index].source || "") && OFFICE_AGGREGATION[item.aggregation].toLowerCase() === before.values[index].aggregation.toLowerCase());
    const values = sameValues ? undefined : asked;
    if (a.sort) {
      const rowsAfter = a.rows ?? before.rows;
      if (!rowsAfter.some((name) => sameName(name, a.sort!.field))) throw new ToolError(`Сортировать можно поле строк: ${rowsAfter.map((name) => `«${name}»`).join(", ")}.`);
      if (a.sort.order !== "asc" && a.sort.order !== "desc") throw new ToolError("sort.order — asc или desc.");
    }
    // Подписи и формат — к полям значений по порядку: новым, если values задан, иначе нынешним.
    const finishValues = (a.values ?? []).flatMap((item, index) => (item.label || item.numberFormat
      ? [{ index, ...(item.label ? { label: item.label } : {}), ...(item.numberFormat ? { numberFormat: item.numberFormat } : {}) }]
      : []));
    const finish: PivotFinish = {
      values: finishValues,
      ...(a.grandTotals ? { grandTotals: a.grandTotals } : {}),
      ...(a.subtotals === false ? { subtotals: false } : {})
    };
    if (!a.rows && !a.columns && !a.values && !a.grandTotals && a.subtotals === undefined && !a.sort) {
      throw new ToolError("Что поменять в сводной? Укажите rows, columns, values, grandTotals, subtotals или sort.");
    }
    const preview = [
      ...(a.rows ? [`Строки: ${before.rows.join(", ") || "—"} → ${a.rows.join(", ")}`] : []),
      ...(a.columns ? [`Столбцы: ${before.columns.join(", ") || "—"} → ${a.columns.join(", ") || "нет"}`] : []),
      ...(values ? [`Значения: ${before.values.map((item) => item.name).join(", ") || "—"} → ${values.map((item) => `${item.field} (${item.aggregation})`).join(", ")}`] : []),
      ...(a.sort ? [`Порядок «${a.sort.field}»: ${a.sort.order === "desc" ? "по убыванию" : "по возрастанию"} итога`] : []),
      ...finishText(finish, (values ?? before.values.map((item) => ({ field: item.name }))).map((item: any) => item.field))
    ];
    return {
      kind: "update_pivot" as const,
      id: newId(),
      target: { ...target, sheetName: sheet.name },
      pivot: pivot.name,
      sheet: sheet.name,
      before,
      ...(a.rows ? { rows: a.rows } : {}),
      ...(a.columns ? { columns: a.columns } : {}),
      ...(values ? { values } : {}),
      ...(a.sort ? { sort: { field: a.sort.field, order: a.sort.order } } : {}),
      finish,
      preview,
      undoAvailable: isCustomUndoAvailable(),
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

async function setFields(ctx: Excel.RequestContext, pivot: Excel.PivotTable, rows?: readonly string[], columns?: readonly string[], values?: readonly { field: string; aggregation: string; office?: boolean }[]) {
  if (values) {
    const data = pivot.dataHierarchies;
    data.load("items/name");
    await ctx.sync();
    for (const item of [...data.items]) pivot.dataHierarchies.remove(item);
    await ctx.sync();
  }
  if (rows) {
    const current = pivot.rowHierarchies;
    current.load("items/name");
    await ctx.sync();
    for (const item of [...current.items]) pivot.rowHierarchies.remove(item);
    await ctx.sync();
    for (const name of rows) pivot.rowHierarchies.add(pivot.hierarchies.getItem(name));
    await ctx.sync();
  }
  if (columns) {
    const current = pivot.columnHierarchies;
    current.load("items/name");
    await ctx.sync();
    for (const item of current.items.filter((hierarchy) => !isValuesHierarchy(hierarchy.name))) pivot.columnHierarchies.remove(item);
    await ctx.sync();
    for (const name of columns) pivot.columnHierarchies.add(pivot.hierarchies.getItem(name));
    await ctx.sync();
  }
  if (values) {
    for (const item of values) {
      const added = pivot.dataHierarchies.add(pivot.hierarchies.getItem(item.field));
      added.summarizeBy = (item.office ? item.aggregation : OFFICE_AGGREGATION[item.aggregation as Aggregation]) as any;
    }
    await ctx.sync();
  }
}

export async function executeUpdatePivotPlan(plan: UpdatePivotPlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const pivot = await findPivot(ctx, plan.pivot);
    const now = await readShape(ctx, pivot);
    if (JSON.stringify(now) !== JSON.stringify(plan.before)) {
      throw new ToolExecutionError(`Сводную «${plan.pivot}» изменили после предпросмотра. Операция не выполнялась — сделайте новый предпросмотр.`, "failed_before_write");
    }
    let finished: Awaited<ReturnType<typeof applyPivotFinish>>;
    try {
      await setFields(ctx, pivot, plan.rows, plan.columns, plan.values);
      finished = await applyPivotFinish(ctx, pivot, plan.finish, plan.before.sourceFields);
      if (plan.sort) {
        const data = pivot.dataHierarchies;
        data.load("items/name");
        await ctx.sync();
        const rows = pivot.rowHierarchies;
        rows.load("items/name");
        await ctx.sync();
        const row = rows.items.find((item) => sameName(item.name, plan.sort!.field))!;
        row.fields.getItem(row.name).sortByValues((plan.sort.order === "desc" ? "Descending" : "Ascending") as any, data.items[0]);
        await ctx.sync();
      }
    } catch (error: any) {
      throw new ToolExecutionError(`Excel отказал в изменении сводной «${plan.pivot}»: ${error?.message ?? error}. Посмотрите на неё — часть изменений могла встать.`, "unknown");
    }
    const after = await readShape(ctx, pivot);
    const problems = [...finished.problems];
    if (plan.rows && JSON.stringify(after.rows.map((x) => x.toLowerCase())) !== JSON.stringify(plan.rows.map((x) => x.toLowerCase()))) problems.push(`строки — ${after.rows.join(", ")}`);
    if (plan.columns && JSON.stringify(after.columns.map((x) => x.toLowerCase())) !== JSON.stringify(plan.columns.map((x) => x.toLowerCase()))) problems.push(`столбцы — ${after.columns.join(", ") || "нет"}`);
    if (plan.values && after.values.length !== plan.values.length) problems.push(`полей значений ${after.values.length} вместо ${plan.values.length}`);

    let undoRecorded = false;
    if (plan.undoAvailable) {
      const before = plan.before;
      undoRecorded = push(action(`изменение сводной ${plan.pivot}`, async () => {
        await Excel.run(async (undoCtx) => {
          const undoPivot = await findPivot(undoCtx, plan.pivot);
          await setFields(undoCtx, undoPivot, before.rows, before.columns,
            before.values.map((item) => ({ field: item.source || item.name, aggregation: item.aggregation, office: true })));
          const layout = undoPivot.layout;
          layout.showRowGrandTotals = before.showRowGrandTotals;
          layout.showColumnGrandTotals = before.showColumnGrandTotals;
          await undoCtx.sync();
          // Подписи полей значений — прежние.
          const data = undoPivot.dataHierarchies;
          data.load("items/name");
          await undoCtx.sync();
          before.values.forEach((item, index) => { if (data.items[index] && data.items[index].name !== item.name) data.items[index].name = item.name; });
          await undoCtx.sync();
        });
      }));
    }
    if (problems.length) {
      throw new ToolExecutionError(`Сводная «${plan.pivot}» изменена, но обратное чтение расходится: ${problems.join("; ")}. ` +
        (undoRecorded ? "Изменение можно отменить кнопкой «Отменить»." : "Проверьте сводную."), "applied");
    }
    const area = pivot.layout.getRange();
    area.load("address,values");
    await ctx.sync();
    return {
      ok: true,
      executionState: "verified",
      pivot: plan.pivot,
      sheet: plan.sheet,
      rows: after.rows,
      columns: after.columns,
      values: after.values.map((item) => item.name),
      finish: finished.applied,
      address: String(area.address).replace(/^.*!/, ""),
      firstRows: (area.values as unknown[][]).slice(0, 6),
      note: "Поля и настройки прочитаны обратно из Excel. Числа сводная считает сама по своему источнику.",
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: "Автоматическая отмена недоступна." })
    };
  });
}

/* -------------------------------------------------------- refresh_pivot */

export interface RefreshPivotPlan {
  readonly kind: "refresh_pivot";
  readonly id: string;
  readonly target: WorkbookTarget;
  /** Пусто — все сводные книги. */
  readonly pivots: readonly string[];
  readonly all: boolean;
  readonly createdAt: string;
}

export async function prepareRefreshPivotPlan(args: unknown): Promise<RefreshPivotPlan> {
  preflightToolArgs("refresh_pivot", args);
  const a = args as { sheet?: string; pivot?: string };
  const target = await captureTarget(a.sheet);
  const prepared = await Excel.run(async (ctx) => {
    const pivots = ctx.workbook.pivotTables;
    pivots.load("items/name");
    await ctx.sync();
    if (!pivots.items.length) throw new ToolError("В книге нет сводных таблиц: обновлять нечего.");
    const names = a.pivot?.trim() ? [(await findPivot(ctx, a.pivot)).name] : pivots.items.map((item) => item.name);
    return { kind: "refresh_pivot" as const, id: newId(), target, pivots: names, all: !a.pivot?.trim(), createdAt: new Date().toISOString() };
  });
  return deepFreeze(prepared);
}

export async function executeRefreshPivotPlan(plan: RefreshPivotPlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const read = async () => {
      const out: Record<string, string> = {};
      for (const name of plan.pivots) {
        const range = (await findPivot(ctx, name)).layout.getRange();
        range.load("address");
        await ctx.sync();
        out[name] = String(range.address).replace(/^.*!/, "");
      }
      return out;
    };
    const before = await read();
    try {
      if (plan.all) ctx.workbook.pivotTables.refreshAll();
      else for (const name of plan.pivots) (await findPivot(ctx, name)).refresh();
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(`Excel отказал в обновлении: ${error?.message ?? error}.`, "unknown");
    }
    const after = await read();
    return {
      ok: true,
      executionState: "verified",
      refreshed: plan.pivots.map((name) => ({ pivot: name, before: before[name], after: after[name] })),
      note: "Сводные пересчитаны по своим источникам. Сводная по таблице Excel подхватывает новые строки; по адресу — только строки внутри адреса.",
      undoable: false,
      undoNote: "Обновление не отменяется: сводная просто показывает текущие данные источника."
    };
  });
}

/* ----------------------------------------------------------- add_slicer */

export interface AddSlicerPlan {
  readonly kind: "add_slicer";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly pivot: string;
  readonly fields: readonly string[];
  readonly destSheet: string;
  readonly anchorCell?: string;
  readonly undoAvailable: boolean;
  readonly createdAt: string;
}

export async function prepareAddSlicerPlan(args: unknown): Promise<AddSlicerPlan> {
  preflightToolArgs("add_slicer", args);
  const a = args as { sheet?: string; pivot: string; fields: string[]; destSheet?: string; anchorCell?: string };
  if (!Array.isArray(a.fields) || !a.fields.length) throw new ToolError("Укажите поля для срезов (fields).");
  const target = await captureTarget(a.sheet);
  const prepared = await Excel.run(async (ctx) => {
    const pivot = await findPivot(ctx, a.pivot);
    const shape = await readShape(ctx, pivot);
    const unknown = a.fields.filter((name) => !shape.sourceFields.some((field) => sameName(field, name)));
    if (unknown.length) throw new ToolError(`В сводной «${pivot.name}» нет полей ${unknown.map((n) => `«${n}»`).join(", ")}. Поля: ${shape.sourceFields.map((n) => `«${n}»`).join(", ")}.`);
    const own = pivot.worksheet;
    own.load("name");
    await ctx.sync();
    const dest = a.destSheet?.trim() ? ctx.workbook.worksheets.getItemOrNullObject(a.destSheet.trim()) : own;
    dest.load(["name", "isNullObject"]);
    await ctx.sync();
    if ((dest as any).isNullObject) throw new ToolError(`Листа «${a.destSheet}» нет. Создайте его через create_sheet.`);
    const slicers = ctx.workbook.slicers;
    slicers.load("items/name");
    await ctx.sync();
    const taken = a.fields.filter((name) => slicers.items.some((item) => sameName(item.name, name)));
    if (taken.length) throw new ToolError(`Срез ${taken.map((n) => `«${n}»`).join(", ")} в книге уже есть.`);
    return {
      kind: "add_slicer" as const,
      id: newId(),
      target: { ...target, sheetName: own.name },
      pivot: pivot.name,
      fields: shape.sourceFields.filter((field) => a.fields.some((name) => sameName(field, name))),
      destSheet: dest.name,
      ...(a.anchorCell?.trim() ? { anchorCell: a.anchorCell.trim().toUpperCase() } : {}),
      undoAvailable: isCustomUndoAvailable(),
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executeAddSlicerPlan(plan: AddSlicerPlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const pivot = await findPivot(ctx, plan.pivot);
    const dest = ctx.workbook.worksheets.getItem(plan.destSheet);
    let left = 0;
    let top = 0;
    if (plan.anchorCell) {
      const cell = dest.getRange(plan.anchorCell);
      cell.load(["left", "top"]);
      await ctx.sync();
      left = cell.left;
      top = cell.top;
    }
    const created: string[] = [];
    try {
      for (const [index, field] of plan.fields.entries()) {
        const slicer = ctx.workbook.slicers.add(pivot, field, dest);
        slicer.caption = field;
        if (plan.anchorCell) { slicer.left = left + index * 152; slicer.top = top; }
        slicer.load("name");
        await ctx.sync();
        created.push(slicer.name);
      }
    } catch (error: any) {
      throw new ToolExecutionError(`Excel отказал в создании среза: ${error?.message ?? error}. Создано: ${created.join(", ") || "ничего"}.`, created.length ? "applied" : "unknown");
    }
    const all = ctx.workbook.slicers;
    all.load("items/name");
    await ctx.sync();
    const missing = created.filter((name) => !all.items.some((item) => item.name === name));
    let undoRecorded = false;
    if (plan.undoAvailable) {
      undoRecorded = push(action(`срезы ${created.join(", ")}`, async () => {
        await Excel.run(async (undoCtx) => {
          for (const name of created) undoCtx.workbook.slicers.getItemOrNullObject(name).delete();
          await undoCtx.sync();
        });
      }));
    }
    if (missing.length) throw new ToolExecutionError(`Срезы ${missing.join(", ")} после создания не найдены.`, "applied");
    return {
      ok: true,
      executionState: "verified",
      slicers: created,
      pivot: plan.pivot,
      sheet: plan.destSheet,
      note: "Срез отбирает только свою сводную: подключить его к нескольким сводным Excel надстройкам не даёт — это делается правой кнопкой по срезу → «Подключение к отчётам». Закрепить срез от сдвига тоже можно только вручную (Формат среза → Свойства).",
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: "Автоматическая отмена недоступна: срез удаляется вручную." })
    };
  });
}

/* ------------------------------------------------------- set_sheet_view */

export interface SheetViewPlan {
  readonly kind: "set_sheet_view";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly sheet: string;
  readonly gridlines?: boolean;
  readonly visible?: boolean;
  readonly before: { gridlines: boolean; visibility: string };
  readonly undoAvailable: boolean;
  readonly createdAt: string;
}

export async function prepareSheetViewPlan(args: unknown): Promise<SheetViewPlan> {
  preflightToolArgs("set_sheet_view", args);
  const a = args as { sheet: string; gridlines?: boolean; visible?: boolean };
  if (a.gridlines === undefined && a.visible === undefined) throw new ToolError("Укажите gridlines или visible.");
  const target = await captureTarget(a.sheet);
  const prepared = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    sheet.load(["name", "showGridlines", "visibility"]);
    const all = ctx.workbook.worksheets;
    all.load("items/name,items/visibility");
    await ctx.sync();
    if (a.visible === false) {
      const visible = all.items.filter((item) => String(item.visibility) === "Visible");
      if (visible.length <= 1) throw new ToolError("Это единственный видимый лист: Excel не даст скрыть все листы.");
    }
    return {
      kind: "set_sheet_view" as const,
      id: newId(),
      target: { ...target, sheetName: sheet.name },
      sheet: sheet.name,
      ...(a.gridlines !== undefined ? { gridlines: a.gridlines } : {}),
      ...(a.visible !== undefined ? { visible: a.visible } : {}),
      before: { gridlines: Boolean(sheet.showGridlines), visibility: String(sheet.visibility) },
      undoAvailable: isCustomUndoAvailable(),
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executeSheetViewPlan(plan: SheetViewPlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(plan.target.sheetId);
    try {
      if (plan.gridlines !== undefined) sheet.showGridlines = plan.gridlines;
      if (plan.visible !== undefined) sheet.visibility = (plan.visible ? "Visible" : "Hidden") as any;
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(`Excel отказал: ${error?.message ?? error}.`, "unknown");
    }
    sheet.load(["showGridlines", "visibility"]);
    await ctx.sync();
    const problems: string[] = [];
    if (plan.gridlines !== undefined && Boolean(sheet.showGridlines) !== plan.gridlines) problems.push("сетка не переключилась");
    if (plan.visible !== undefined && (String(sheet.visibility) === "Visible") !== plan.visible) problems.push("видимость листа не переключилась");
    let undoRecorded = false;
    if (plan.undoAvailable) {
      undoRecorded = push(action(`вид листа ${plan.sheet}`, async () => {
        await Excel.run(async (undoCtx) => {
          const undoSheet = undoCtx.workbook.worksheets.getItem(plan.target.sheetId);
          undoSheet.showGridlines = plan.before.gridlines;
          undoSheet.visibility = plan.before.visibility as any;
          await undoCtx.sync();
        });
      }));
    }
    if (problems.length) throw new ToolExecutionError(`Лист ${plan.sheet}: ${problems.join("; ")}.`, "applied");
    return {
      ok: true,
      executionState: "verified",
      sheet: plan.sheet,
      gridlines: Boolean(sheet.showGridlines),
      visible: String(sheet.visibility) === "Visible",
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: "Автоматическая отмена недоступна." })
    };
  });
}
