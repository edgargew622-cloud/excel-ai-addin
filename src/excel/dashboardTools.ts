/**
 * Дашборд (срез 10.8, 07.10.2026): изменить готовую диаграмму, расставить
 * диаграммы сеткой по ячейкам и поставить один фильтр во все сводные.
 *
 * Замер в Excel 07.10.2026: тип, заголовок из ячейки (title.setFormula),
 * кнопки полей сводной диаграммы (pivotOptions) и положение по ячейкам
 * (setPosition) работают. Подключить срез к нескольким сводным Excel
 * надстройкам не даёт — поэтому filter_pivots ставит один и тот же отбор
 * во все сводные с этим полем: «покажи 2025 год во всех графиках».
 */

import { assertPlanWorkbook, deepFreeze, preflightToolArgs, ToolError, ToolExecutionError } from "./excelTools";
import { CHART_KINDS, type ChartKind } from "./chartModel";
import { chartNumberFormat } from "./pivotFinish";
import { columnLetters } from "./formulaFill";
import { action, isCustomUndoAvailable, push } from "./undo";
import { captureTarget, type WorkbookTarget } from "./workbookContext";

const newId = () => globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`;
const sameName = (a: string, b: string) => a.trim().toLowerCase() === b.trim().toLowerCase();
const LEGEND = ["Top", "Bottom", "Left", "Right", "None"] as const;
type Legend = typeof LEGEND[number];

async function findChart(ctx: Excel.RequestContext, sheet: Excel.Worksheet, name: string): Promise<Excel.Chart> {
  const charts = sheet.charts;
  charts.load("items/name");
  sheet.load("name");
  await ctx.sync();
  const found = charts.items.find((item) => sameName(item.name, name));
  if (!found) throw new ToolError(`На листе ${sheet.name} нет диаграммы «${name}». Есть: ${charts.items.map((item) => `«${item.name}»`).join(", ") || "диаграмм нет"}.`);
  return found;
}

/** Ячейка заголовка → формула «='Лист'!$B$1». */
export function titleCellFormula(input: string, defaultSheet: string): string {
  const text = input.trim().replace(/^=/, "");
  const bang = text.lastIndexOf("!");
  const cell = (bang >= 0 ? text.slice(bang + 1) : text).replace(/\$/g, "").toUpperCase();
  if (!/^[A-Z]{1,3}\d{1,7}$/.test(cell)) throw new ToolError(`Заголовок из ячейки — одна ячейка, например «Дашборд!B1»; получено «${input}».`);
  const sheet = bang >= 0 ? text.slice(0, bang).replace(/^'(.*)'$/, "$1") : defaultSheet;
  const column = cell.match(/^[A-Z]+/)![0];
  return `='${sheet.replace(/'/g, "''")}'!$${column}$${cell.slice(column.length)}`;
}

/* ----------------------------------------------------------- edit_chart */

interface ChartState {
  chartType: string;
  title: string;
  titleVisible: boolean;
  legendVisible: boolean;
  legendPosition: string;
  width: number;
  height: number;
  fieldButtons: boolean | null;
}

async function readChart(ctx: Excel.RequestContext, chart: Excel.Chart): Promise<ChartState> {
  chart.load(["chartType", "width", "height"]);
  chart.title.load(["text", "visible"]);
  chart.legend.load(["visible", "position"]);
  await ctx.sync();
  let fieldButtons: boolean | null = null;
  try {
    const options = (chart as any).pivotOptions;
    options.load("showAxisFieldButtons");
    await ctx.sync();
    fieldButtons = Boolean(options.showAxisFieldButtons);
  } catch { fieldButtons = null; }
  return {
    chartType: String(chart.chartType),
    title: String(chart.title.text ?? ""),
    titleVisible: Boolean(chart.title.visible),
    legendVisible: Boolean(chart.legend.visible),
    legendPosition: String(chart.legend.position),
    width: chart.width,
    height: chart.height,
    fieldButtons
  };
}

export interface EditChartPlan {
  readonly kind: "edit_chart";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly chart: string;
  readonly sheet: string;
  readonly before: ChartState;
  readonly chartType?: ChartKind;
  readonly title?: string;
  readonly titleFormula?: string;
  readonly legend?: Legend;
  readonly dataLabels?: boolean;
  readonly fieldButtons?: boolean;
  /** Формат чисел оси значений и подписей, английская запись. */
  readonly numberFormat?: string;
  readonly preview: readonly string[];
  readonly undoAvailable: boolean;
  readonly createdAt: string;
}

export async function prepareEditChartPlan(args: unknown): Promise<EditChartPlan> {
  preflightToolArgs("edit_chart", args);
  const a = args as { sheet?: string; chart: string; chartType?: string; title?: string; titleFromCell?: string; legend?: string; dataLabels?: boolean; fieldButtons?: boolean; numberFormat?: string };
  if (a.chartType !== undefined && !CHART_KINDS.includes(a.chartType as ChartKind)) throw new ToolError(`Тип диаграммы — один из ${CHART_KINDS.join(", ")}.`);
  if (a.legend !== undefined && !LEGEND.includes(a.legend as Legend)) throw new ToolError(`legend — ${LEGEND.join(", ")}.`);
  if (a.title !== undefined && a.titleFromCell !== undefined) throw new ToolError("title и titleFromCell вместе не задаются: заголовок либо текстом, либо из ячейки.");
  if ([a.chartType, a.title, a.titleFromCell, a.legend, a.dataLabels, a.fieldButtons, a.numberFormat].every((item) => item === undefined)) {
    throw new ToolError("Что изменить в диаграмме? chartType, title, titleFromCell, legend, dataLabels, fieldButtons или numberFormat.");
  }
  const target = await captureTarget(a.sheet);
  const prepared = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    const chart = await findChart(ctx, sheet, a.chart);
    const before = await readChart(ctx, chart);
    if (a.fieldButtons !== undefined && before.fieldButtons === null) throw new ToolError(`«${chart.name}» — не сводная диаграмма: кнопок полей у неё нет.`);
    const titleFormula = a.titleFromCell ? titleCellFormula(a.titleFromCell, sheet.name) : undefined;
    const preview = [
      ...(a.chartType ? [`Тип: ${before.chartType} → ${a.chartType}`] : []),
      ...(a.title !== undefined ? [`Заголовок: «${a.title}»`] : []),
      ...(titleFormula ? [`Заголовок из ячейки ${titleFormula.slice(1)} — будет меняться вместе с ней`] : []),
      ...(a.legend ? [a.legend === "None" ? "Без легенды" : `Легенда: ${a.legend}`] : []),
      ...(a.dataLabels !== undefined ? [a.dataLabels ? "Подписи значений" : "Без подписей значений"] : []),
      ...(a.fieldButtons !== undefined ? [a.fieldButtons ? "Показать кнопки полей" : "Скрыть серые кнопки полей сводной"] : []),
      ...(a.numberFormat ? [`Формат чисел оси и подписей: ${a.numberFormat}`] : [])
    ];
    return {
      kind: "edit_chart" as const,
      id: newId(),
      target: { ...target, sheetName: sheet.name },
      chart: chart.name,
      sheet: sheet.name,
      before,
      ...(a.chartType ? { chartType: a.chartType as ChartKind } : {}),
      ...(a.title !== undefined ? { title: a.title } : {}),
      ...(titleFormula ? { titleFormula } : {}),
      ...(a.legend ? { legend: a.legend as Legend } : {}),
      ...(a.dataLabels !== undefined ? { dataLabels: a.dataLabels } : {}),
      ...(a.fieldButtons !== undefined ? { fieldButtons: a.fieldButtons } : {}),
      ...(a.numberFormat?.trim() ? { numberFormat: a.numberFormat.trim() } : {}),
      preview,
      undoAvailable: isCustomUndoAvailable(),
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

function setFieldButtons(chart: Excel.Chart, show: boolean) {
  const options = (chart as any).pivotOptions;
  options.showAxisFieldButtons = show;
  options.showLegendFieldButtons = show;
  options.showReportFilterFieldButtons = show;
  options.showValueFieldButtons = show;
}

export async function executeEditChartPlan(plan: EditChartPlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(plan.target.sheetId);
    const chart = await findChart(ctx, sheet, plan.chart);
    const now = await readChart(ctx, chart);
    if (JSON.stringify(now) !== JSON.stringify(plan.before)) {
      throw new ToolExecutionError(`Диаграмму «${plan.chart}» изменили после предпросмотра. Операция не выполнялась.`, "failed_before_write");
    }
    try {
      if (plan.chartType) chart.chartType = plan.chartType as any;
      if (plan.title !== undefined) { chart.title.text = plan.title; chart.title.visible = plan.title !== ""; }
      if (plan.titleFormula) { chart.title.setFormula(plan.titleFormula); chart.title.visible = true; }
      if (plan.legend) {
        chart.legend.visible = plan.legend !== "None";
        if (plan.legend !== "None") chart.legend.position = plan.legend as any;
      }
      if (plan.dataLabels !== undefined) chart.dataLabels.showValue = plan.dataLabels;
      if (plan.fieldButtons !== undefined) setFieldButtons(chart, plan.fieldButtons);
      await ctx.sync();
      if (plan.numberFormat) {
        const local = await chartNumberFormat(ctx, plan.numberFormat);
        // У кольцевой и круговой оси нет — формат только у подписей.
        if (!["Pie", "Doughnut"].includes(plan.chartType ?? plan.before.chartType)) chart.axes.valueAxis.numberFormat = local;
        chart.dataLabels.numberFormat = local;
        await ctx.sync();
      }
    } catch (error: any) {
      throw new ToolExecutionError(`Excel отказал в изменении диаграммы «${plan.chart}»: ${error?.message ?? error}. Часть изменений могла встать — посмотрите на неё.`, "unknown");
    }
    const after = await readChart(ctx, chart);
    const problems: string[] = [];
    if (plan.chartType && after.chartType !== plan.chartType) problems.push(`тип — ${after.chartType}`);
    if (plan.title !== undefined && after.title !== plan.title) problems.push(`заголовок — «${after.title}»`);
    if (plan.legend && (after.legendVisible !== (plan.legend !== "None") || (plan.legend !== "None" && after.legendPosition !== plan.legend))) problems.push("легенда не переключилась");
    if (plan.fieldButtons !== undefined && after.fieldButtons !== plan.fieldButtons) problems.push("кнопки полей не переключились");

    let undoRecorded = false;
    if (plan.undoAvailable) {
      const before = plan.before;
      undoRecorded = push(action(`диаграмма ${plan.chart}`, async () => {
        await Excel.run(async (undoCtx) => {
          const undoChart = await findChart(undoCtx, undoCtx.workbook.worksheets.getItem(plan.target.sheetId), plan.chart);
          undoChart.chartType = before.chartType as any;
          undoChart.title.text = before.title;
          undoChart.title.visible = before.titleVisible;
          undoChart.legend.visible = before.legendVisible;
          if (before.legendVisible) undoChart.legend.position = before.legendPosition as any;
          if (before.fieldButtons !== null && plan.fieldButtons !== undefined) setFieldButtons(undoChart, before.fieldButtons);
          await undoCtx.sync();
        });
      }));
    }
    if (problems.length) {
      throw new ToolExecutionError(`Диаграмма «${plan.chart}» изменена, но обратное чтение расходится: ${problems.join("; ")}.`, "applied");
    }
    return {
      ok: true,
      executionState: "verified",
      chart: plan.chart,
      sheet: plan.sheet,
      after: { chartType: after.chartType, title: after.title, legend: after.legendVisible ? after.legendPosition : "None", fieldButtons: after.fieldButtons },
      ...(plan.titleFormula ? { titleNote: `Заголовок связан с ${plan.titleFormula.slice(1)}: поменяйте ячейку — поменяется и он.` } : {}),
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: "Автоматическая отмена недоступна." })
    };
  });
}

/* ------------------------------------------------------- arrange_charts */

export interface ArrangeChartsPlan {
  readonly kind: "arrange_charts";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly sheet: string;
  readonly charts: readonly string[];
  /** Ячейки каждой диаграммы: левая верхняя и правая нижняя. */
  readonly cells: readonly { chart: string; from: string; to: string }[];
  readonly before: readonly { chart: string; left: number; top: number; width: number; height: number }[];
  readonly undoAvailable: boolean;
  readonly createdAt: string;
}

/** Сетка диаграмм по ячейкам: столбцы, размер в ячейках, промежуток. */
export function chartGrid(count: number, start: { row: number; column: number }, columns: number, widthCells: number, heightRows: number, gap: number) {
  return Array.from({ length: count }, (_, index) => {
    const r = Math.floor(index / columns);
    const c = index % columns;
    const row = start.row + r * (heightRows + gap);
    const column = start.column + c * (widthCells + gap);
    return { from: `${columnLetters(column)}${row}`, to: `${columnLetters(column + widthCells - 1)}${row + heightRows - 1}` };
  });
}

export async function prepareArrangeChartsPlan(args: unknown): Promise<ArrangeChartsPlan> {
  preflightToolArgs("arrange_charts", args);
  const a = args as { sheet?: string; charts?: string[]; columns?: number; startCell?: string; widthCells?: number; heightRows?: number; gapCells?: number };
  const columns = a.columns ?? 2;
  const widthCells = a.widthCells ?? 8;
  const heightRows = a.heightRows ?? 15;
  const gap = a.gapCells ?? 1;
  const start = /^([A-Za-z]{1,3})(\d{1,7})$/.exec((a.startCell ?? "B2").trim());
  if (!start) throw new ToolError(`startCell — одна ячейка, например B2; получено «${a.startCell}».`);
  const startPos = { column: [...start[1].toUpperCase()].reduce((total, ch) => total * 26 + ch.charCodeAt(0) - 64, 0), row: Number(start[2]) };
  const target = await captureTarget(a.sheet);
  const prepared = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    sheet.load("name");
    const all = sheet.charts;
    all.load("items/name,items/left,items/top,items/width,items/height");
    await ctx.sync();
    if (!all.items.length) throw new ToolError(`На листе ${sheet.name} нет диаграмм.`);
    const chosen = a.charts?.length
      ? a.charts.map((name) => {
        const found = all.items.find((item) => sameName(item.name, name));
        if (!found) throw new ToolError(`На листе ${sheet.name} нет диаграммы «${name}». Есть: ${all.items.map((item) => `«${item.name}»`).join(", ")}.`);
        return found;
      })
      : [...all.items].sort((x, y) => x.top - y.top || x.left - y.left);
    const grid = chartGrid(chosen.length, startPos, columns, widthCells, heightRows, gap);
    return {
      kind: "arrange_charts" as const,
      id: newId(),
      target: { ...target, sheetName: sheet.name },
      sheet: sheet.name,
      charts: chosen.map((item) => item.name),
      cells: chosen.map((item, index) => ({ chart: item.name, ...grid[index] })),
      before: chosen.map((item) => ({ chart: item.name, left: item.left, top: item.top, width: item.width, height: item.height })),
      undoAvailable: isCustomUndoAvailable(),
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executeArrangeChartsPlan(plan: ArrangeChartsPlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(plan.target.sheetId);
    // По очереди: параллельные загрузки одной коллекции в одном контексте Excel
    // не дожидается («items недоступно» — живая проверка 10.8).
    const charts: Excel.Chart[] = [];
    for (const item of plan.cells) charts.push(await findChart(ctx, sheet, item.chart));
    try {
      plan.cells.forEach((item, index) => charts[index].setPosition(item.from, item.to));
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(`Excel отказал в расстановке диаграмм: ${error?.message ?? error}.`, "unknown");
    }
    // Проверка: левый верхний угол каждой диаграммы — на углу своей ячейки.
    const corners = plan.cells.map((item) => { const cell = sheet.getRange(item.from); cell.load(["left", "top"]); return cell; });
    for (const chart of charts) chart.load(["left", "top", "width", "height"]);
    await ctx.sync();
    const off = plan.cells.filter((item, index) => Math.abs(charts[index].left - corners[index].left) > 1 || Math.abs(charts[index].top - corners[index].top) > 1).map((item) => item.chart);
    let undoRecorded = false;
    if (plan.undoAvailable) {
      undoRecorded = push(action(`расстановка диаграмм на ${plan.sheet}`, async () => {
        await Excel.run(async (undoCtx) => {
          const undoSheet = undoCtx.workbook.worksheets.getItem(plan.target.sheetId);
          for (const item of plan.before) {
            const chart = await findChart(undoCtx, undoSheet, item.chart);
            chart.left = item.left; chart.top = item.top; chart.width = item.width; chart.height = item.height;
          }
          await undoCtx.sync();
        });
      }));
    }
    if (off.length) throw new ToolExecutionError(`Диаграммы ${off.join(", ")} встали не по ячейкам сетки.`, "applied");
    return {
      ok: true,
      executionState: "verified",
      sheet: plan.sheet,
      placed: plan.cells.map((item) => `${item.chart}: ${item.from}:${item.to}`),
      note: "Диаграммы стоят по границам ячеек, одного размера. Закрепить их от сдвига надстройка не может: Формат диаграммы → Свойства → «Не перемещать и не изменять размеры».",
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: "Автоматическая отмена недоступна." })
    };
  });
}

/* -------------------------------------------------------- filter_pivots */

export interface FilterPivotsPlan {
  readonly kind: "filter_pivots";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly field: string;
  readonly include?: readonly string[];
  readonly clear: boolean;
  /** Сводные с этим полем и где оно у них стоит. */
  readonly pivots: readonly { name: string; axis: "rows" | "columns" | "filters" | "none" }[];
  readonly createdAt: string;
}

async function fieldPlace(ctx: Excel.RequestContext, pivot: Excel.PivotTable, field: string): Promise<"rows" | "columns" | "filters" | "none" | null> {
  const lists = { rows: pivot.rowHierarchies, columns: pivot.columnHierarchies, filters: pivot.filterHierarchies, all: pivot.hierarchies };
  for (const list of Object.values(lists)) list.load("items/name");
  await ctx.sync();
  if (!lists.all.items.some((item) => sameName(item.name, field))) return null;
  for (const axis of ["rows", "columns", "filters"] as const) if (lists[axis].items.some((item) => sameName(item.name, field))) return axis;
  return "none";
}

export async function prepareFilterPivotsPlan(args: unknown): Promise<FilterPivotsPlan> {
  preflightToolArgs("filter_pivots", args);
  const a = args as { sheet?: string; field: string; include?: string[]; clear?: boolean };
  if (!a.clear && !a.include?.length) throw new ToolError("Укажите include — значения, которые оставить, или clear: true, чтобы снять отбор.");
  const target = await captureTarget(a.sheet);
  const prepared = await Excel.run(async (ctx) => {
    const pivots = ctx.workbook.pivotTables;
    pivots.load("items/name");
    await ctx.sync();
    const places: { name: string; axis: "rows" | "columns" | "filters" | "none" }[] = [];
    for (const pivot of pivots.items) {
      const axis = await fieldPlace(ctx, pivot, a.field);
      if (axis) places.push({ name: pivot.name, axis });
    }
    if (!places.length) throw new ToolError(`Ни в одной сводной книги нет поля «${a.field}».`);
    return {
      kind: "filter_pivots" as const,
      id: newId(),
      target,
      field: a.field,
      ...(a.include?.length ? { include: a.include.map(String) } : {}),
      clear: a.clear === true,
      pivots: places,
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executeFilterPivotsPlan(plan: FilterPivotsPlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const done: string[] = [];
    const problems: string[] = [];
    for (const item of plan.pivots) {
      const pivot = ctx.workbook.pivotTables.getItem(item.name);
      const field = plan.pivots.length ? plan.field : plan.field;
      try {
        let axisList: Excel.RowColumnPivotHierarchyCollection | Excel.FilterPivotHierarchyCollection;
        if (item.axis === "none") {
          if (plan.clear) continue;
          // Поля нет ни в строках, ни в столбцах — оно встаёт в фильтры сводной.
          pivot.filterHierarchies.add(pivot.hierarchies.getItem(field));
          await ctx.sync();
          axisList = pivot.filterHierarchies;
        } else {
          axisList = item.axis === "rows" ? pivot.rowHierarchies : item.axis === "columns" ? pivot.columnHierarchies : pivot.filterHierarchies;
        }
        const hierarchies = axisList as any;
        hierarchies.load("items/name");
        await ctx.sync();
        const hierarchy = hierarchies.items.find((h: any) => sameName(h.name, field));
        const pivotField = hierarchy.fields.getItem(hierarchy.name);
        if (plan.clear) pivotField.clearAllFilters();
        else pivotField.applyFilter({ manualFilter: { selectedItems: [...plan.include!] } } as any);
        await ctx.sync();
        done.push(item.name);
      } catch (error: any) {
        problems.push(`${item.name}: ${error?.message ?? error}`);
      }
    }
    if (problems.length) {
      throw new ToolExecutionError(`Отбор встал не во все сводные: ${problems.join("; ")}. Готово: ${done.join(", ") || "ни одной"}.`, done.length ? "applied" : "unknown");
    }
    return {
      ok: true,
      executionState: "applied",
      field: plan.field,
      ...(plan.clear ? { cleared: done } : { include: plan.include, filtered: done }),
      note: "Один и тот же отбор поставлен во все сводные с этим полем — так один «срез» управляет всеми графиками. Снять: тот же инструмент с clear: true.",
      undoable: false,
      undoNote: "Отбор снимается этим же инструментом с clear: true."
    };
  });
}
