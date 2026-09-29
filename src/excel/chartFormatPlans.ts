/**
 * Перекрасить готовую диаграмму (этап 10, 10.2): цвет рядов и отдельных точек.
 *
 * Диаграмма ищется по имени на листе (его показывает обзор листа). Ряды,
 * их типы, число точек и подписи категорий читаются у самой диаграммы, и
 * просьба сводится к ним до карточки. Прежний цвет линий и маркеров
 * читается до записи — отмена возвращает его. Заливку Excel 2021 не даёт
 * прочитать (chartColors.ts): прежний цвет неизвестен, поэтому, если среди
 * целей есть заливка, отмены нет — карточка говорит это до подтверждения.
 */

import { assertPlanWorkbook, deepFreeze, preflightToolArgs, ToolError, ToolExecutionError } from "./excelTools";
import {
  applyChartColors,
  describeColor,
  FILL_UNVERIFIABLE_NOTE,
  paintStyle,
  resolveColors,
  type ColorRequest,
  type PreviousColor,
  type ResolvedColor
} from "./chartColors";
import { action, isCustomUndoAvailable, push } from "./undo";
import { captureTarget, type WorkbookTarget } from "./workbookContext";

export interface FormatChartPlan {
  readonly kind: "format_chart";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly chartId: string;
  readonly chartName: string;
  readonly seriesNames: readonly string[];
  readonly seriesTypes: readonly string[];
  readonly colors: readonly ResolvedColor[];
  readonly previous: readonly PreviousColor[];
  readonly lines: readonly string[];
  readonly hasFill: boolean;
  readonly undoAvailable: boolean;
  readonly undoNote?: string;
  readonly createdAt: string;
}

export async function prepareFormatChartPlan(args: unknown): Promise<FormatChartPlan> {
  preflightToolArgs("format_chart", args);
  const a = args as { sheet?: string; chart: string; colors: ColorRequest[] };
  if (!a.colors?.length) throw new ToolError("Не задано ни одного цвета: colors пуст.");
  const target = await captureTarget(a.sheet);
  const prepared = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    sheet.load("name");
    try { sheet.protection?.load("protected"); } catch { /* среда без сведений о защите */ }
    const chart = sheet.charts.getItemOrNullObject(a.chart);
    chart.load(["isNullObject", "id", "name"]);
    const all = sheet.charts;
    all.load("items/name");
    await ctx.sync();
    if (chart.isNullObject) {
      throw new ToolError(`Диаграммы «${a.chart}» на листе ${sheet.name} нет. Есть: ${all.items.map((item) => `«${item.name}»`).join(", ") || "ни одной"}.`);
    }
    if (sheet.protection?.protected) throw new ToolError(`Лист ${sheet.name} защищён: диаграмму на нём не перекрасить. Операция не выполнялась.`);
    const series = chart.series;
    series.load("items/name,items/chartType");
    await ctx.sync();
    if (!series.items.length) throw new ToolError(`У диаграммы «${chart.name}» нет рядов.`);
    const counts = series.items.map((item) => { const points = item.points; points.load("count"); return points; });
    const categories = series.items[0].getDimensionValues("Categories" as any);
    await ctx.sync();
    const seriesNames = series.items.map((item) => String(item.name ?? ""));
    const seriesTypes = series.items.map((item) => String(item.chartType ?? ""));
    const colors = resolveColors(a.colors, {
      seriesNames,
      seriesTypes,
      categories: (categories.value ?? []).map((value) => String(value ?? "")),
      pointCounts: counts.map((item) => Number(item.count))
    });

    // Прежние цвета — там, где Excel их отдаёт: линия и маркеры.
    const reads = colors.filter((item) => paintStyle(seriesTypes[item.seriesIndex]) === "line").map((item) => {
      const s = series.items[item.seriesIndex];
      if (item.pointIndex === undefined) {
        s.format.line.load("color");
        s.load(["markerBackgroundColor", "markerForegroundColor"]);
        return { item, read: () => ({ seriesIndex: item.seriesIndex, line: s.format.line.color ?? null, markerBackground: s.markerBackgroundColor ?? null, markerForeground: s.markerForegroundColor ?? null }) };
      }
      const point = s.points.getItemAt(item.pointIndex);
      point.load(["markerBackgroundColor", "markerForegroundColor"]);
      return { item, read: () => ({ seriesIndex: item.seriesIndex, pointIndex: item.pointIndex, markerBackground: point.markerBackgroundColor ?? null, markerForeground: point.markerForegroundColor ?? null }) };
    });
    await ctx.sync();
    const previous = reads.map(({ read }) => read());
    const hasFill = colors.some((item) => paintStyle(seriesTypes[item.seriesIndex]) === "fill");
    const monitor = isCustomUndoAvailable();
    const undoAvailable = monitor && !hasFill;
    return {
      kind: "format_chart" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      target: { ...target, sheetName: sheet.name },
      chartId: String(chart.id),
      chartName: chart.name,
      seriesNames,
      seriesTypes,
      colors,
      previous,
      lines: colors.map((item) => `${describeColor(item)}${paintStyle(seriesTypes[item.seriesIndex]) === "fill" ? " (заливка)" : " (линия и маркеры)"}`),
      hasFill,
      undoAvailable,
      ...(undoAvailable ? {} : {
        undoNote: hasFill
          ? "Отмены не будет: прежнюю заливку эта версия Excel не сообщает, вернуть её нечем. Цвет можно поменять новой просьбой."
          : "Отмена недоступна: монитор изменений Excel не активен."
      }),
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

async function restoreColors(sheetId: string, chartId: string, previous: readonly PreviousColor[]) {
  await Excel.run(async (ctx) => {
    const chart = ctx.workbook.worksheets.getItem(sheetId).charts.getItemOrNullObject(chartId);
    chart.load("isNullObject");
    await ctx.sync();
    if (chart.isNullObject) throw new Error("Диаграммы уже нет: её удалили после операции агента. Отменять нечего.");
    for (const item of previous) {
      const series = chart.series.getItemAt(item.seriesIndex);
      if (item.pointIndex === undefined) {
        if (item.line) series.format.line.color = item.line;
        if (item.markerBackground) series.markerBackgroundColor = item.markerBackground;
        if (item.markerForeground) series.markerForegroundColor = item.markerForeground;
      } else {
        const point = series.points.getItemAt(item.pointIndex);
        if (item.markerBackground) point.markerBackgroundColor = item.markerBackground;
        if (item.markerForeground) point.markerForegroundColor = item.markerForeground;
      }
    }
    await ctx.sync();
  });
}

export async function executeFormatChartPlan(plan: FormatChartPlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(plan.target.sheetId);
    sheet.load("name");
    const chart = sheet.charts.getItemOrNullObject(plan.chartId);
    chart.load(["isNullObject", "name"]);
    await ctx.sync();
    if (chart.isNullObject) {
      throw new ToolExecutionError(`Диаграммы «${plan.chartName}» больше нет на листе ${sheet.name}. Ничего не менялось.`, "failed_before_write");
    }
    chart.series.load("items/name,items/chartType");
    await ctx.sync();
    const names = chart.series.items.map((item) => String(item.name ?? ""));
    const types = chart.series.items.map((item) => String(item.chartType ?? ""));
    if (JSON.stringify(names) !== JSON.stringify(plan.seriesNames) || JSON.stringify(types) !== JSON.stringify(plan.seriesTypes)) {
      throw new ToolExecutionError(
        `Ряды диаграммы «${plan.chartName}» изменились после предпросмотра. Цвета не менялись — сделайте новый предпросмотр.`,
        "failed_before_write"
      );
    }

    const result = await applyChartColors(ctx, chart, plan.colors, plan.seriesTypes);
    let undoRecorded = false;
    if (plan.undoAvailable && result.applied.length) {
      const sheetId = plan.target.sheetId;
      undoRecorded = push(action(`цвета диаграммы ${plan.chartName} на листе ${sheet.name}`, () => restoreColors(sheetId, plan.chartId, plan.previous)));
    }
    if (result.problems.length) {
      throw new ToolExecutionError(
        `Диаграмма «${plan.chartName}»: часть цветов Excel не принял или записал иначе: ${result.problems.join("; ")}.` +
          (undoRecorded ? " Прежние цвета линий и маркеров можно вернуть кнопкой «Отменить»." : " Проверьте диаграмму на листе."),
        result.applied.length ? "applied" : "unknown"
      );
    }
    return {
      ok: true,
      executionState: "verified",
      sheet: sheet.name,
      chart: plan.chartName,
      colors: result.applied.map(({ target, color, verified }) => ({ target, color, verified })),
      ...(result.unverifiable.length ? { colorNote: FILL_UNVERIFIABLE_NOTE } : {}),
      note: "Цвета линий и маркеров сверены с тем, что сообщил Excel. Как диаграмма выглядит целиком, панель не видит.",
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: plan.undoNote ?? "Автоматическая отмена этой операции недоступна." })
    };
  });
}
