/**
 * Диаграмма на проверяемом пути.
 *
 * Прежний `create_chart` клал диаграмму куда придётся и ничего не проверял:
 * какие ряды Excel построил, откуда взял подписи, не закрыла ли диаграмма
 * данные. Теперь подготовка по данным называет ряды, точки и подписи, которые
 * должны получиться, и выбирает место правее данных. Исполнение строит
 * диаграмму и сверяет её ряды с ожиданием: расхождение означает, что Excel
 * понял область иначе, и об этом говорится, а не молчится.
 */

import {
  assertPlanWorkbook,
  checkAddress,
  deepFreeze,
  preflightToolArgs,
  rangeOf,
  ToolError,
  ToolExecutionError
} from "./excelTools";
import {
  CHART_KINDS,
  CHARTS_WITH_TRENDLINES,
  CHARTS_WITHOUT_AXES,
  chartsOverlap,
  expectChart,
  freeChartTop,
  placementCell,
  seriesMismatches,
  type ChartBox,
  type ChartExpectation,
  type ChartKind,
  type DataLabelPosition,
  type LegendPosition,
  type SeriesBy,
  type TrendlineType
} from "./chartModel";
import { parseA1Rect, intersects } from "./a1";
import { action, getStructuralRevision, isCustomUndoAvailable, push } from "./undo";
import { captureTarget, officeCapabilities, type WorkbookTarget } from "./workbookContext";

/** Диаграмма по области больше этой не читается и строится долго. */
export const MAX_CHART_CELLS = 5_000;

export interface AxisRequest {
  readonly title?: string;
  readonly minimum?: number;
  readonly maximum?: number;
  readonly numberFormat?: string;
}

export interface DataLabelsRequest {
  readonly show: boolean;
  readonly position?: DataLabelPosition;
  readonly numberFormat?: string;
}

export interface LegendRequest {
  readonly position: LegendPosition;
}

/** Ряды уже сведены к индексам: имя из просьбы могло значить только то,
 * что предсказано по данным, а не то, как Excel сам назовёт ряд без шапки. */
export interface ResolvedTrendline {
  readonly type: TrendlineType;
  readonly movingAveragePeriod?: number;
  readonly indices: readonly number[];
  readonly seriesLabel: string;
}

export interface CreateChartPlan {
  readonly kind: "create_chart";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly address: string;
  readonly resolvedAddress: string;
  readonly chartType: ChartKind;
  readonly title?: string;
  readonly expectation: ChartExpectation;
  /** Левый верхний угол диаграммы. */
  readonly anchorCell: string;
  readonly anchorWarning?: string;
  /** Сколько диаграмм уже стоит на листе: новая встанет под ними. */
  readonly chartsOnSheet: number;
  /** Слепок данных: ручная правка до подтверждения меняет ожидание. */
  readonly signature: string;
  readonly axes?: { value?: AxisRequest; category?: { title: string } };
  readonly dataLabels?: DataLabelsRequest;
  readonly legend?: LegendRequest;
  readonly trendlines?: readonly ResolvedTrendline[];
  readonly undoAvailable: boolean;
  readonly undoNote?: string;
  readonly createdAt: string;
}

function withoutSheet(address: string): string {
  return address.slice(address.lastIndexOf("!") + 1);
}

/** Ряд без указания в просьбе значит «на все ряды». Указанный — ищется по
 * тому имени, которое ожидание уже посчитало (шапка или синтетическое
 * «столбец B»): как Excel в итоге назовёт ряд сам, роли не играет — линия
 * тренда ставится по позиции ряда, а не по его имени в Excel. */
function resolveTrendlines(
  requests: ReadonlyArray<{ series?: string; type: TrendlineType; movingAveragePeriod?: number }>,
  seriesNames: readonly string[]
): ResolvedTrendline[] {
  return requests.map((request) => {
    if (request.series === undefined) {
      return { type: request.type, movingAveragePeriod: request.movingAveragePeriod, indices: seriesNames.map((_, index) => index), seriesLabel: "все ряды" };
    }
    const index = seriesNames.indexOf(request.series);
    if (index === -1) {
      throw new ToolError(
        `Ряд «${request.series}» не найден для линии тренда: в диаграмме будут ряды ${seriesNames.map((name) => `«${name}»`).join(", ") || "—"}.`
      );
    }
    return { type: request.type, movingAveragePeriod: request.movingAveragePeriod, indices: [index], seriesLabel: request.series };
  });
}

export async function prepareCreateChartPlan(args: unknown): Promise<CreateChartPlan> {
  preflightToolArgs("create_chart", args);
  const a = args as {
    sheet?: string; address: string; chartType: string; title?: string; seriesBy?: SeriesBy; anchorCell?: string;
    axes?: { value?: { title?: string; minimum?: number; maximum?: number; numberFormat?: string }; category?: { title?: string } };
    dataLabels?: { show: boolean; position?: DataLabelPosition; numberFormat?: string };
    legend?: { position: LegendPosition };
    trendlines?: Array<{ series?: string; type: TrendlineType; movingAveragePeriod?: number }>;
  };
  if (!CHART_KINDS.includes(a.chartType as ChartKind)) throw new ToolError(`Неподдерживаемый тип диаграммы ${a.chartType}.`);
  const chartType = a.chartType as ChartKind;
  // Ограничения самого Excel: у круговой и кольцевой нет осей, линия тренда
  // не строится на них и на составных диаграммах. Проверяется до Excel —
  // это не то, что Excel «понял иначе», а то, чего он не умеет вовсе.
  if (a.axes && CHARTS_WITHOUT_AXES.has(chartType)) {
    throw new ToolError(`У диаграммы ${chartType} нет осей значений и категорий: круговая и кольцевая построены без них. Уберите axes или выберите другой тип.`);
  }
  if (a.trendlines?.length && !CHARTS_WITH_TRENDLINES.has(chartType)) {
    throw new ToolError(
      `Линия тренда не строится на ${chartType}: Excel не поддерживает её для круговых, кольцевых и составных диаграмм. ` +
        "Уберите trendlines или выберите ColumnClustered, BarClustered, Line, Area или XYScatter."
    );
  }
  for (const request of a.trendlines ?? []) {
    if (request.movingAveragePeriod !== undefined && request.type !== "MovingAverage") {
      throw new ToolError(`movingAveragePeriod задан для типа ${request.type}: он действует только для MovingAverage.`);
    }
  }
  const seriesBy: SeriesBy = a.seriesBy === "rows" ? "rows" : "columns";
  const address = checkAddress(a.address);
  if (a.anchorCell !== undefined) {
    const anchor = parseA1Rect(a.anchorCell);
    if (!anchor || anchor.kind !== "cells" || anchor.rowStart !== anchor.rowEnd || anchor.columnStart !== anchor.columnEnd) {
      throw new ToolError(`anchorCell должен быть одной ячейкой, например H2; получено «${a.anchorCell}».`);
    }
  }
  const target = await captureTarget(a.sheet);

  const prepared = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    const range = await rangeOf(ctx, sheet, address);
    range.load(["address", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
    sheet.load(["id", "name"]);
    try { sheet.protection?.load("protected"); } catch { /* среда без сведений о защите */ }
    const used = officeCapabilities().usedRangeOrNull
      ? sheet.getUsedRangeOrNullObject(true)
      : sheet.getUsedRange(true);
    used.load(["isNullObject", "address", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
    await ctx.sync();
    if (sheet.protection?.protected) {
      throw new ToolError(`Лист ${sheet.name} защищён: диаграмму на нём создать нельзя. Операция не выполнялась.`);
    }
    const cells = range.rowCount * range.columnCount;
    if (cells > MAX_CHART_CELLS) {
      throw new ToolError(
        `Диаграмма строится по области до ${MAX_CHART_CELLS} ячеек; ${range.address} содержит ${cells}. ` +
        "Для больших данных сначала сведите их, например сводной таблицей или итогами."
      );
    }
    range.load(["values", "formulas"]);
    await ctx.sync();
    const values = range.values as unknown[][];

    const expectation = expectChart(values, chartType, seriesBy, { rowIndex: range.rowIndex, columnIndex: range.columnIndex });
    if (expectation.warnings[0]?.startsWith("В области нет чисел")) {
      throw new ToolError(`${range.address}: в области нет чисел, строить не из чего. Операция не выполнялась.`);
    }
    const trendlines = a.trendlines?.length ? resolveTrendlines(a.trendlines, expectation.seriesNames) : undefined;

    const empty = Boolean((used as any).isNullObject);
    const anchorCell = a.anchorCell?.trim().toUpperCase()
      ?? placementCell(empty ? null : { rowIndex: used.rowIndex, columnIndex: used.columnIndex, columnCount: used.columnCount }, range.rowIndex);
    let anchorWarning: string | undefined;
    if (a.anchorCell && !empty) {
      const anchor = parseA1Rect(anchorCell);
      const data = parseA1Rect(withoutSheet(used.address));
      if (anchor && data && intersects(anchor, data)) {
        anchorWarning = `Ячейка ${anchorCell} внутри занятой области ${withoutSheet(used.address)}: диаграмма ляжет поверх данных и закроет их.`;
      }
    }

    const charts = sheet.charts;
    charts.load("items/name");
    await ctx.sync();

    const undo = isCustomUndoAvailable();
    return {
      kind: "create_chart" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      target: { ...target, sheetName: sheet.name },
      address,
      resolvedAddress: withoutSheet(range.address),
      chartType,
      ...(a.title?.trim() ? { title: a.title.trim() } : {}),
      expectation,
      anchorCell,
      chartsOnSheet: charts.items.length,
      ...(anchorWarning ? { anchorWarning } : {}),
      signature: JSON.stringify(range.formulas),
      ...(a.axes ? { axes: a.axes as { value?: AxisRequest; category?: { title: string } } } : {}),
      ...(a.dataLabels ? { dataLabels: a.dataLabels } : {}),
      ...(a.legend ? { legend: a.legend } : {}),
      ...(trendlines ? { trendlines } : {}),
      undoAvailable: undo,
      ...(undo ? {} : { undoNote: "Отмена недоступна: монитор изменений Excel не активен." }),
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executeCreateChartPlan(plan: CreateChartPlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(plan.target.sheetId);
    sheet.load(["id", "name"]);
    const range = sheet.getRange(plan.resolvedAddress);
    range.load("formulas");
    await ctx.sync();
    if (JSON.stringify(range.formulas) !== plan.signature) {
      throw new ToolExecutionError(
        `Данные ${sheet.name}!${plan.resolvedAddress} изменились после предпросмотра. Диаграмма не строилась — сделайте новый предпросмотр.`,
        "failed_before_write"
      );
    }

    let chart: Excel.Chart;
    let existing: Excel.ChartCollection;
    try {
      chart = sheet.charts.add(
        plan.chartType as any,
        range,
        (plan.expectation.seriesBy === "rows" ? "Rows" : "Columns") as any
      );
      chart.setPosition(plan.anchorCell);
      if (plan.title) chart.title.text = plan.title;
      chart.load(["id", "name", "chartType", "top", "left", "height", "width"]);
      chart.series.load("items/name");
      existing = sheet.charts;
      existing.load("items/name,items/top,items/left,items/height,items/width");
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(
        `Excel отказал в построении диаграммы по ${sheet.name}!${plan.resolvedAddress}: ${error?.message ?? error}. ` +
        "Неизвестно, успела ли она появиться — посмотрите на лист.",
        "unknown"
      );
    }

    // Проверка в Excel 20 сентября 2026 года: вторая диаграмма встала в ту же
    // ячейку, что и первая, и легла поверх неё. Место правее данных знает
    // только про ячейки, а диаграммы лежат над ними — поэтому новая
    // опускается под те, с которыми пересеклась, — до места, свободного от
    // всех (S6: однократный сдвиг клал третью диаграмму на вторую).
    let movedNote: string | undefined;
    let placementProblem: string | undefined;
    const others: ChartBox[] = existing.items
      .filter((item) => item.name !== chart.name)
      .map((item) => ({ name: item.name, left: item.left, top: item.top, width: item.width, height: item.height }));
    const place = freeChartTop({ left: chart.left, top: chart.top, width: chart.width, height: chart.height }, others);
    if (!place) {
      placementProblem = `свободного места под ${others.length} диаграммами листа не нашлось — новая легла поверх других`;
    } else if (place.passed.length) {
      chart.top = place.top;
      movedNote = `На листе уже ${others.length} диаграмм: новая опущена под ${place.passed.join(", ")}, чтобы не закрыть их.`;
    }
    // Положение сверяется по тому, что Excel отдаёт после записи, а не по
    // расчёту: успешный sync не доказывает, что диаграмма сдвинулась.
    chart.load(["top", "left", "height", "width"]);
    await ctx.sync();
    const position = { top: chart.top, left: chart.left, width: chart.width, height: chart.height };
    const covered = others.filter((item) => chartsOverlap(position, item)).map((item) => item.name);
    if (!placementProblem && covered.length) placementProblem = `новая легла поверх ${covered.join(", ")}`;

    const chartId = String(chart.id);
    // Отмена записывается сразу: диаграмма уже есть, и даже при расхождении
    // пользователю нужен способ её убрать.
    let undoRecorded = false;
    if (plan.undoAvailable) {
      const sheetId = plan.target.sheetId;
      undoRecorded = push(action(`диаграмма ${chart.name} на листе ${sheet.name}`, async () => {
        const revision = getStructuralRevision();
        await Excel.run(async (undoCtx) => {
          const existing = undoCtx.workbook.worksheets.getItem(sheetId).charts.getItemOrNullObject(chartId);
          existing.load("isNullObject");
          await undoCtx.sync();
          if (existing.isNullObject) throw new Error("Диаграммы уже нет: её удалили после операции агента. Отменять нечего.");
          if (getStructuralRevision() !== revision) throw new Error("Структура книги изменилась во время отмены. Отмена остановлена.");
          existing.delete();
          await undoCtx.sync();
        });
      }));
    }

    // Оформление (8.1): каждая группа свойств — свой sync, чтобы отказ Excel
    // в одной (например, в положении подписи, которое годится не для всякого
    // типа диаграммы) не скрывал, что остальные применились. Сверяется то,
    // что Excel подтвердил обратным чтением, а не то, что было запрошено.
    const formattingProblems: string[] = [];
    let appliedAxes: { value?: Record<string, unknown>; category?: Record<string, unknown> } | undefined;
    let appliedDataLabels: Record<string, unknown> | undefined;
    let appliedLegend: { position: string } | undefined;
    let appliedTrendlines: Array<{ series: string; type: string; movingAveragePeriod?: number }> | undefined;

    if (plan.axes?.value) {
      const v = plan.axes.value;
      try {
        const axis = chart.axes.getItem("Value");
        if (v.title !== undefined) { axis.title.text = v.title; axis.title.visible = true; }
        if (v.minimum !== undefined) axis.minimum = v.minimum;
        if (v.maximum !== undefined) axis.maximum = v.maximum;
        if (v.numberFormat !== undefined) axis.numberFormat = v.numberFormat;
        axis.load(["minimum", "maximum", "numberFormat"]);
        axis.title.load(["text", "visible"]);
        await ctx.sync();
        appliedAxes = { ...appliedAxes, value: {
          ...(v.title !== undefined ? { title: axis.title.visible ? String(axis.title.text ?? "") : null } : {}),
          ...(v.minimum !== undefined ? { minimum: Number(axis.minimum) } : {}),
          ...(v.maximum !== undefined ? { maximum: Number(axis.maximum) } : {}),
          ...(v.numberFormat !== undefined ? { numberFormat: String(axis.numberFormat ?? "") } : {})
        } };
        if (v.title !== undefined && (!axis.title.visible || String(axis.title.text ?? "") !== v.title)) {
          formattingProblems.push(`заголовок оси значений — «${axis.title.visible ? axis.title.text : ""}» вместо «${v.title}»`);
        }
        if (v.minimum !== undefined && Number(axis.minimum) !== v.minimum) {
          formattingProblems.push(`минимум оси значений — ${axis.minimum} вместо ${v.minimum}`);
        }
        if (v.maximum !== undefined && Number(axis.maximum) !== v.maximum) {
          formattingProblems.push(`максимум оси значений — ${axis.maximum} вместо ${v.maximum}`);
        }
        if (v.numberFormat !== undefined && String(axis.numberFormat ?? "") !== v.numberFormat) {
          formattingProblems.push(`числовой формат оси значений — «${axis.numberFormat}» вместо «${v.numberFormat}»`);
        }
      } catch (error: any) {
        formattingProblems.push(`ось значений: Excel отказал (${error?.message ?? error})`);
      }
    }
    if (plan.axes?.category?.title !== undefined) {
      const requestedTitle = plan.axes.category.title;
      try {
        const axis = chart.axes.getItem("Category");
        axis.title.text = requestedTitle;
        axis.title.visible = true;
        axis.title.load(["text", "visible"]);
        await ctx.sync();
        appliedAxes = { ...appliedAxes, category: { title: axis.title.visible ? String(axis.title.text ?? "") : null } };
        if (!axis.title.visible || String(axis.title.text ?? "") !== requestedTitle) {
          formattingProblems.push(`заголовок оси категорий — «${axis.title.visible ? axis.title.text : ""}» вместо «${requestedTitle}»`);
        }
      } catch (error: any) {
        formattingProblems.push(`ось категорий: Excel отказал (${error?.message ?? error})`);
      }
    }
    if (plan.legend) {
      const requested = plan.legend.position;
      try {
        if (requested === "None") {
          chart.legend.visible = false;
          chart.legend.load("visible");
          await ctx.sync();
          appliedLegend = { position: "None" };
          if (chart.legend.visible !== false) formattingProblems.push("легенда — не скрылась");
        } else {
          chart.legend.visible = true;
          chart.legend.position = requested;
          chart.legend.load(["visible", "position"]);
          await ctx.sync();
          const actual = chart.legend.visible ? String(chart.legend.position) : "None";
          appliedLegend = { position: actual };
          if (actual !== requested) {
            formattingProblems.push(`легенда — ${actual === "None" ? "скрыта" : `положение ${actual}`} вместо ${requested}`);
          }
        }
      } catch (error: any) {
        formattingProblems.push(`легенда: Excel отказал (${error?.message ?? error})`);
      }
    }
    if (plan.dataLabels) {
      const { show, position, numberFormat } = plan.dataLabels;
      try {
        chart.dataLabels.showValue = show;
        chart.dataLabels.load("showValue");
        await ctx.sync();
        appliedDataLabels = { show: Boolean(chart.dataLabels.showValue) };
        if (Boolean(chart.dataLabels.showValue) !== show) {
          formattingProblems.push(`подписи данных — ${chart.dataLabels.showValue ? "показаны" : "скрыты"} вместо ${show ? "показанных" : "скрытых"}`);
        }
      } catch (error: any) {
        formattingProblems.push(`подписи данных: Excel отказал (${error?.message ?? error})`);
      }
      // Положение подписи годится не для всякого типа диаграммы — отдельный
      // sync, чтобы отказ здесь не выглядел так, будто подписи вовсе не включились.
      if (show && (position || numberFormat)) {
        try {
          if (position) chart.dataLabels.position = position;
          if (numberFormat) chart.dataLabels.numberFormat = numberFormat;
          chart.dataLabels.load(["position", "numberFormat"]);
          await ctx.sync();
          appliedDataLabels = {
            ...appliedDataLabels,
            ...(position ? { position: String(chart.dataLabels.position) } : {}),
            ...(numberFormat ? { numberFormat: String(chart.dataLabels.numberFormat ?? "") } : {})
          };
          if (position && String(chart.dataLabels.position) !== position) {
            formattingProblems.push(`подписи данных — положение «${chart.dataLabels.position}» вместо «${position}»`);
          }
          if (numberFormat && String(chart.dataLabels.numberFormat ?? "") !== numberFormat) {
            formattingProblems.push(`подписи данных — числовой формат «${chart.dataLabels.numberFormat}» вместо «${numberFormat}»`);
          }
        } catch (error: any) {
          formattingProblems.push(`подписи данных: положение или формат Excel не принял (${error?.message ?? error})`);
        }
      }
    }
    if (plan.trendlines?.length) {
      try {
        const added: Array<{ label: string; type: TrendlineType; line: Excel.ChartTrendline }> = [];
        for (const request of plan.trendlines) {
          for (const index of request.indices) {
            const line = chart.series.items[index].trendlines.add(request.type);
            if (request.type === "MovingAverage" && request.movingAveragePeriod) {
              line.movingAveragePeriod = request.movingAveragePeriod;
            }
            line.load(["type", "movingAveragePeriod"]);
            added.push({ label: plan.expectation.seriesNames[index] ?? `ряд ${index + 1}`, type: request.type, line });
          }
        }
        await ctx.sync();
        appliedTrendlines = added.map(({ label, type, line }) => ({
          series: label,
          type: String(line.type),
          ...(type === "MovingAverage" ? { movingAveragePeriod: Number(line.movingAveragePeriod) } : {})
        }));
        for (const { label, type, line } of added) {
          if (String(line.type) !== type) formattingProblems.push(`линия тренда на «${label}» — тип ${line.type} вместо ${type}`);
        }
      } catch (error: any) {
        formattingProblems.push(`линия тренда: Excel отказал (${error?.message ?? error})`);
      }
    }

    const series = chart.series.items;
    const points = series.map((item) => {
      const collection = item.points;
      collection.load("count");
      return collection;
    });
    chart.title.load("text");
    await ctx.sync();
    const title = String(chart.title.text ?? "");

    const actual = { names: series.map((item) => String(item.name ?? "")), pointCounts: points.map((item) => Number(item.count)) };
    const problems = seriesMismatches(plan.expectation, actual);
    if (placementProblem) problems.push(placementProblem);
    if (String(chart.chartType) !== plan.chartType) problems.unshift(`тип ${chart.chartType} вместо ${plan.chartType}`);
    if (plan.title && title !== plan.title) problems.push(`заголовок «${title}» вместо «${plan.title}»`);
    if (problems.length) {
      // Что именно построил Excel, важнее самих чисел: по именам рядов и
      // заголовку сразу видно, какой столбец он принял за данные.
      const built = `Excel построил ряды ${actual.names.map((name) => `«${name}»`).join(", ") || "без имён"}` +
        (plan.title ? "" : `, заголовок «${title}»`) + ".";
      throw new ToolExecutionError(
        `Диаграмма ${chart.name} построена, но Excel понял область иначе, чем ожидалось: ${problems.join("; ")}. ${built} ` +
        (undoRecorded ? "Её можно убрать кнопкой «Отменить» и построить заново, например с другим направлением рядов (seriesBy)." : "Проверьте её на листе."),
        "applied"
      );
    }
    // Ряды поняты верно; отдельно — оформление, которое Excel мог принять не всё.
    if (formattingProblems.length) {
      throw new ToolExecutionError(
        `Диаграмма ${chart.name} построена, ряды и точки совпадают с ожиданием, но часть оформления Excel не принял: ${formattingProblems.join("; ")}. ` +
        (undoRecorded ? "Её можно убрать кнопкой «Отменить» и построить заново с другими свойствами." : "Проверьте её на листе."),
        "applied"
      );
    }

    return {
      ok: true,
      executionState: "verified",
      sheet: sheet.name,
      chart: chart.name,
      chartType: String(chart.chartType),
      source: plan.resolvedAddress,
      anchorCell: plan.anchorCell,
      ...(movedNote ? { placementNote: movedNote } : {}),
      // Фактическое положение, прочитанное после записи, в пунктах.
      position,
      series: actual.names,
      pointsPerSeries: plan.expectation.pointCount,
      ...(plan.expectation.categories.length ? { categories: plan.expectation.categories } : {}),
      ...(title ? { title } : {}),
      ...(plan.expectation.warnings.length ? { warnings: plan.expectation.warnings } : {}),
      ...(appliedAxes ? { axes: appliedAxes } : {}),
      ...(appliedDataLabels ? { dataLabels: appliedDataLabels } : {}),
      ...(appliedLegend ? { legend: appliedLegend } : {}),
      ...(appliedTrendlines ? { trendlines: appliedTrendlines } : {}),
      note: "Ряды, точки и запрошенное оформление сверены с тем, что сообщил Excel о построенной диаграмме. Как она выглядит целиком, панель не видит.",
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: plan.undoNote ?? "Автоматическая отмена этой операции недоступна." })
    };
  });
}
