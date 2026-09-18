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
  checkAddress,
  deepFreeze,
  preflightToolArgs,
  rangeOf,
  ToolError,
  ToolExecutionError
} from "./excelTools";
import {
  CHART_KINDS,
  expectChart,
  placementCell,
  seriesMismatches,
  type ChartExpectation,
  type ChartKind,
  type SeriesBy
} from "./chartModel";
import { parseA1Rect, intersects } from "./a1";
import { action, getStructuralRevision, isCustomUndoAvailable, push } from "./undo";
import { captureTarget, officeCapabilities, type WorkbookTarget } from "./workbookContext";

/** Диаграмма по области больше этой не читается и строится долго. */
export const MAX_CHART_CELLS = 5_000;

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
  /** Слепок данных: ручная правка до подтверждения меняет ожидание. */
  readonly signature: string;
  readonly undoAvailable: boolean;
  readonly undoNote?: string;
  readonly createdAt: string;
}

function withoutSheet(address: string): string {
  return address.slice(address.lastIndexOf("!") + 1);
}

export async function prepareCreateChartPlan(args: unknown): Promise<CreateChartPlan> {
  preflightToolArgs("create_chart", args);
  const a = args as { sheet?: string; address: string; chartType: string; title?: string; seriesBy?: SeriesBy; anchorCell?: string };
  if (!CHART_KINDS.includes(a.chartType as ChartKind)) throw new ToolError(`Неподдерживаемый тип диаграммы ${a.chartType}.`);
  const chartType = a.chartType as ChartKind;
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
      ...(anchorWarning ? { anchorWarning } : {}),
      signature: JSON.stringify(range.formulas),
      undoAvailable: undo,
      ...(undo ? {} : { undoNote: "Отмена недоступна: монитор изменений Excel не активен." }),
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executeCreateChartPlan(plan: CreateChartPlan) {
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
    try {
      chart = sheet.charts.add(
        plan.chartType as any,
        range,
        (plan.expectation.seriesBy === "rows" ? "Rows" : "Columns") as any
      );
      chart.setPosition(plan.anchorCell);
      if (plan.title) chart.title.text = plan.title;
      chart.load(["id", "name", "chartType"]);
      chart.series.load("items/name");
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(
        `Excel отказал в построении диаграммы по ${sheet.name}!${plan.resolvedAddress}: ${error?.message ?? error}. ` +
        "Неизвестно, успела ли она появиться — посмотрите на лист.",
        "unknown"
      );
    }

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

    const series = chart.series.items;
    const points = series.map((item) => {
      const collection = item.points;
      collection.load("count");
      return collection;
    });
    let title: string | null = null;
    if (plan.title) chart.title.load("text");
    await ctx.sync();
    if (plan.title) title = String(chart.title.text ?? "");

    const actual = { names: series.map((item) => String(item.name ?? "")), pointCounts: points.map((item) => Number(item.count)) };
    const problems = seriesMismatches(plan.expectation, actual);
    if (String(chart.chartType) !== plan.chartType) problems.unshift(`тип ${chart.chartType} вместо ${plan.chartType}`);
    if (plan.title && title !== plan.title) problems.push(`заголовок «${title}» вместо «${plan.title}»`);
    if (problems.length) {
      throw new ToolExecutionError(
        `Диаграмма ${chart.name} построена, но Excel понял область иначе, чем ожидалось: ${problems.join("; ")}. ` +
        (undoRecorded ? "Её можно убрать кнопкой «Отменить» и построить заново, например с другим направлением рядов (seriesBy)." : "Проверьте её на листе."),
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
      series: actual.names,
      pointsPerSeries: plan.expectation.pointCount,
      ...(plan.expectation.categories.length ? { categories: plan.expectation.categories } : {}),
      ...(title !== null ? { title } : {}),
      ...(plan.expectation.warnings.length ? { warnings: plan.expectation.warnings } : {}),
      note: "Ряды и точки сверены с тем, что сообщил Excel о построенной диаграмме. Как она выглядит, панель не видит.",
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: plan.undoNote ?? "Автоматическая отмена этой операции недоступна." })
    };
  });
}
