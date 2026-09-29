/**
 * Цвета диаграмм (этап 10, 10.2): цвет ряда и отдельной точки.
 *
 * Замер 30.09.2026 на Office 2021 (ExcelApi 1.14):
 * - линия ряда (`format.line.color`), маркеры ряда и точки
 *   (`markerBackgroundColor`/`markerForegroundColor`), граница точки —
 *   записываются и читаются обратно;
 * - заливка столбцов, полос, долей и областей (`format.fill.setSolidColor`) —
 *   записывается, а прочитать её нельзя: `getSolidColor` появился в ExcelApi
 *   1.16. Поэтому заливку на такой версии не сверить и прежний цвет не узнать.
 *
 * Как красить, зависит от типа ряда: у столбцов, полос, долей и областей —
 * заливка, у графиков и точечной — линия и маркеры.
 */

import { ToolError } from "./excelTools";

export interface ColorRequest {
  /** Ряд по имени; можно не указывать, если ряд один. */
  readonly series?: string;
  /** Точка по номеру, с 1. */
  readonly point?: number;
  /** Точка по подписи категории, например «Мар». */
  readonly category?: string;
  readonly color: string;
}

export interface ResolvedColor {
  readonly seriesIndex: number;
  readonly seriesName: string;
  /** С 0; нет — цвет всего ряда. */
  readonly pointIndex?: number;
  readonly pointLabel?: string;
  readonly color: string;
}

export type PaintStyle = "fill" | "line";

const LINE_TYPES = new Set(["Line", "LineMarkers", "LineStacked", "LineMarkersStacked", "LineStacked100", "LineMarkersStacked100", "XYScatter", "XYScatterLines", "XYScatterLinesNoMarkers", "XYScatterSmooth", "XYScatterSmoothNoMarkers", "Radar", "RadarMarkers"]);
const POINTS_ONLY = new Set(["Pie", "PieExploded", "Pie3D", "Doughnut", "DoughnutExploded"]);

export function paintStyle(chartType: string): PaintStyle {
  return LINE_TYPES.has(chartType) ? "line" : "fill";
}

export function normalizeColor(value: unknown): string {
  const text = String(value ?? "").trim();
  if (!/^#?[0-9A-Fa-f]{6}$/.test(text)) throw new ToolError(`Цвет «${text}» не в формате HEX, ожидается #RRGGBB, например #C00000.`);
  return `#${text.replace(/^#/, "").toUpperCase()}`;
}

/**
 * Просьбы о цвете — к рядам и точкам по данным диаграммы. Ряд ищется по
 * имени, точка — по номеру или подписи категории; неоднозначное и
 * несуществующее отклоняется до Excel.
 */
export function resolveColors(
  requests: readonly ColorRequest[],
  chart: { seriesNames: readonly string[]; seriesTypes: readonly string[]; categories: readonly string[]; pointCounts: readonly number[] }
): ResolvedColor[] {
  if (requests.length > 60) throw new ToolError("За раз можно задать не больше 60 цветов.");
  const out: ResolvedColor[] = [];
  const seen = new Set<string>();
  for (const request of requests) {
    const color = normalizeColor(request.color);
    let seriesIndex: number;
    if (request.series === undefined) {
      if (chart.seriesNames.length !== 1) {
        throw new ToolError(`Укажите ряд (series): в диаграмме ряды ${chart.seriesNames.map((name) => `«${name}»`).join(", ")}.`);
      }
      seriesIndex = 0;
    } else {
      seriesIndex = chart.seriesNames.indexOf(request.series);
      if (seriesIndex === -1) {
        throw new ToolError(`Ряд «${request.series}» не найден: в диаграмме ряды ${chart.seriesNames.map((name) => `«${name}»`).join(", ") || "—"}.`);
      }
    }
    const seriesName = chart.seriesNames[seriesIndex];
    if (request.point !== undefined && request.category !== undefined) throw new ToolError("Укажите точку либо номером (point), либо подписью (category), не обоими.");
    let pointIndex: number | undefined;
    let pointLabel: string | undefined;
    if (request.point !== undefined) {
      const count = chart.pointCounts[seriesIndex] ?? 0;
      if (!Number.isInteger(request.point) || request.point < 1 || request.point > count) {
        throw new ToolError(`Точки ${request.point} нет в ряду «${seriesName}»: в нём ${count} точек.`);
      }
      pointIndex = request.point - 1;
      pointLabel = chart.categories[pointIndex] ?? `точка ${request.point}`;
    } else if (request.category !== undefined) {
      const matches = chart.categories.map((label, index) => (label === request.category ? index : -1)).filter((index) => index >= 0);
      if (!matches.length) {
        throw new ToolError(`Категории «${request.category}» нет: подписи ${chart.categories.slice(0, 20).map((label) => `«${label}»`).join(", ")}${chart.categories.length > 20 ? " …" : ""}.`);
      }
      if (matches.length > 1) throw new ToolError(`Подпись «${request.category}» встречается ${matches.length} раза — укажите точку номером (point).`);
      pointIndex = matches[0];
      pointLabel = request.category;
    }
    if (pointIndex === undefined && POINTS_ONLY.has(chart.seriesTypes[seriesIndex] ?? "")) {
      throw new ToolError("У круговой и кольцевой цвет задаётся долям: укажите category или point для каждой доли.");
    }
    const key = `${seriesIndex}:${pointIndex ?? "*"}`;
    if (seen.has(key)) throw new ToolError(`Цвет для ${pointLabel ? `«${pointLabel}» ряда «${seriesName}»` : `ряда «${seriesName}»`} указан дважды.`);
    seen.add(key);
    out.push({ seriesIndex, seriesName, ...(pointIndex !== undefined ? { pointIndex, pointLabel } : {}), color });
  }
  // Цвет ряда кладётся раньше цветов его точек: иначе он перекрасил бы их.
  return out.sort((a, b) => Number(a.pointIndex !== undefined) - Number(b.pointIndex !== undefined));
}

export function describeColor(item: ResolvedColor): string {
  return item.pointIndex === undefined ? `ряд «${item.seriesName}» — ${item.color}` : `«${item.pointLabel}» в ряду «${item.seriesName}» — ${item.color}`;
}

/** Прежний цвет — только то, что Excel даёт прочитать (линия и маркеры). */
export interface PreviousColor {
  readonly seriesIndex: number;
  readonly pointIndex?: number;
  readonly line?: string | null;
  readonly markerBackground?: string | null;
  readonly markerForeground?: string | null;
}

export interface ColorResult {
  readonly applied: Array<{ target: string; style: PaintStyle; color: string; verified: boolean; read?: string | null }>;
  readonly problems: string[];
  /** Заливку этой версии Excel не прочитать — сверки нет. */
  readonly unverifiable: string[];
}

function fillReadable(): boolean {
  try {
    return Office.context.requirements.isSetSupported("ExcelApi", "1.16");
  } catch {
    return false;
  }
}

/** Записать цвета в диаграмму и сверить то, что Excel даёт прочитать. */
export async function applyChartColors(
  ctx: Excel.RequestContext,
  chart: Excel.Chart,
  colors: readonly ResolvedColor[],
  seriesTypes: readonly string[]
): Promise<ColorResult> {
  const applied: Array<{ target: string; style: PaintStyle; color: string; verified: boolean; read?: string | null }> = [];
  const problems: string[] = [];
  const unverifiable: string[] = [];
  const canReadFill = fillReadable();
  for (const item of colors) {
    const style = paintStyle(seriesTypes[item.seriesIndex] ?? "");
    const target = item.pointIndex === undefined ? `ряд «${item.seriesName}»` : `«${item.pointLabel}» в ряду «${item.seriesName}»`;
    try {
      const series = chart.series.getItemAt(item.seriesIndex);
      if (style === "line") {
        if (item.pointIndex === undefined) {
          series.format.line.color = item.color;
          series.markerBackgroundColor = item.color;
          series.markerForegroundColor = item.color;
          series.format.line.load("color");
          series.load(["markerBackgroundColor", "markerForegroundColor"]);
          await ctx.sync();
          const read = String(series.format.line.color ?? "").toUpperCase();
          const ok = read === item.color && String(series.markerBackgroundColor ?? "").toUpperCase() === item.color;
          applied.push({ target, style, color: item.color, verified: ok, read });
          if (!ok) problems.push(`${target} — линия ${read || "не прочиталась"}, маркер ${series.markerBackgroundColor} вместо ${item.color}`);
        } else {
          const point = series.points.getItemAt(item.pointIndex);
          point.markerBackgroundColor = item.color;
          point.markerForegroundColor = item.color;
          point.load(["markerBackgroundColor", "markerForegroundColor"]);
          await ctx.sync();
          const read = String(point.markerBackgroundColor ?? "").toUpperCase();
          const ok = read === item.color;
          applied.push({ target, style, color: item.color, verified: ok, read });
          if (!ok) problems.push(`${target} — маркер ${read || "не прочитался"} вместо ${item.color}`);
        }
      } else {
        const fill = item.pointIndex === undefined ? series.format.fill : series.points.getItemAt(item.pointIndex).format.fill;
        fill.setSolidColor(item.color);
        await ctx.sync();
        if (canReadFill) {
          const read = (fill as any).getSolidColor();
          await ctx.sync();
          const value = String(read.value ?? "").toUpperCase();
          const ok = value === item.color;
          applied.push({ target, style, color: item.color, verified: ok, read: value });
          if (!ok) problems.push(`${target} — заливка ${value || "не прочиталась"} вместо ${item.color}`);
        } else {
          applied.push({ target, style, color: item.color, verified: false });
          unverifiable.push(target);
        }
      }
    } catch (error: any) {
      problems.push(`${target}: Excel не принял цвет (${error?.message ?? error})`);
    }
  }
  return { applied, problems, unverifiable };
}

export const FILL_UNVERIFIABLE_NOTE =
  "Заливку столбцов, полос, долей и областей эта версия Excel записывает, но не даёт прочитать обратно (нужен ExcelApi 1.16): она принята без ошибки, но не сверена — посмотрите на диаграмму.";
