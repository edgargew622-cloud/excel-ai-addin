/**
 * Измерение длительности запросов и размера результатов.
 *
 * Метрика по построению не может содержать данные книги: в ней только имя
 * операции, время и размер в байтах. Это свойство зафиксировано тестом, потому
 * что технические журналы не должны содержать содержимое ячеек и ключи.
 */

export type MetricKind = "model" | "tool";

export interface OpMetric {
  kind: MetricKind;
  /** Имя провайдера с моделью или имя инструмента. */
  name: string;
  ms: number;
  /** Размер результата в байтах. Для модели — размер полученного текста. */
  bytes: number;
  ok: boolean;
}

export interface MetricsSummary {
  count: number;
  totalMs: number;
  totalBytes: number;
  modelCount: number;
  toolCount: number;
  slowest: OpMetric | null;
  largest: OpMetric | null;
}

export interface MetricsStore {
  record(metric: OpMetric): void;
  all(): OpMetric[];
  summary(): MetricsSummary;
  reset(): void;
}

/** Держим ограниченное окно: панель живёт долго, память не резиновая. */
export const MAX_RETAINED_METRICS = 200;

export function createMetrics(limit: number = MAX_RETAINED_METRICS): MetricsStore {
  let metrics: OpMetric[] = [];

  return {
    record(metric) {
      metrics.push(metric);
      if (metrics.length > limit) metrics = metrics.slice(-limit);
    },
    all() {
      return [...metrics];
    },
    summary() {
      return summarize(metrics);
    },
    reset() {
      metrics = [];
    }
  };
}

export function summarize(metrics: OpMetric[]): MetricsSummary {
  let totalMs = 0;
  let totalBytes = 0;
  let modelCount = 0;
  let toolCount = 0;
  let slowest: OpMetric | null = null;
  let largest: OpMetric | null = null;

  for (const m of metrics) {
    totalMs += m.ms;
    totalBytes += m.bytes;
    if (m.kind === "model") modelCount += 1;
    else toolCount += 1;
    if (!slowest || m.ms > slowest.ms) slowest = m;
    if (!largest || m.bytes > largest.bytes) largest = m;
  }

  return { count: metrics.length, totalMs, totalBytes, modelCount, toolCount, slowest, largest };
}

export function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} Б`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} КБ`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} МБ`;
}

export function formatMs(ms: number): string {
  if (ms < 1000) return `${Math.round(ms)} мс`;
  return `${(ms / 1000).toFixed(1)} с`;
}

/** Одна строка для технического журнала: только имена, время и размеры. */
export function formatMetricLog(metric: OpMetric): string {
  return `kind=${metric.kind} name=${metric.name} ms=${Math.round(metric.ms)} bytes=${metric.bytes} ok=${metric.ok}`;
}

/** Размер полезной нагрузки в байтах UTF-8, а не в символах. */
export function byteLength(payload: string): number {
  if (typeof TextEncoder !== "undefined") return new TextEncoder().encode(payload).length;
  return Buffer.byteLength(payload, "utf8");
}
