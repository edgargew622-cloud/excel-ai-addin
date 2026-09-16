/**
 * Измерение обращений к моделям.
 *
 * План требует знать длительность, размеры запросов, число обращений
 * и фактический usage — и при этом не писать содержимое книги в технический
 * журнал. Поэтому запись метрики по построению состоит только из чисел,
 * имени провайдера и модели: положить туда данные ячеек нечего. Это свойство
 * закреплено тестом, а не намерением.
 *
 * Usage приходит от провайдера, поэтому измерение живёт на сервере: панель
 * видит только текст ответа и посчитать токены может лишь на глаз.
 */

export interface Usage {
  promptTokens: number;
  completionTokens: number;
  totalTokens: number;
  /** Токены размышлений, если провайдер их отделяет. */
  reasoningTokens?: number;
  /** Prompt-токены, обслуженные кэшем: по ним видна экономия на повторах. */
  cachedTokens?: number;
}

export interface RequestMetric {
  provider: string;
  model: string;
  api: "chat" | "responses";
  requestBytes: number;
  responseBytes: number;
  /** От отправки запроса до первого байта ответа. */
  firstByteMs: number;
  totalMs: number;
  ok: boolean;
  /** Сколько обращений к провайдеру потребовалось, включая повторы маршрута. */
  attempts: number;
  usage?: Usage;
  at: string;
}

const number = (value: unknown): number => (typeof value === "number" && Number.isFinite(value) ? value : 0);

/** Приводит оба принятых написания к одному виду: chat/completions называет
 * поля prompt и completion, /v1/responses — input и output. */
export function normalizeUsage(raw: unknown): Usage | null {
  if (!raw || typeof raw !== "object") return null;
  const source = raw as Record<string, any>;
  const prompt = number(source.prompt_tokens ?? source.input_tokens);
  const completion = number(source.completion_tokens ?? source.output_tokens);
  const total = number(source.total_tokens) || prompt + completion;
  if (prompt === 0 && completion === 0 && total === 0) return null;

  const reasoning = number(
    source.completion_tokens_details?.reasoning_tokens ?? source.output_tokens_details?.reasoning_tokens
  );
  const cached = number(
    source.prompt_tokens_details?.cached_tokens ?? source.input_tokens_details?.cached_tokens
  );
  return {
    promptTokens: prompt,
    completionTokens: completion,
    totalTokens: total,
    ...(reasoning > 0 ? { reasoningTokens: reasoning } : {}),
    ...(cached > 0 ? { cachedTokens: cached } : {})
  };
}

/** Достаёт usage из строки SSE. Оба интерфейса шлют его в самом конце,
 * причём chat/completions — только если запрошен явно. */
export function extractUsage(line: string): Usage | null {
  const trimmed = line.trim();
  if (!trimmed.startsWith("data:")) return null;
  const payload = trimmed.slice(5).trim();
  if (!payload || payload === "[DONE]") return null;
  if (!payload.includes("usage")) return null;
  let parsed: any;
  try { parsed = JSON.parse(payload); }
  catch { return null; }
  return normalizeUsage(parsed?.usage ?? parsed?.response?.usage);
}

/** Следит за потоком и запоминает последний увиденный usage. Провайдеры
 * присылают его финальным кадром, но порядок кадров не гарантирован. */
export class UsageScanner {
  private tail = "";
  private found: Usage | null = null;

  get usage(): Usage | null {
    return this.found;
  }

  push(text: string): void {
    this.tail += text;
    const lines = this.tail.split(/\r?\n/);
    this.tail = lines.pop() ?? "";
    for (const line of lines) {
      const usage = extractUsage(line);
      if (usage) this.found = usage;
    }
    // Хвост не должен расти бесконечно, если провайдер шлёт поток без переводов
    // строк: тогда usage в нём всё равно не разобрать.
    if (this.tail.length > 1_000_000) this.tail = this.tail.slice(-100_000);
  }

  finish(): Usage | null {
    if (this.tail.trim()) {
      const usage = extractUsage(this.tail);
      if (usage) this.found = usage;
    }
    this.tail = "";
    return this.found;
  }
}

export interface MetricsSummary {
  requests: number;
  failed: number;
  totalMs: number;
  requestBytes: number;
  responseBytes: number;
  promptTokens: number;
  completionTokens: number;
  totalTokens: number;
  /** Обращений без usage: провайдер его не прислал, и счёт токенов неполон. */
  withoutUsage: number;
  byModel: Record<string, { requests: number; totalTokens: number; totalMs: number }>;
}

export const MAX_RETAINED_METRICS = 200;

export class MetricsStore {
  private records: RequestMetric[] = [];

  constructor(private readonly limit = MAX_RETAINED_METRICS) {}

  record(metric: RequestMetric): void {
    this.records.push(metric);
    // Сервер живёт неделями: окно ограничено, иначе память течёт.
    if (this.records.length > this.limit) this.records.splice(0, this.records.length - this.limit);
  }

  all(): RequestMetric[] {
    return [...this.records];
  }

  reset(): void {
    this.records = [];
  }

  summary(): MetricsSummary {
    const byModel: MetricsSummary["byModel"] = {};
    const summary: MetricsSummary = {
      requests: this.records.length,
      failed: 0,
      totalMs: 0,
      requestBytes: 0,
      responseBytes: 0,
      promptTokens: 0,
      completionTokens: 0,
      totalTokens: 0,
      withoutUsage: 0,
      byModel
    };
    for (const item of this.records) {
      if (!item.ok) summary.failed += 1;
      summary.totalMs += item.totalMs;
      summary.requestBytes += item.requestBytes;
      summary.responseBytes += item.responseBytes;
      if (item.usage) {
        summary.promptTokens += item.usage.promptTokens;
        summary.completionTokens += item.usage.completionTokens;
        summary.totalTokens += item.usage.totalTokens;
      } else {
        summary.withoutUsage += 1;
      }
      const key = `${item.provider}/${item.model}`;
      const bucket = byModel[key] ?? (byModel[key] = { requests: 0, totalTokens: 0, totalMs: 0 });
      bucket.requests += 1;
      bucket.totalTokens += item.usage?.totalTokens ?? 0;
      bucket.totalMs += item.totalMs;
    }
    return summary;
  }
}

/** Строка для журнала. Содержит только числа и имена — содержимому книги
 * здесь взяться неоткуда, и это проверено тестом. */
export function formatMetricLine(metric: RequestMetric): string {
  const usage = metric.usage
    ? `tokens=${metric.usage.totalTokens} (prompt ${metric.usage.promptTokens}, completion ${metric.usage.completionTokens}` +
      `${metric.usage.cachedTokens ? `, из кэша ${metric.usage.cachedTokens}` : ""}` +
      `${metric.usage.reasoningTokens ? `, размышления ${metric.usage.reasoningTokens}` : ""})`
    : "tokens=нет данных";
  return `[метрика] ${metric.provider}/${metric.model} ${metric.api} ` +
    `${metric.ok ? "ок" : "сбой"} обращений=${metric.attempts} ` +
    `запрос=${metric.requestBytes}б ответ=${metric.responseBytes}б ` +
    `первый_байт=${metric.firstByteMs}мс всего=${metric.totalMs}мс ${usage}`;
}
