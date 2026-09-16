/**
 * Выбор reasoning_effort для запроса к провайдеру.
 *
 * Агенту нужен быстрый цикл «вызов инструмента — ответ», поэтому у OpenAI мы
 * просим минимум размышлений. Но набор допустимых значений различается от
 * модели к модели: `gpt-6-astra` отвергает `none` и требует хотя бы `low`.
 * Держать список моделей в коде — значит ломаться на каждой следующей, поэтому
 * непринятое значение выясняется из ответа самого провайдера и запоминается.
 */

export const DEFAULT_OPENAI_EFFORT = "none";
/** Запасной вариант, если разобрать список допустимых значений не удалось. */
export const FALLBACK_OPENAI_EFFORT = "low";

const PREFERENCE = ["none", "low", "medium", "high", "xhigh"];

const learned = new Map<string, string>();

function key(providerId: string, model: string): string {
  return `${providerId}:${model}`;
}

/** Что просить у этой модели сейчас: выученное значение либо желаемое. */
export function reasoningEffortFor(providerId: string, model: string, preferred = DEFAULT_OPENAI_EFFORT): string {
  return learned.get(key(providerId, model)) ?? preferred;
}

/** Достаёт допустимые значения из текста ошибки провайдера.
 * Формат: Supported values are: 'low', 'medium', 'high', and 'xhigh'. */
export function parseSupportedEfforts(message: string): string[] {
  const tail = message.slice(message.toLowerCase().indexOf("supported values"));
  if (!tail) return [];
  const found = [...tail.matchAll(/'([a-z]+)'/gi)].map((match) => match[1].toLowerCase());
  return found.filter((value) => PREFERENCE.includes(value));
}

/** Самое дешёвое из допустимых: цикл агента не выигрывает от долгих раздумий. */
export function cheapestSupportedEffort(supported: string[]): string {
  for (const candidate of PREFERENCE) if (supported.includes(candidate)) return candidate;
  return FALLBACK_OPENAI_EFFORT;
}

export interface UpstreamErrorShape {
  error?: { message?: string; param?: string; code?: string };
}

/** Распознаёт именно отказ по reasoning_effort, а не любую ошибку 400.
 * Тело приходит от стороннего сервиса, поэтому разбирается осторожно. */
export function rejectedReasoningEffort(status: number, body: string): string | null {
  if (status !== 400) return null;
  let parsed: UpstreamErrorShape;
  try { parsed = JSON.parse(body) as UpstreamErrorShape; }
  catch { return null; }
  const error = parsed?.error;
  if (!error || typeof error !== "object") return null;
  if (error.param !== "reasoning_effort") return null;
  const message = typeof error.message === "string" ? error.message : "";
  return cheapestSupportedEffort(parseSupportedEfforts(message));
}

/** Запоминает рабочее значение, чтобы следующий запрос не повторял ошибку. */
export function rememberEffort(providerId: string, model: string, effort: string): void {
  learned.set(key(providerId, model), effort);
}

export function resetLearnedEfforts(): void {
  learned.clear();
}
