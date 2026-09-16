/**
 * Выбор интерфейса и reasoning_effort для OpenAI.
 *
 * Агенту нужен быстрый цикл «вызов инструмента — ответ», поэтому по умолчанию
 * просим минимум размышлений на /v1/chat/completions. Но у моделей разные
 * ограничения, и `gpt-6-astra` на chat/completions с функциями недоступна
 * вообще: замер 16 сентября 2026 года дал три отказа подряд —
 *
 *   none    → «does not support 'none' with this model. Supported: low, medium…»
 *   low     → «Function tools with reasoning_effort are not supported… use /v1/responses»
 *   без параметра → тот же отказ про функции, действует значение по умолчанию
 *
 * Выход указан в самой ошибке: /v1/responses. Список моделей в коде ломался бы
 * на каждой следующей, поэтому маршрут выясняется из ответов провайдера
 * и запоминается для этой модели до перезапуска сервера.
 */

export interface OpenAiRoute {
  /** Какой интерфейс использовать. */
  api: "chat" | "responses";
  /** Значение reasoning_effort; null — не отправлять параметр вовсе. */
  effort: string | null;
}

export const DEFAULT_ROUTE: OpenAiRoute = { api: "chat", effort: "none" };

/** Дешёвые значения идут первыми: долгие раздумья циклу инструментов не нужны. */
const PREFERENCE = ["none", "low", "medium", "high", "xhigh"];

const learned = new Map<string, OpenAiRoute>();

const key = (providerId: string, model: string) => `${providerId}:${model}`;
const sameRoute = (a: OpenAiRoute, b: OpenAiRoute) => a.api === b.api && a.effort === b.effort;

export function routeFor(providerId: string, model: string): OpenAiRoute {
  return learned.get(key(providerId, model)) ?? DEFAULT_ROUTE;
}

export function rememberRoute(providerId: string, model: string, route: OpenAiRoute): void {
  learned.set(key(providerId, model), route);
}

export function resetLearnedRoutes(): void {
  learned.clear();
}

/** Достаёт допустимые значения из текста ошибки.
 * Формат: Supported values are: 'low', 'medium', 'high', and 'xhigh'. */
export function parseSupportedEfforts(message: string): string[] {
  const at = message.toLowerCase().indexOf("supported values");
  if (at < 0) return [];
  return [...message.slice(at).matchAll(/'([a-z]+)'/gi)]
    .map((match) => match[1].toLowerCase())
    .filter((value) => PREFERENCE.includes(value));
}

export function cheapestSupportedEffort(supported: string[]): string | null {
  for (const candidate of PREFERENCE) if (supported.includes(candidate)) return candidate;
  return null;
}

interface UpstreamErrorShape {
  error?: { message?: string; param?: string; code?: string };
}

export interface RouteRetry {
  route: OpenAiRoute;
  reason: string;
}

/**
 * Разбирает отказ и предлагает следующий маршрут. Возвращает null, если ошибка
 * не про это или пробовать больше нечего. `tried` не даёт ходить по кругу
 * между противоречащими друг другу отказами.
 */
export function nextRouteAfterRejection(
  status: number,
  body: string,
  current: OpenAiRoute,
  tried: readonly OpenAiRoute[] = []
): RouteRetry | null {
  if (status !== 400) return null;
  let parsed: UpstreamErrorShape;
  try { parsed = JSON.parse(body) as UpstreamErrorShape; }
  catch { return null; }
  const message = typeof parsed?.error?.message === "string" ? parsed.error.message : "";
  if (!message) return null;

  const seen = [current, ...tried];
  const unseen = (route: OpenAiRoute, reason: string): RouteRetry | null =>
    seen.some((item) => sameRoute(item, route)) ? null : { route, reason };

  // Провайдер прямо называет пригодный интерфейс — доверяем этому указанию.
  if (message.includes("/v1/responses")) {
    return unseen({ api: "responses", effort: null }, "модель требует /v1/responses для функций");
  }

  // Отказ по значению параметра: берём самое дешёвое из названных допустимых.
  if (parsed.error?.param === "reasoning_effort") {
    const cheapest = cheapestSupportedEffort(parseSupportedEfforts(message));
    if (cheapest) {
      const retry = unseen({ api: current.api, effort: cheapest }, `модель назвала допустимым ${cheapest}`);
      if (retry) return retry;
    }
    return unseen({ api: current.api, effort: null }, "модель отвергает параметр при любом значении");
  }
  return null;
}
