/**
 * Сервисы поиска в интернете (этап 8, 8.7): Tavily — основной, Serper
 * (выдача Google) — второй. Ключ у каждого пользователя свой и хранится
 * там же, где ключи провайдеров моделей: в «Ключах», зашифрованным.
 */

let storedKey: (id: string) => string | undefined = () => undefined;

export function setSearchKeyLookup(lookup: (id: string) => string | undefined): void {
  storedKey = lookup;
}

export interface SearchService {
  id: "tavily" | "serper" | "deepseek";
  label: string;
  envKey: string;
  site: string;
}

export const SEARCH_SERVICES: SearchService[] = [
  { id: "tavily", label: "Tavily — поиск в интернете", envKey: "TAVILY_API_KEY", site: "https://tavily.com" },
  { id: "serper", label: "Serper — выдача Google", envKey: "SERPER_API_KEY", site: "https://serper.dev" }
];

/**
 * Поиск DeepSeek (22.10.2026): свой поисковик DeepSeek по тому же ключу, что и
 * модель, — отдельный ключ поиска не нужен. В «Ключах» отдельной строкой не
 * показывается: это ключ провайдера DeepSeek. $0,75 за 1000 запросов.
 */
export const DEEPSEEK_SEARCH: SearchService = { id: "deepseek", label: "DeepSeek — поиск", envKey: "DEEPSEEK_API_KEY", site: "https://api-docs.deepseek.com/guides/search/" };

export function getSearchService(id: string): SearchService | undefined {
  return [...SEARCH_SERVICES, DEEPSEEK_SEARCH].find((service) => service.id === id);
}

export function searchKey(service: SearchService): string | undefined {
  return storedKey(service.id) || process.env[service.envKey]?.trim() || undefined;
}

export function searchKeySource(service: SearchService): "panel" | "env" | null {
  if (storedKey(service.id)) return "panel";
  if (process.env[service.envKey]?.trim()) return "env";
  return null;
}

/**
 * Готовые сервисы в порядке предпочтения: DeepSeek (тот же ключ, что у модели,
 * даты публикаций, глубокие ссылки на cbr.ru и armstat — замер 22.10.2026),
 * затем Tavily и Serper — запасные, если DeepSeek не ответил.
 */
export function readySearchServices(): SearchService[] {
  return [DEEPSEEK_SEARCH, ...SEARCH_SERVICES].filter((service) => searchKey(service));
}
