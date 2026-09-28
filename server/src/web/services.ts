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
  id: "tavily" | "serper";
  label: string;
  envKey: string;
  site: string;
}

export const SEARCH_SERVICES: SearchService[] = [
  { id: "tavily", label: "Tavily — поиск в интернете", envKey: "TAVILY_API_KEY", site: "https://tavily.com" },
  { id: "serper", label: "Serper — выдача Google", envKey: "SERPER_API_KEY", site: "https://serper.dev" }
];

export function getSearchService(id: string): SearchService | undefined {
  return SEARCH_SERVICES.find((service) => service.id === id);
}

export function searchKey(service: SearchService): string | undefined {
  return storedKey(service.id) || process.env[service.envKey]?.trim() || undefined;
}

export function searchKeySource(service: SearchService): "panel" | "env" | null {
  if (storedKey(service.id)) return "panel";
  if (process.env[service.envKey]?.trim()) return "env";
  return null;
}

/** Готовые сервисы в порядке предпочтения: Tavily, затем Serper. */
export function readySearchServices(): SearchService[] {
  return SEARCH_SERVICES.filter((service) => searchKey(service));
}
