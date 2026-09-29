/**
 * Ollama на этом компьютере: бесплатная модель без ключа и без облака.
 *
 * Список моделей берётся у самой Ollama, а не из кода: у каждого скачаны
 * свои. Модели без вызова функций (например, для эмбеддингов) агенту не
 * годятся и в список не попадают. Запрос короткий, результат держится
 * несколько секунд: панель спрашивает список часто, а Ollama может быть
 * просто не установлена — тогда провайдера в списке нет, и ничего не ломается.
 */

export interface OllamaModels {
  /** Модели, которые умеют вызывать функции. */
  models: string[];
  /** Ollama ответила, но подходящих моделей нет — панели есть что подсказать. */
  running: boolean;
}

type FetchLike = (url: string, init?: RequestInit) => Promise<Response>;

/** Корень Ollama из адреса OpenAI-совместимого API: http://…:11434/v1 → http://…:11434. */
export function ollamaRoot(baseURL: string): string {
  return baseURL.replace(/\/+$/, "").replace(/\/v1$/, "");
}

async function getJson(fetchImpl: FetchLike, url: string, timeoutMs: number, body?: unknown): Promise<any> {
  const response = await fetchImpl(url, {
    method: body === undefined ? "GET" : "POST",
    headers: body === undefined ? undefined : { "Content-Type": "application/json" },
    body: body === undefined ? undefined : JSON.stringify(body),
    signal: AbortSignal.timeout(timeoutMs)
  });
  if (!response.ok) throw new Error(`HTTP ${response.status}`);
  return response.json();
}

export async function detectOllamaModels(baseURL: string, fetchImpl: FetchLike = fetch, timeoutMs = 1500): Promise<OllamaModels> {
  const root = ollamaRoot(baseURL);
  let tags: any;
  try {
    tags = await getJson(fetchImpl, `${root}/api/tags`, timeoutMs);
  } catch {
    return { models: [], running: false };
  }
  const names: string[] = Array.isArray(tags?.models)
    ? tags.models.map((m: any) => String(m?.name ?? m?.model ?? "")).filter(Boolean)
    : [];
  const models: string[] = [];
  for (const name of names) {
    try {
      const info = await getJson(fetchImpl, `${root}/api/show`, timeoutMs, { model: name });
      const capabilities: unknown = info?.capabilities;
      // Старые Ollama не сообщают возможностей — тогда модель не отсеиваем.
      if (!Array.isArray(capabilities) || capabilities.includes("tools")) models.push(name);
    } catch {
      // Модель не описалась — пропускаем её, остальные остаются.
    }
  }
  return { models: models.sort(), running: true };
}

/** Кэш на несколько секунд, чтобы частые запросы панели не дёргали Ollama. */
export class OllamaWatcher {
  private last: OllamaModels = { models: [], running: false };
  private checkedAt = Number.NEGATIVE_INFINITY;
  private pending: Promise<OllamaModels> | null = null;

  constructor(
    private readonly baseURL: () => string,
    private readonly onChange: (found: OllamaModels) => void,
    private readonly fetchImpl: FetchLike = fetch,
    private readonly ttlMs = 5000,
    private readonly now: () => number = Date.now
  ) {}

  async refresh(): Promise<OllamaModels> {
    if (this.now() - this.checkedAt < this.ttlMs) return this.last;
    if (!this.pending) {
      this.pending = detectOllamaModels(this.baseURL(), this.fetchImpl)
        .then((found) => {
          this.last = found;
          this.checkedAt = this.now();
          this.onChange(found);
          return found;
        })
        .finally(() => { this.pending = null; });
    }
    return this.pending;
  }
}
