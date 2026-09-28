/**
 * Поиск и чтение страниц (этап 8, 8.7). Только чтение: сервер ничего не
 * отправляет на сайты, кроме обычного запроса страницы, и ничего не
 * сохраняет на диск. В сервис поиска уходит только текст запроса.
 */

import { randomUUID } from "node:crypto";
import type { Express } from "express";
import { chunkText } from "../files/types.js";
import { parseInWorker } from "../files/fileStore.js";
import { decodeHtml, htmlToText } from "./html.js";
import { safeFetch, WebFetchError } from "./safeFetch.js";
import { getSearchService, readySearchServices, searchKey, type SearchService } from "./services.js";

export const WEB_LIMITS = { results: 10, queryChars: 400, domains: 20, charsPerRead: 20_000, pagesCached: 20, cacheMs: 30 * 60 * 1000 };

/**
 * «Официальные источники»: регуляторы, статистика, раскрытие информации.
 * Список — отправная точка; пользователь может назвать свои сайты.
 */
export const OFFICIAL_DOMAINS = [
  "cbr.ru", "minfin.gov.ru", "rosstat.gov.ru", "nalog.gov.ru", "economy.gov.ru", "e-disclosure.ru", "moex.com",
  "sec.gov", "federalreserve.gov", "ecb.europa.eu", "imf.org", "worldbank.org", "oecd.org", "bis.org", "eurostat.ec.europa.eu"
];

export interface SearchResult { title: string; url: string; snippet: string; published?: string }

export interface SearchRequest { query: string; domains?: string[]; service?: string; max?: number }

type Fetcher = (url: string, init: { method: string; headers: Record<string, string>; body: string }) => Promise<{ ok: boolean; status: number; json(): Promise<any> }>;

function cleanDomains(domains: unknown): string[] {
  if (!Array.isArray(domains)) return [];
  return [...new Set(domains.map((item) => String(item).trim().toLowerCase().replace(/^https?:\/\//, "").replace(/\/.*$/, "").replace(/^www\./, ""))
    .filter((item) => /^[a-z0-9.-]+\.[a-z]{2,}$/.test(item)))].slice(0, WEB_LIMITS.domains);
}

function upstreamError(service: SearchService, status: number): WebFetchError {
  if (status === 401 || status === 403) return new WebFetchError(`${service.label}: ключ не подошёл. Проверьте его в «Ключах».`);
  if (status === 429 || status === 432 || status === 402) return new WebFetchError(`${service.label}: лимит запросов исчерпан — пополните тариф или подключите второй сервис в «Ключах».`);
  return new WebFetchError(`${service.label} ответил ошибкой ${status}.`);
}

export async function webSearch(request: SearchRequest, fetcher: Fetcher = fetch as any): Promise<{ service: string; query: string; domains: string[]; results: SearchResult[] }> {
  const query = String(request.query ?? "").replace(/\s+/g, " ").trim().slice(0, WEB_LIMITS.queryChars);
  if (!query) throw new WebFetchError("Пустой запрос поиска.");
  const ready = readySearchServices();
  if (!request.service && !ready.length) throw new WebFetchError("Не подключён сервис поиска: добавьте ключ Tavily или Serper в «Ключах».");
  const service = request.service ? getSearchService(request.service) : ready[0];
  if (!service) throw new WebFetchError(`Сервиса поиска «${request.service}» нет. Есть: tavily, serper.`);
  const key = searchKey(service);
  if (!key) throw new WebFetchError("Не подключён сервис поиска: добавьте ключ Tavily или Serper в «Ключах».");
  const domains = cleanDomains(request.domains);
  const max = Math.min(Math.max(1, Math.floor(request.max ?? 5)), WEB_LIMITS.results);

  if (service.id === "tavily") {
    const response = await fetcher("https://api.tavily.com/search", {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${key}` },
      body: JSON.stringify({ query, max_results: max, search_depth: "basic", include_answer: false, include_raw_content: false, ...(domains.length ? { include_domains: domains } : {}) })
    });
    if (!response.ok) throw upstreamError(service, response.status);
    const data = await response.json();
    const results = (Array.isArray(data?.results) ? data.results : []).map((item: any) => ({
      title: String(item?.title ?? ""), url: String(item?.url ?? ""), snippet: String(item?.content ?? "").slice(0, 600),
      ...(item?.published_date ? { published: String(item.published_date) } : {})
    }));
    return { service: service.label, query, domains, results };
  }

  const response = await fetcher("https://google.serper.dev/search", {
    method: "POST",
    headers: { "Content-Type": "application/json", "X-API-KEY": key },
    body: JSON.stringify({ q: domains.length ? `${query} (${domains.map((domain) => `site:${domain}`).join(" OR ")})` : query, num: max, hl: "ru" })
  });
  if (!response.ok) throw upstreamError(service, response.status);
  const data = await response.json();
  const results = (Array.isArray(data?.organic) ? data.organic : []).slice(0, max).map((item: any) => ({
    title: String(item?.title ?? ""), url: String(item?.link ?? ""), snippet: String(item?.snippet ?? "").slice(0, 600),
    ...(item?.date ? { published: String(item.date) } : {})
  }));
  return { service: service.label, query, domains, results };
}

interface CachedPage { id: string; url: string; title: string; kind: "html" | "pdf" | "text"; fetchedAt: string; parts: string[]; at: number }

export class PageCache {
  private pages = new Map<string, CachedPage>();

  constructor(private fetchPage: typeof safeFetch = safeFetch, private now: () => number = Date.now) {}

  private sweep() {
    for (const [key, page] of this.pages) if (this.now() - page.at > WEB_LIMITS.cacheMs) this.pages.delete(key);
    while (this.pages.size > WEB_LIMITS.pagesCached) this.pages.delete(this.pages.keys().next().value!);
  }

  async open(url: string): Promise<CachedPage> {
    this.sweep();
    const cached = this.pages.get(url);
    if (cached) return cached;
    const fetched = await this.fetchPage(url);
    const type = fetched.contentType.toLowerCase();
    let page: CachedPage;
    const base = { id: randomUUID(), url: fetched.url, fetchedAt: new Date(this.now()).toISOString(), at: this.now() };
    if (type.includes("application/pdf") || fetched.body.subarray(0, 5).toString("latin1") === "%PDF-") {
      const parsed = await parseInWorker(fetched.body, "page.pdf");
      page = { ...base, title: fetched.url, kind: "pdf", parts: parsed.text };
    } else if (type.includes("html") || type.includes("xml") || !type) {
      const { title, text } = htmlToText(decodeHtml(fetched.body, type));
      if (!text.trim()) throw new WebFetchError("На странице нет текста — возможно, она собирается скриптами, которые панель не выполняет.");
      page = { ...base, title, kind: "html", parts: chunkText(text) };
    } else if (type.startsWith("text/")) {
      page = { ...base, title: fetched.url, kind: "text", parts: chunkText(decodeHtml(fetched.body, type)) };
    } else {
      throw new WebFetchError(`Страница — не текст (${type.split(";")[0]}): такие не читаются.`);
    }
    this.pages.set(url, page);
    return page;
  }

  /** Часть страницы для модели — до charsPerRead знаков, с продолжением. */
  async read(url: string, from = 0) {
    const page = await this.open(url);
    const start = Math.max(0, Math.floor(from));
    const out: string[] = [];
    let chars = 0;
    let index = start;
    for (; index < page.parts.length; index++) {
      if (out.length && chars + page.parts[index].length > WEB_LIMITS.charsPerRead) break;
      chars += page.parts[index].length;
      out.push(page.parts[index]);
    }
    return {
      url: page.url, title: page.title, kind: page.kind, fetchedAt: page.fetchedAt, parts: page.parts.length,
      from: start + 1, to: index, text: out.join("\n"), ...(index < page.parts.length ? { continueFrom: index } : {})
    };
  }
}

export function registerWebRoutes(app: Express, cache = new PageCache()): void {
  const fail = (res: any, error: any) => res.status(error instanceof WebFetchError ? 400 : 500).json({ error: { message: String(error?.message ?? error) } });
  app.get("/api/web", (_req, res) => res.json({ services: readySearchServices().map((service) => ({ id: service.id, label: service.label })), official: OFFICIAL_DOMAINS }));
  app.post("/api/web/search", (req, res) => {
    const body = req.body ?? {};
    const domains = body.official ? [...OFFICIAL_DOMAINS, ...(Array.isArray(body.domains) ? body.domains : [])] : body.domains;
    webSearch({ query: body.query, domains, service: body.service, max: body.max }).then((result) => res.json(result), (error) => fail(res, error));
  });
  app.post("/api/web/page", (req, res) => {
    cache.read(String(req.body?.url ?? ""), Number(req.body?.from ?? 0)).then((result) => res.json(result), (error) => fail(res, error));
  });
}
