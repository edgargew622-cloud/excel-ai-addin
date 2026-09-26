/**
 * Проверка новой версии (этап 8, 8.8.3).
 *
 * Сервер не чаще раза в сутки спрашивает у GitHub номер последнего выпуска
 * и сравнивает со своей версией. GitHub при этом видит только адрес
 * компьютера: ни данных книги, ни ключей в запросе нет. Отключается строкой
 * UPDATE_CHECK=off в server/.env. Без сети проверка молча не срабатывает —
 * работе панели это не мешает.
 */

export const RELEASES_API = "https://api.github.com/repos/edgargew622-cloud/excel-ai-addin/releases/latest";
const DAY_MS = 24 * 60 * 60_000;

/** «v1.0.3» и «1.0.3» → [1, 0, 3]; не версия — null. */
export function parseVersion(text: unknown): number[] | null {
  const match = /^v?(\d+)\.(\d+)\.(\d+)$/.exec(String(text ?? "").trim());
  return match ? match.slice(1).map(Number) : null;
}

/** Положительное — a новее b. */
export function compareVersions(a: number[], b: number[]): number {
  for (let index = 0; index < 3; index++) if (a[index] !== b[index]) return a[index] - b[index];
  return 0;
}

export interface UpdateInfo {
  current: string;
  checked: boolean;
  latest?: string;
  url?: string;
  newer?: boolean;
  disabled?: boolean;
}

export class UpdateChecker {
  private cache: { at: number; info: UpdateInfo } | null = null;

  constructor(
    private readonly current: string,
    private readonly fetcher: (url: string, init: RequestInit) => Promise<Response> = fetch,
    private readonly now: () => number = Date.now,
    private readonly disabled = false
  ) {}

  async check(): Promise<UpdateInfo> {
    if (this.disabled) return { current: this.current, checked: false, disabled: true };
    if (this.cache && this.now() - this.cache.at < DAY_MS) return this.cache.info;
    let info: UpdateInfo = { current: this.current, checked: false };
    try {
      const res = await this.fetcher(RELEASES_API, {
        headers: { Accept: "application/vnd.github+json", "User-Agent": "excel-ai-addin-update-check" },
        signal: AbortSignal.timeout(8_000)
      });
      if (res.ok) {
        const data = (await res.json()) as { tag_name?: string; html_url?: string };
        const latest = parseVersion(data.tag_name);
        const mine = parseVersion(this.current);
        if (latest && mine) {
          info = {
            current: this.current,
            checked: true,
            latest: latest.join("."),
            url: typeof data.html_url === "string" && data.html_url.startsWith("https://github.com/") ? data.html_url : undefined,
            newer: compareVersions(latest, mine) > 0
          };
        }
      }
    } catch { /* нет сети или GitHub недоступен — не мешаем работе */ }
    // Неудачу тоже запоминаем, но на час: не стучаться при каждом открытии панели.
    this.cache = { at: info.checked ? this.now() : this.now() - DAY_MS + 60 * 60_000, info };
    return info;
  }
}
