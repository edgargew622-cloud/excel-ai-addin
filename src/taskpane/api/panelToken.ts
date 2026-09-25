/**
 * Токен панели (этап 8, 8.0.1): без него локальный сервер не отвечает на /api.
 *
 * Excel открывает панель по адресу из манифеста каталога, и туда при
 * установке дописан ?t=<токен>. Каталог лежит в папке надстройки, закрытой
 * для других пользователей компьютера. Копия в sessionStorage — на случай,
 * если панель перезагрузится без параметра.
 */

const STORAGE_KEY = "excel-ai-addin.panel-token";
let cached: string | null = null;

export function panelToken(): string {
  if (cached !== null) return cached;
  let token = "";
  try { token = new URLSearchParams(globalThis.location?.search ?? "").get("t") ?? ""; } catch { /* нет адреса — тесты */ }
  try {
    if (token) sessionStorage.setItem(STORAGE_KEY, token);
    else token = sessionStorage.getItem(STORAGE_KEY) ?? "";
  } catch { /* хранилище недоступно */ }
  cached = token;
  return token;
}

export function apiHeaders(): Record<string, string> {
  const token = panelToken();
  return { "Content-Type": "application/json", ...(token ? { "X-Panel-Token": token } : {}) };
}
