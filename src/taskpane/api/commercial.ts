import { streamChat } from "./client";

export const C2_AVAILABLE = typeof __COMMERCIAL_C2__ !== "undefined" && __COMMERCIAL_C2__ === true;
export interface PilotTask { id: string; state: string; steps: number; max_steps: number; cap: string; spent: string; reserved: string; stop_reason?: string; synthetic?: boolean; requests?: Array<{ id: string; state: string }> }
/** Strip workbook location metadata, not the cell values explicitly returned by authorized reads. */
function toolContent(content: string): string {
  try { return JSON.stringify(JSON.parse(content, (key, value) => key === "documentUrl" ? undefined : value)); }
  catch { return content; }
}
export class CommercialClient {
  private token = "";
  private orgId = "";
  organizations: Array<{ id: string; name: string; role: string }> = [];
  constructor(private fetcher: typeof fetch = fetch) {}
  get signedIn() { return Boolean(this.token); }
  get organization() { return this.orgId; }
  selectOrganization(orgId: string) {
    if (!this.organizations.some(o => o.id === orgId)) throw new Error("Организация недоступна.");
    this.orgId = orgId;
  }
  async request(path: string, body?: unknown, signal?: AbortSignal): Promise<any> {
    const res = await this.fetcher(`/api/v1${path}`, { method: body === undefined ? "GET" : "POST", credentials: "omit", cache: "no-store", signal,
      headers: { "Content-Type": "application/json", ...(this.token ? { Authorization: `Bearer ${this.token}` } : {}) }, ...(body === undefined ? {} : { body: JSON.stringify(body) }) });
    if (!res.ok) {
      const data = await res.json().catch(() => null);
      if (res.status === 401) { this.token = ""; this.orgId = ""; this.organizations = []; }
      throw new Error(`Шлюз am.AI: ${data?.error?.code ?? res.status}. Новые действия остановлены; локальная отмена доступна.`);
    }
    return res.status === 204 ? null : res.json();
  }
  async login() {
    const bytes = crypto.getRandomValues(new Uint8Array(32));
    const proof = btoa(String.fromCharCode(...bytes)).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
    const challenge = [...new Uint8Array(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(proof)))].map(b => b.toString(16).padStart(2, "0")).join("");
    const { dialogUrl } = await this.request("/auth/start", { parentChallenge: challenge });
    const url = new URL(dialogUrl);
    if (url.origin !== location.origin || url.pathname !== "/api/v1/auth/dialog") throw new Error("Неверный адрес входа.");
    const ticket = await new Promise<string>((resolve, reject) => {
      Office.context.ui.displayDialogAsync(url.href, { height: 65, width: 45, displayInIframe: false }, result => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) { reject(new Error("Не удалось открыть окно входа Microsoft.")); return; }
        const dialog = result.value;
        const timer = setTimeout(() => { dialog.close(); reject(new Error("Время входа истекло.")); }, 300_000);
        let done = false;
        const fail = () => { if (done) return; done = true; clearTimeout(timer); dialog.close(); reject(new Error("Вход не завершён. Повторите вход Microsoft.")); };
        dialog.addEventHandler(Office.EventType.DialogEventReceived, fail);
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, arg => {
          if (done) return;
          if (!("message" in arg) || !("origin" in arg) || arg.origin !== location.origin) { fail(); return; }
          try {
            const data = JSON.parse(arg.message);
            if (data.type !== "amai-auth" || typeof data.ticket !== "string" || !/^[\w-]{43}$/.test(data.ticket)) { fail(); return; }
            done = true; clearTimeout(timer); dialog.close(); resolve(data.ticket);
          } catch { fail(); }
        });
      });
    });
    await this.redeem(ticket, proof);
  }
  async redeem(ticket: string, proof: string) {
    const session = await this.request("/auth/redeem", { ticket, proof });
    this.token = session.token;
    try {
      const me = await this.request("/me"); this.organizations = me.organizations;
      this.orgId = this.organizations[0]?.id ?? "";
      if (!this.orgId) throw new Error("Нет активного приглашения в пилотную организацию.");
    } catch (e) { await this.logout().catch(() => undefined); throw e; }
  }
  async logout() {
    try { if (this.token) await this.request("/auth/logout", {}); }
    finally { this.token = ""; this.orgId = ""; this.organizations = []; }
  }
  private scope() { if (!this.token || !this.orgId) throw new Error("Войдите в am.AI через Microsoft."); return `/orgs/${this.orgId}`; }
  async budgets() { return this.request(`${this.scope()}/budgets`); }
  async task(id: string): Promise<PilotTask> { return this.request(`${this.scope()}/tasks/${encodeURIComponent(id)}`); }
  async createTask(): Promise<PilotTask> { return this.request(`${this.scope()}/tasks`, { keyMode: "hosted" }); }
  async finishTask(id: string, state: "completed" | "cancelled" | "stopped") { return this.request(`${this.scope()}/tasks/${encodeURIComponent(id)}/finish`, { state }); }
  chatForTask(taskId: string): typeof streamChat {
    const scope = this.scope(), token = this.token;
    return opts => {
      // No retries. The same server task can be inspected after a lost response, never replayed into Excel.
      const requestId = crypto.randomUUID();
      return streamChat({ ...opts, transport: async () => {
        try {
          const response = await this.fetcher(`/api/v1${scope}/chat`, { method: "POST", credentials: "omit", signal: opts.signal,
            headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}`, "Idempotency-Key": requestId },
            body: JSON.stringify({ provider: "fake", model: "c2-pilot", keyMode: "hosted", taskId,
              messages: opts.messages.map(m => m.role === "assistant" ? { role: m.role, content: m.content, ...(m.tool_calls ? { tool_calls: m.tool_calls } : {}) } : m.role === "tool" ? { ...m, content: toolContent(m.content) } : m),
              toolNames: opts.tools.map(t => (t as { function: { name: string } }).function.name) }) });
          if (response.status === 401) { this.token = ""; this.orgId = ""; this.organizations = []; }
          return response;
        } catch {
          if (opts.signal?.aborted) throw new DOMException("Остановлено пользователем", "AbortError");
          throw new Error(`Связь со шлюзом прервана. Задание ${taskId}, запрос ${requestId}. Проверьте состояние; автоматического повтора нет.`);
        }
      } });
    };
  }
}
export const commercial = new CommercialClient();
