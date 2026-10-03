import { useEffect, useRef, useState } from "react";
import { fetchProviders, fetchUpdate, type ChatMessage, type ProviderInfo, type UpdateInfo } from "./api/client";
import KeysPanel from "./KeysPanel";
import { stepCost } from "./api/prices";
import { WEB_PANEL } from "./panelMode";
import MemoryPanel from "./MemoryPanel";
import ConversationsPanel from "./ConversationsPanel";
import { Icon, Logo, Markdown, Sparkle, Wordmark, actionsWord, toolLabel } from "./chatView";
import { CATEGORY_TEXT, fetchMemory, memoryPrompt, type Scenario } from "./api/memory";
import { apiHeaders } from "./api/panelToken";
import { documentConversationId, documentConversationKey, ensureDocumentConversationId } from "./documentId";
import { describeFile, filesPrompt, listFiles, removeFile, uploadFile, type AttachedFile } from "./api/files";
import { runAgent, type ToolEvent } from "../agent/loop";
import {
  depth as undoDepth,
  describeNextUndo,
  isCustomUndoAvailable,
  subscribeUndoAvailability,
  undoLast
} from "../excel/undo";
import { ensureStructuralChangeMonitor, subscribeStructuralInvalidation } from "../excel/workbookEvents";
import { getActiveContext } from "../excel/workbookContext";
import {
  conversationIdentity,
  deleteConversation,
  bindingDecision,
  loadConversation,
  saveConversation,
  type PersistedEntry
} from "./conversationStore";
import { BUSY, busyReason, createWorkbookLock, withWorkbookLock, type LockOwner } from "./workbookLock";

type Entry = PersistedEntry;

interface Pending {
  name: string;
  args: unknown;
  resolve: (ok: boolean) => void;
}

/** Вытаскиваем адрес из аргументов, чтобы показать его отдельно — это главное в операции. */
function addressOf(args: unknown): string {
  const a = args as Record<string, unknown> | null;
  if (!a || typeof a !== "object") return "";
  const sheet = typeof a.sheet === "string" && a.sheet ? `${a.sheet}!` : "";
  const targetSheet = a.target && typeof a.target === "object" && typeof (a.target as any).sheetName === "string"
    ? `${(a.target as any).sheetName}!`
    : "";
  const addr = (a.address ?? a.sourceAddress ?? a.destAddress) as string | undefined;
  if (addr) return `${sheet || targetSheet}${addr}`;
  if (typeof a.startRow === "number") return `${sheet}строки ${a.startRow}–${a.startRow + Number(a.count ?? 1) - 1}`;
  return "";
}

type EntryGroup =
  | { kind: "ops"; start: number; items: Array<Extract<Entry, { kind: "op" }>> }
  | { kind: "single"; start: number; entry: Exclude<Entry, { kind: "op" }> };

/** Подряд идущие действия агента — одной сворачиваемой строкой. */
function groupEntries(entries: readonly Entry[]): EntryGroup[] {
  const groups: EntryGroup[] = [];
  entries.forEach((entry, index) => {
    const last = groups[groups.length - 1];
    if (entry.kind === "op") {
      if (last?.kind === "ops") last.items.push(entry);
      else groups.push({ kind: "ops", start: index, items: [entry] });
    } else groups.push({ kind: "single", start: index, entry });
  });
  return groups;
}

const PANEL_BUILD = typeof __PANEL_BUILD__ === "string" ? __PANEL_BUILD__ : "разработка";
const PANEL_VERSION = typeof __PANEL_VERSION__ === "string" ? __PANEL_VERSION__ : "разработка";

/**
 * Имя собственного файла панели. В собранной панели это taskpane-<хеш>.js;
 * при разработке — исходник, и сверять его не с чем.
 */
function ownBundle(): string | null {
  try {
    const name = new URL(import.meta.url).pathname.split("/").pop() ?? "";
    return /^taskpane-[\w-]+\.js$/.test(name) ? name : null;
  } catch {
    return null;
  }
}

/**
 * Устарела ли загруженная панель.
 *
 * Excel держит панель открытой, пока её не закроют, и новая сборка на сервере
 * её не касается. Проверки 18 сентября 2026 года трижды шли на старой панели:
 * агент не видел новых инструментов, и это выяснялось лишь по его ответам.
 * Сервер отдаёт taskpane.html текущей сборки — если в нём другой файл
 * панели, загруженная устарела.
 */
async function panelIsStale(): Promise<boolean> {
  const own = ownBundle();
  if (!own) return false;
  try {
    const response = await fetch("taskpane.html", { cache: "no-store" });
    if (!response.ok) return false;
    return !(await response.text()).includes(own);
  } catch {
    return false;
  }
}

/** «Расход: 13 обращений · 420 тыс. токенов (из кэша 380 тыс.) · $0,21». */
export function spendingNote(spent: { calls: number; prompt: number; cached: number; completion: number; cost: number; costKnown: boolean; estimated?: boolean }): string {
  const thousands = (n: number) => (n >= 10_000 ? `${Math.round(n / 1000)} тыс.` : n >= 1000 ? `${(n / 1000).toFixed(1).replace(".", ",")} тыс.` : String(n));
  const calls = `${spent.calls} ${spent.calls % 10 === 1 && spent.calls % 100 !== 11 ? "обращение" : [2, 3, 4].includes(spent.calls % 10) && ![12, 13, 14].includes(spent.calls % 100) ? "обращения" : "обращений"}`;
  const cached = spent.cached ? ` (из кэша ${thousands(spent.cached)})` : "";
  // «≈» — цена по прайсу панели (prices.ts), без него — цена от поставщика.
  const money = spent.costKnown
    ? ` · ${spent.estimated ? "≈ " : ""}$${spent.cost < 0.01 ? spent.cost.toFixed(4) : spent.cost.toFixed(2)}`.replace(".", ",")
    : " · цену считает поставщик";
  return `Расход задачи: ${calls} к модели · ${thousands(spent.prompt + spent.completion)} токенов${cached}${money}.`;
}

/** Итог беседы — из строк «Расход задачи: …» в ленте: так он переживает перезагрузку панели. */
export function conversationSpending(entries: readonly { kind: string; text?: string; cost?: number }[]): { calls: number; cost: number; costKnown: boolean; estimated: boolean; tasks: number } {
  const total = { calls: 0, cost: 0, costKnown: true, estimated: false, tasks: 0 };
  for (const entry of entries) {
    if (entry.kind !== "notice" || !entry.text?.startsWith("Расход задачи:")) continue;
    const task = entry.text.split(" Всего за беседу")[0];
    const calls = /Расход задачи: (\d+)/.exec(task);
    const cost = /\$(\d+(?:,\d+)?)/.exec(task);
    total.tasks += 1;
    total.calls += calls ? Number(calls[1]) : 0;
    if (cost) {
      // Точная цена, если сохранена: сумма округлённых строк расходилась
      // с журналом (02.10.2026: $0,15 вместо $0,16).
      total.cost += entry.cost ?? Number(cost[1].replace(",", "."));
      if (task.includes("≈")) total.estimated = true;
    }
    else total.costKnown = false;
  }
  return total;
}

function moneyText(cost: number): string {
  return `$${cost < 0.01 ? cost.toFixed(4) : cost.toFixed(2)}`.replace(".", ",");
}

export function conversationTotalText(total: { calls: number; cost: number; costKnown: boolean; estimated?: boolean }): string {
  const word = total.calls % 10 === 1 && total.calls % 100 !== 11 ? "обращение" : [2, 3, 4].includes(total.calls % 10) && ![12, 13, 14].includes(total.calls % 100) ? "обращения" : "обращений";
  return `${total.calls} ${word}${total.costKnown ? ` · ${total.estimated ? "≈ " : ""}${moneyText(total.cost)}` : total.cost ? ` · от ${moneyText(total.cost)} (часть цен считает поставщик)` : ""}`;
}

export default function Taskpane() {
  const [providers, setProviders] = useState<ProviderInfo[]>([]);
  const [apiStatus, setApiStatus] = useState<"checking" | "ready" | "error">("checking");
  const [showKeys, setShowKeys] = useState(false);
  // Интернет (8.7): выключен по умолчанию; включение — явное действие пользователя.
  const [webEnabled, setWebEnabled] = useState(() => { try { return localStorage.getItem("amai.web") === "1"; } catch { return false; } });
  const [webServices, setWebServices] = useState<string[] | null>(null);
  // Прикреплённые файлы (8.6): лежат в памяти локального сервера.
  const [files, setFiles] = useState<AttachedFile[]>([]);
  const [uploading, setUploading] = useState(false);
  const fileInput = useRef<HTMLInputElement>(null);
  const [showMemory, setShowMemory] = useState(false);
  const [showConversations, setShowConversations] = useState(false);
  const [apiError, setApiError] = useState("");
  const [provider, setProvider] = useState("");
  const [model, setModel] = useState("");
  const [entries, setEntries] = useState<Entry[]>([]);
  const [draft, setDraft] = useState("");
  // Владелец изменений книги один: задача и отмена исключают друг друга.
  // Замок захватывается синхронно, поэтому два нажатия подряд не запускают
  // два цикла, даже пока React не перерисовал панель.
  const lock = useRef(createWorkbookLock());
  const [lockOwner, setLockOwner] = useState<LockOwner | null>(null);
  const busy = lockOwner !== null;
  const taskRunning = lockOwner === "task";
  const [streaming, setStreaming] = useState("");
  const [pending, setPending] = useState<Pending | null>(null);
  const [canUndo, setCanUndo] = useState(false);
  const [undoLabel, setUndoLabel] = useState<string | null>(null);
  const [undoAvailable, setUndoAvailable] = useState(false);
  const [monitorStatus, setMonitorStatus] = useState<"connecting" | "ready" | "unsupported" | "error">("connecting");
  const [analysisOnly, setAnalysisOnly] = useState(true);
  const [contextLabel, setContextLabel] = useState("Книга: проверка…");
  const [persistenceNote, setPersistenceNote] = useState("История: проверка привязки…");
  const [stale, setStale] = useState(false);
  const [update, setUpdate] = useState<UpdateInfo | null>(null);

  const history = useRef<ChatMessage[]>([]);
  const workbookBinding = useRef<{ key: string; url: string } | null>(null);
  const persistenceReady = useRef(false);
  /** Панель уже знает свою книгу: дальше смена адреса — это её сохранение. */
  const bindingInitialized = useRef(false);
  const abort = useRef<AbortController | null>(null);
  const logEnd = useRef<HTMLDivElement>(null);

  /** Какие сервисы поиска готовы (8.7): без ключа галочка «Интернет» недоступна. */
  async function loadWeb() {
    // Панель из интернета (Mac): поиск живёт в локальном сервере — его нет.
    if (WEB_PANEL) { setWebServices([]); return; }
    try {
      const response = await fetch("/api/web", { headers: apiHeaders() });
      const data = await response.json();
      setWebServices(response.ok ? (data.services ?? []).map((service: { label: string }) => service.label.split(" — ")[0]) : []);
    } catch {
      setWebServices([]);
    }
  }

  useEffect(() => {
    void loadProviders();
    void loadWeb();
    void refreshContext();
    void panelIsStale().then(setStale);
    void fetchUpdate().then((info) => setUpdate(info?.newer ? info : null));
    return lock.current.subscribe(setLockOwner);
  }, []);

  useEffect(() => {
    if (!persistenceReady.current || !workbookBinding.current) return;
    const binding = workbookBinding.current;
    const title = entries.find((entry) => entry.kind === "user")?.text.slice(0, 80) || "Новая беседа";
    const timer = window.setTimeout(() => {
      const saved = saveConversation(localStorage, {
        workbookKey: binding.key,
        documentUrl: binding.url,
        title,
        entries,
        history: history.current
      });
      if (!saved) setPersistenceNote("История не сохранена: достигнут локальный лимит 2 МБ.");
    }, 250);
    return () => window.clearTimeout(timer);
  }, [entries]);

  async function refreshContext() {
    try {
      let permissionsReset = false;
      const context = await getActiveContext();
      const url = context.workbook.documentUrl;
      const book = url ? decodeURIComponent(url.split(/[\\/]/).pop() || url) : "несохранённая книга";
      setContextLabel(`${book} · ${context.activeSheet.name} · ${context.selectedAreas.join(", ") || "выделена не ячейка"}`);
      // Номер беседы в самой книге важнее адреса: адрес меняется при сохранении.
      const documentId = documentConversationId();
      const urlKey = conversationIdentity(url);
      const key = documentId ? documentConversationKey(documentId) : urlKey;
      const decision = key ? bindingDecision(workbookBinding.current, bindingInitialized.current, key) : "keep";
      if (!key) {
        // Книга без адреса: беседа живёт, пока открыта панель, и не
        // восстанавливается при следующем открытии.
        workbookBinding.current = null;
        persistenceReady.current = true;
        bindingInitialized.current = true;
        setPersistenceNote("Несохранённая книга: беседа не восстанавливается автоматически.");
      } else if (decision === "migrate") {
        // Та же книга сохранена впервые или под новым именем: беседа остаётся,
        // сохраняется под новым адресом, права записи не сбрасываются.
        workbookBinding.current = { key, url };
        persistenceReady.current = true;
        setPersistenceNote("Беседа хранится локально 30 дней; ответы инструментов могут содержать данные ячеек.");
        setEntries((current) => [...current]);
      } else if (decision === "load") {
        bindingInitialized.current = true;
        persistenceReady.current = false;
        workbookBinding.current = { key, url };
        // Беседа, сохранённая до номера в книге, — по адресу; дальше живёт под номером.
        const restored = loadConversation(localStorage, key) ?? (documentId && urlKey ? loadConversation(localStorage, urlKey) : null);
        history.current = restored
          ? [
              {
                role: "system",
                content: "Эта беседа восстановлена после закрытия панели. Все прежние сведения о книге исторические и могут быть устаревшими; перечитайте необходимые диапазоны. Не повторяйте прежние записи автоматически."
              },
              ...restored.history
            ]
          : [];
        setEntries(restored?.entries ?? []);
        setAnalysisOnly(true);
        permissionsReset = true;
        persistenceReady.current = true;
        setPersistenceNote(restored
          ? "Беседа восстановлена локально; данные книги считаются устаревшими, запись отключена."
          : "Беседа хранится локально 30 дней; ответы инструментов могут содержать данные ячеек.");
      }
      return { context, permissionsReset };
    } catch (error: any) {
      setContextLabel(`Контекст книги недоступен: ${error?.message ?? String(error)}`);
      return null;
    }
  }

  function loadProviders() {
    setApiStatus("checking");
    setApiError("");
    return fetchProviders()
      .then((list) => {
        setApiStatus("ready");
        setProviders(list);
        if (list.length) {
          // После правки ключей выбор пользователя сохраняется, если он ещё доступен.
          const keep = list.find((p) => p.id === provider);
          const chosen = keep ?? list[0];
          setProvider(chosen.id);
          setModel(keep && keep.models.includes(model) ? model : chosen.defaultModel);
        } else {
          setProvider("");
          setModel("");
          setApiError("Нет ни одного ключа провайдера. Добавьте хотя бы один в «Ключи».");
          setShowKeys(true);
        }
      })
      .catch((err) => {
        setApiStatus("error");
        setApiError(`Локальный API недоступен: ${String(err.message ?? err)}`);
      });
  }

  useEffect(() => {
    const unsubscribeInvalidation = subscribeStructuralInvalidation((notice) => {
      refreshUndoState();
      setEntries((e) => [
        ...e,
        {
          kind: "assistant",
          text:
            `Обнаружено структурное изменение Excel (${notice.changeType}` +
            `${notice.address ? `, ${notice.address}` : ""}). ` +
            `История custom undo очищена (${notice.removedUndo}), чтобы старый адрес не затронул другую ячейку.`
        }
      ]);
    });
    const unsubscribeAvailability = subscribeUndoAvailability((ready) => {
      setUndoAvailable(ready);
      refreshUndoState();
    });

    setUndoAvailable(isCustomUndoAvailable());
    void connectUndoMonitor();

    return () => {
      unsubscribeInvalidation();
      unsubscribeAvailability();
    };
  }, []);

  useEffect(() => {
    logEnd.current?.scrollIntoView({ block: "end" });
  }, [entries, streaming, pending]);

  function pickProvider(id: string) {
    setProvider(id);
    const p = providers.find((x) => x.id === id);
    if (p) setModel(p.defaultModel);
  }

  function pushOp(event: ToolEvent) {
    setEntries((prev) => {
      const idx = prev.findIndex((e) => e.kind === "op" && e.event.id === event.id);
      if (idx === -1) return [...prev, { kind: "op", event }];
      const next = [...prev];
      next[idx] = { kind: "op", event };
      return next;
    });
  }

  function refreshUndoState() {
    const available = isCustomUndoAvailable();
    setUndoAvailable(available);
    setCanUndo(available && undoDepth() > 0);
    setUndoLabel(available ? describeNextUndo() : null);
  }

  async function connectUndoMonitor() {
    setMonitorStatus("connecting");
    try {
      const supported = await ensureStructuralChangeMonitor();
      if (!supported) {
        setMonitorStatus("unsupported");
        refreshUndoState();
        setEntries((e) => [
          ...e,
          {
            kind: "error",
            text: "ExcelApi 1.9 недоступен: безопасный custom undo отключён. Изменения книги продолжат выполняться, но надстройка не будет создавать собственные точки отмены."
          }
        ]);
        return;
      }
      setMonitorStatus("ready");
      refreshUndoState();
    } catch (err: any) {
      setMonitorStatus("error");
      refreshUndoState();
      setEntries((e) => [
        ...e,
        {
          kind: "error",
          text: `Не удалось включить контроль структурных изменений: ${err?.message ?? String(err)}. Custom undo отключён.`
        }
      ]);
    }
  }

  function send(override?: string) {
    const text = (override ?? draft).trim();
    if (!text || !provider || !model) return;
    const started = withWorkbookLock(lock.current, "task", () => runTask(text));
    if (started === BUSY) {
      setEntries((e) => [...e, { kind: "notice", text: busyReason(lock.current.owner()) }]);
      return;
    }
    void started;
  }

  async function runTask(text: string) {
    setDraft("");
    setStreaming("");
    // Номер беседы — в книгу до первого сообщения: так беседа переживёт
    // сохранение книги, в том числе в OneDrive, где Excel перезапускает панель.
    await ensureDocumentConversationId();

    const spent = { calls: 0, prompt: 0, cached: 0, completion: 0, cost: 0, costKnown: true, estimated: false };
    const controller = new AbortController();
    abort.current = controller;

    try {
      const refreshed = await refreshContext();
      setEntries((e) => [...e, { kind: "user", text }]);
      history.current.push({ role: "user", content: text });
      // Новая сборка могла выйти, пока панель открыта.
      void panelIsStale().then(setStale);
      const budgetMinutes = providers.find((item) => item.id === provider)?.taskBudgetMinutes;
      // Память (8.5): не прочиталась — задача идёт без неё, а не падает.
      const memory = WEB_PANEL ? null : await fetchMemory().catch(() => null);
      const attached = WEB_PANEL ? [] : await listFiles().catch(() => files);
      setFiles(attached);
      await runAgent({
        memoryPrompt: memoryPrompt(memory),
        filesPrompt: filesPrompt(attached),
        webEnabled: webEnabled && Boolean(webServices?.length),
        provider,
        model,
        ...(budgetMinutes ? { taskBudgetMs: budgetMinutes * 60_000 } : {}),
        history: history.current,
        analysisOnly: refreshed?.permissionsReset ? true : analysisOnly,
        ...(refreshed ? { initialContext: refreshed.context } : {}),
        signal: controller.signal,
        hooks: {
          onUsage: (u) => {
            spent.calls += 1;
            spent.prompt += u.promptTokens;
            spent.cached += u.cachedTokens;
            spent.completion += u.completionTokens;
            const priced = stepCost(provider, model, u);
            if (!priced) spent.costKnown = false;
            else {
              spent.cost += priced.cost;
              if (priced.estimated) spent.estimated = true;
            }
          },
          onDelta: (d) => setStreaming((s) => s + d),
          onStepEnd: (t) => {
            setStreaming("");
            if (t.trim()) setEntries((e) => [...e, { kind: "assistant", text: t }]);
          },
          onToolEvent: pushOp,
          confirm: (name, args) =>
            new Promise<boolean>((resolve) => {
              if (controller.signal.aborted) return resolve(false);
              const onAbort = () => {
                setPending(null);
                resolve(false);
              };
              controller.signal.addEventListener("abort", onAbort, { once: true });
              setPending({
                name,
                args,
                resolve: (ok) => {
                  controller.signal.removeEventListener("abort", onAbort);
                  resolve(ok);
                }
              });
            })
        }
      });
    } catch (err: any) {
      if (err?.name === "AbortError") {
        // Прерванные вызовы инструментов уже помечены «отменено». Эта запись
        // закрывает второй случай: остановку во время ответа модели, когда
        // активного вызова нет и в ленте иначе не остаётся ничего.
        setEntries((e) => [
          ...e,
          { kind: "notice", text: "Остановлено вами. Начатое не продолжается; что успело выполниться, показано выше." }
        ]);
      } else {
        setEntries((e) => [...e, { kind: "error", text: err?.message ?? String(err) }]);
      }
    } finally {
      // Расход задачи (01.10.2026: пользователь увидел $1,30 за беседу только в кабинете OpenRouter).
      if (spent.calls) {
        setEntries((e) => {
          const before = conversationSpending(e);
          const note = spendingNote(spent);
          const exact = spent.costKnown ? { cost: spent.cost } : {};
          if (!before.tasks) return [...e, { kind: "notice", text: note, ...exact }];
          const total = conversationSpending([...e, { kind: "notice", text: note, ...exact }]);
          return [...e, { kind: "notice", text: `${note} Всего за беседу: ${conversationTotalText(total)}.`, ...exact }];
        });
      }
      setStreaming("");
      setPending(null);
      abort.current = null;
      refreshUndoState();
    }
  }

  function decide(ok: boolean) {
    pending?.resolve(ok);
    setPending(null);
  }

  function undo() {
    const started = withWorkbookLock(lock.current, "undo", runUndo);
    if (started === BUSY) {
      setEntries((e) => [...e, { kind: "notice", text: busyReason(lock.current.owner()) }]);
      return;
    }
    void started;
  }

  async function runUndo() {
    try {
      const msg = await undoLast();
      setEntries((e) => [...e, { kind: "assistant", text: msg }]);
      // Отмена идёт мимо агента. Проверка 18 сентября 2026 года: после семи
      // отмен агент увидел, что курсив «пропал», и сочинил объяснение — будто
      // его операция не сохранилась. Поэтому каждая отмена записывается
      // в его историю. Роль пользовательская: системное сообщение посреди
      // беседы принимают не все провайдеры, — а пометка говорит, кто автор.
      if (msg.startsWith("Отменено")) {
        history.current.push({
          role: "user",
          content:
            `[Сообщение панели, не от пользователя] Пользователь нажал «Отменить» в панели. ${msg} ` +
            "Состояние этой области вернулось к тому, что было до операции; прежние чтения и отчёты о ней устарели."
        });
      }
    } catch (err: any) {
      setEntries((e) => [...e, { kind: "error", text: `Отмена не удалась: ${err?.message ?? err}` }]);
    } finally {
      refreshUndoState();
    }
  }

  function reset() {
    if (workbookBinding.current) deleteConversation(localStorage, workbookBinding.current.key);
    history.current = [];
    setEntries([]);
    setStreaming("");
  }

  const current = providers.find((p) => p.id === provider);
  const spendingTotal = conversationSpending(entries as Array<{ kind: string; text?: string }>);
  const groups = groupEntries(entries);
  // Знак am.AI — у первого ответа после просьбы: так видно, где начинается ход панели.
  const awaitingLead = groups.length > 0 && groups[groups.length - 1].kind === "single" &&
    (groups[groups.length - 1] as { entry: Entry }).entry.kind === "user";

  return (
    <div className="pane">
      <div className="head">
        <div className="brand">
          <Logo size={26} />
          <Wordmark />
          <span className="brand-version" title={`Сборка панели ${PANEL_BUILD} (UTC)`}>{PANEL_VERSION}</span>
        </div>
        <span className="spacer" />
        <button className="ghost" onClick={() => setShowConversations((open) => !open)} disabled={busy} aria-expanded={showConversations}>
          Беседы
        </button>
        {!WEB_PANEL && (
          <button className="ghost" onClick={() => setShowMemory((open) => !open)} disabled={busy} aria-expanded={showMemory}>
            Память
          </button>
        )}
        <button className="ghost" onClick={() => setShowKeys((open) => !open)} disabled={busy} aria-expanded={showKeys}>
          Ключи
        </button>
      </div>

      <div className="toolbar">
        <select value={provider} onChange={(e) => pickProvider(e.target.value)} disabled={busy} aria-label="Провайдер">
          {providers.map((p) => (
            <option key={p.id} value={p.id}>
              {p.label}
            </option>
          ))}
        </select>
        <select value={model} onChange={(e) => setModel(e.target.value)} disabled={busy} aria-label="Модель">
          {current?.models.map((m) => (
            <option key={m} value={m}>
              {m}
            </option>
          ))}
        </select>
        <span className="spacer" />
        <button
          className="ghost icon-text undo"
          onClick={undo}
          disabled={!canUndo || busy}
          title={
            !undoAvailable
              ? "Custom undo недоступен: монитор структуры книги не активен"
              : undoLabel
                ? `Отменить: ${undoLabel}`
                : "Вернуть последнюю правку"
          }
        >
          <Icon.undo />
          {/* Неактивная отмена — только значок: место в строке нужнее модели. */}
          {(canUndo || !undoAvailable) && <span>{!undoAvailable ? "Undo недоступен" : "Отменить"}</span>}
        </button>
        {monitorStatus === "error" && (
          <button className="ghost" onClick={() => void connectUndoMonitor()} disabled={busy}>
            Повторить защиту undo
          </button>
        )}
        <button className="ghost" onClick={reset} disabled={busy} title="Начать беседу заново">
          Очистить
        </button>
      </div>

      <div className="context-bar">
        <span className="context-label" title={`${contextLabel}
${persistenceNote}`}><Icon.sheet />{contextLabel}</span>
        <label className={`toggle${analysisOnly ? " on" : ""}`} title="Пока галочка стоит, am.AI только читает книгу и ничего не меняет.">
          <input type="checkbox" checked={analysisOnly} onChange={(event) => setAnalysisOnly(event.target.checked)} disabled={busy} />
          Только анализ
        </label>
        {!WEB_PANEL && <label
          className="toggle"
          title={webServices?.length
            ? `Поиск через ${webServices.join(", ")}. В сервис уходит только текст запроса; страницы читает этот компьютер.`
            : "Добавьте ключ Tavily или Serper в «Ключах», чтобы включить поиск в интернете."}
        >
          <input
            type="checkbox"
            checked={webEnabled && Boolean(webServices?.length)}
            disabled={busy || !webServices?.length}
            onChange={(event) => {
              setWebEnabled(event.target.checked);
              try { localStorage.setItem("amai.web", event.target.checked ? "1" : "0"); } catch { /* хранилище недоступно */ }
            }}
          />
          Интернет
        </label>}
      </div>
      {/* Обычное «беседа хранится локально» — в подсказке строки книги; видна только важная заметка. */}
      {persistenceNote && !persistenceNote.startsWith("Беседа хранится локально") && <div className="persistence-note">{persistenceNote}</div>}
      {update && (
        <div className="undo-note">
          Вышла версия {update.latest} (у вас {update.current}).{" "}
          {update.url ? <a href={update.url} target="_blank" rel="noreferrer">Скачать с GitHub</a> : "Её можно скачать на GitHub."}
        </div>
      )}
      {stale && (
        <div className="warn-note">
          Панель устарела: на сервере уже новая сборка. Закройте панель и откройте заново — иначе агент работает
          со старым набором инструментов.
        </div>
      )}

      {(apiStatus !== "ready" || apiError) && (
        <div className="api-status" role="status">
          {apiStatus === "checking" ? "Проверка локального API…" : apiError}
          {apiStatus !== "checking" && <button className="ghost" onClick={() => void loadProviders()}>Повторить подключение</button>}
        </div>
      )}

      {showKeys && !busy && <KeysPanel onChanged={() => { void loadProviders(); void loadWeb(); }} onClose={() => setShowKeys(false)} />}
      {showConversations && !busy && (
        <ConversationsPanel
          currentKey={workbookBinding.current?.key ?? null}
          onClose={() => setShowConversations(false)}
          onDeletedCurrent={() => {
            history.current = [];
            setEntries([]);
            setStreaming("");
          }}
        />
      )}
      {showMemory && !busy && (
        <MemoryPanel
          onClose={() => setShowMemory(false)}
          onRun={(scenario: Scenario) => {
            setShowMemory(false);
            send(`Выполни сценарий «${scenario.name}» по шагам, каждое изменение — через карточку:\n${scenario.steps.map((step, index) => `${index + 1}. ${step}`).join("\n")}`);
          }}
        />
      )}

      <div className="log">
        {entries.length === 0 && (
          <div className="empty">
            <div className="hero">
              <Logo size={88} />
            </div>
            <h2>Чем помочь с <span className="gradient-text">книгой</span>?</h2>
            <p>Напишите обычными словами — am.AI сам прочитает нужные ячейки. Перед изменениями он покажет план и спросит разрешения.</p>
            <div className="examples">
              {["Посчитай итоги по месяцам", "Найди ошибки в формулах", "Сделай диаграмму по таблице", "Убери дубли в списке"].map((example) => (
                <button key={example} className="example" onClick={() => setDraft(example)} disabled={busy}><i />{example}</button>
              ))}
            </div>
          </div>
        )}

        {groups.map((group, groupIndex) => {
          const previous = groups[groupIndex - 1];
          const lead = group.kind === "ops" || group.entry.kind !== "user"
            ? !previous || (previous.kind === "single" && previous.entry.kind === "user")
            : false;
          const withAvatar = (node: JSX.Element) => lead
            ? <div key={`lead-${group.start}`} className="lead"><Logo size={24} className="avatar" />{node}</div>
            : node;
          if (group.kind === "ops") {
            const ops = group.items;
            const open = ops.some(({ event }) => ["running", "error", "uncertain", "rejected"].includes(event.status));
            const failed = ops.some(({ event }) => event.status === "error" || event.status === "uncertain");
            const running = ops.some(({ event }) => event.status === "running");
            const names = [...new Set(ops.map(({ event }) => toolLabel(event.name)))];
            return withAvatar(
              <details key={`ops-${group.start}`} className={`steps${failed ? " has-error" : ""}`} open={open || undefined}>
                <summary>
                  <span className="steps-icon">{failed ? <Icon.alert /> : running ? <span className="spinner" /> : <Icon.check />}</span>
                  <span className="steps-title">{actionsWord(ops.length)}</span>
                  <span className="steps-names">{names.slice(0, 3).join(", ")}{names.length > 3 ? "…" : ""}</span>
                  <span className="steps-chevron"><Icon.chevron /></span>
                </summary>
                {ops.map(({ event }, n) => {
                  const addr = addressOf(event.args);
                  return (
                    <div key={`${event.id}-${n}`} className={`op ${event.status}`} title={event.name}>
                      <span className="name">{toolLabel(event.name)}</span>
                      {addr && (
                        <>
                          {" "}
                          <span className="addr">{addr}</span>
                        </>
                      )}
                      {event.status === "rejected" && " — отклонено"}
                      {event.status === "cancelled" && " — отменено"}
                      {event.status === "uncertain" && " — проверьте книгу перед новой правкой"}
                      {/* Причину остановки показываем: без неё ни человек, ни разбор
                          не видят, что именно вернул Excel. */}
                      {event.status === "uncertain" && event.result && <div className="undo-note">{String(event.result)}</div>}
                      {event.status === "done" && event.undoable === false && " — без автоматической отмены"}
                      {event.undoNote && <div className="undo-note">{event.undoNote}</div>}
                      {event.status === "error" && ` — ${event.result}`}
                    </div>
                  );
                })}
              </details>
            );
          }
          const e = group.entry;
          if (e.kind === "assistant") {
            return withAvatar(
              <div key={group.start} className="msg assistant">
                <Markdown text={e.text} />
              </div>
            );
          }
          if (e.kind === "notice" && e.text.startsWith("Расход задачи:")) {
            return withAvatar(
              <div key={group.start} className="msg notice spend" title={e.text}>
                <Icon.coin /><span>{e.text.replace(/^Расход задачи: /, "").replace(/ Всего за беседу:.*$/, "")}</span>
              </div>
            );
          }
          return e.kind === "user" ? (
            <div key={group.start} className={`msg ${e.kind}`}>
              {e.text}
            </div>
          ) : withAvatar(
            <div key={group.start} className={`msg ${e.kind}`}>
              {e.text}
            </div>
          );
        })}

        {pending && (
          <div className="confirm">
            {pending.name === "__read_sheets" ? (() => {
              // Границы чтения (8.0.6): лист вне просьбы читается только с разрешения.
              const request = pending.args as { sheets: string[]; tool: string };
              const many = request.sheets.length > 1;
              return (
                <>
                  <p>
                    Агент хочет прочитать {many ? "листы" : "лист"} <strong>{request.sheets.map((sheet) => `«${sheet}»`).join(", ")}</strong>,
                    {" "}{many ? "которые" : "который"} вы не называли в просьбе.
                  </p>
                  <div className="preview">
                    <p>
                      Данные с {many ? "этих листов" : "этого листа"} уйдут провайдеру модели ({provider || "выбранному"}). Книга не меняется.
                      Разрешение действует до конца этой задачи.
                    </p>
                    <p className="undo-note">
                      Если вы не просили об этом, лучше не разрешать: так срабатывает и текст в ячейке, уговаривающий агента прочитать лишнее.
                    </p>
                  </div>
                </>
              );
            })() : (
              <p>
                Разрешить <strong>{pending.name}</strong>
                {addressOf(pending.args) ? ` в ${addressOf(pending.args)}` : ""}? {pending.name === "create_workbook_backup"
                  ? "Книга не изменится, но на диске появится файл."
                  : pending.name === "remember_preference" || pending.name === "save_scenario"
                    ? "Книга не изменится: запись добавится в память панели на этом компьютере."
                    : "Операция изменит книгу."}
              </p>
            )}
            {pending.name === "set_range_values" && (() => {
              const plan = pending.args as any;
              return (
                <div className="preview">
                  <p>
                    {plan.cellCount} ячеек; режим записи: <strong>{plan.isFormula ? "формулы" : "литеральные значения"}</strong>;
                    {" "}заменяемых формул: {plan.replacedFormulaCount}; проверка: {plan.errorScanAddress}
                  </p>
                  {plan.tableWarning && <p className="warn-note">{plan.tableWarning}</p>}
                  {plan.mergeWarning && <p className="warn-note">{plan.mergeWarning}</p>}
                  {plan.functionCheck && <p className="undo-note">{plan.functionCheck.note}</p>}
                  <div><strong>Было</strong><pre>{JSON.stringify(plan.before, null, 2)}</pre></div>
                  <div><strong>Станет</strong><pre>{JSON.stringify(plan.after, null, 2)}</pre></div>
                  <p className="undo-note">{plan.undoAvailable ? "После проверки будет доступна безопасная отмена, если диапазон не изменится." : "Автоматическая отмена сейчас недоступна."}</p>
                </div>
              );
            })()}
            {pending.name === "set_ranges_values" && (() => {
              const plan = pending.args as any;
              const items: any[] = plan.items ?? [];
              return (
                <div className="preview">
                  <p>
                    {items.length} операций, {plan.cellCount} ячеек всего. Выполняются по порядку.
                  </p>
                  <p className="warn-note">
                    Это не единая транзакция: после сбоя оставшиеся операции не начнутся, а уже выполненные
                    не откатятся автоматически.
                  </p>
                  {items.map((item, index) => (
                    <div key={item.id ?? index}>
                      <strong>
                        {index + 1}. {item.target?.sheetName}!{item.resolvedAddress}
                      </strong>
                      {" — "}
                      {item.cellCount} ячеек, {item.isFormula ? "формулы" : "литеральные значения"}
                      {item.replacedFormulaCount > 0 && `, заменяемых формул: ${item.replacedFormulaCount}`}
                      {item.mergeWarning && <p className="warn-note">{item.mergeWarning}</p>}
                      {item.functionCheck && <p className="undo-note">{item.functionCheck.note}</p>}
                      <div>
                        <pre>{JSON.stringify(item.before, null, 2)}</pre>
                        <pre>{JSON.stringify(item.after, null, 2)}</pre>
                      </div>
                    </div>
                  ))}
                </div>
              );
            })()}
            {pending.name === "fill_range" && (() => {
              const plan = pending.args as any;
              return (
                <div className="preview">
                  {plan.template ? (
                    <>
                      <p>
                        {plan.cellCount} ячеек ({plan.rows} × {plan.columns}); первая строка {plan.templateRowAddress} по шаблону,
                        каждый столбец протягивается вниз своей формулой, ссылки подстраиваются по строкам.
                      </p>
                      <pre>{plan.template.map((item: unknown) => String(item)).join("  |  ")}</pre>
                    </>
                  ) : (
                    <p>
                      {plan.cellCount} ячеек ({plan.rows} × {plan.columns}); {plan.isFormula ? "формула" : "значение"}{" "}
                      <strong>{String(plan.value)}</strong> из первой ячейки {plan.anchorAddress}
                      {plan.isFormula ? "; Excel протянет её по области, подстраивая ссылки" : " во все ячейки"}.
                    </p>
                  )}
                  {plan.occupiedCells > 0 && (
                    <p className="warn-note">Непустых ячеек в области: {plan.occupiedCells} — их содержимое будет заменено.</p>
                  )}
                  {plan.tableWarning && <p className="warn-note">{plan.tableWarning}</p>}
                  {plan.mergeWarning && <p className="warn-note">{plan.mergeWarning}</p>}
                  {plan.functionCheck && <p className="undo-note">{plan.functionCheck.note}</p>}
                  <p className="undo-note">
                    {plan.undoAvailable ? "После проверки будет доступна отмена, если данные не изменятся." : plan.undoNote}
                  </p>
                </div>
              );
            })()}
            {(pending.name === "trim_text" || pending.name === "convert_values" || pending.name === "change_case") && (() => {
              const plan = pending.args as any;
              const skipped = Object.entries(plan.skipped ?? {}) as [string, { count: number; examples: string[] }][];
              return (
                <div className="preview">
                  <p>{plan.description}</p>
                  <p>
                    Изменится ячеек: <strong>{plan.changes.length}</strong> в {plan.target?.sheetName}!{plan.resolvedAddress}.
                  </p>
                  <div><strong>Было → станет</strong><pre>{plan.sample.join("\n")}</pre></div>
                  {skipped.length > 0 && (
                    <div className="warn-note">
                      Не изменятся:
                      <ul>{skipped.map(([reason, item]) => <li key={reason}>{reason} — {item.count} ({item.examples.join("; ")})</li>)}</ul>
                    </div>
                  )}
                  <p className="undo-note">
                    {plan.undoAvailable ? "После проверки будет доступна отмена, если данные не изменятся." : plan.undoNote}
                  </p>
                </div>
              );
            })()}
            {pending.name === "remove_duplicates" && (() => {
              const plan = pending.args as any;
              return (
                <div className="preview">
                  <p>
                    {plan.target?.sheetName}!{plan.resolvedAddress}: ключ — {plan.keyNames.join(", ")}. Дубликатов: <strong>{plan.removed.length}</strong> из {plan.dataRows} строк;
                    {plan.dest
                      ? ` уникальные строки с шапкой будут скопированы значениями на лист «${plan.dest.sheetName}» (${plan.dest.address}), источник не изменится.`
                      : " остаются первые вхождения, строки ниже поднимутся внутри области."}
                  </p>
                  <div><strong>Удаляемые строки</strong><pre>{plan.sampleRemoved.join("\n")}</pre></div>
                  {plan.risks.length > 0 && (
                    <div className="warn-note">
                      Формулы, которые станут смотреть на другие строки (ссылки не подстроятся):
                      <ul>{plan.risks.map((risk: any) => <li key={risk.sheet + risk.cell}>{risk.sheet}!{risk.cell}: {risk.formula}</li>)}</ul>
                      {plan.riskOverflow > 0 && <p>…и ещё {plan.riskOverflow}.</p>}
                    </div>
                  )}
                  {plan.unscannedSheets?.length > 0 && (
                    <p className="warn-note">Листы {plan.unscannedSheets.join(", ")} слишком велики для обхода формул: про них ничего не проверено.</p>
                  )}
                  <p className={plan.dest ? "undo-note" : "warn-note"}>{plan.undoNote}</p>
                  {!plan.dest && (plan.backup
                    ? <p className="undo-note">Последняя резервная копия: {plan.backup.name}.</p>
                    : <p className="warn-note">Резервной копии в этом сеансе не создавалось.</p>)}
                </div>
              );
            })()}
            {pending.name === "remember_preference" && (() => {
              const plan = pending.args as any;
              return (
                <div className="preview">
                  <p>Запомнить на этом компьютере предпочтение ({CATEGORY_TEXT[plan.category as keyof typeof CATEGORY_TEXT] ?? plan.category}):</p>
                  <p><strong>«{plan.text}»</strong></p>
                  <p className="undo-note">Со следующей задачи агент получит его в начале. Изменить или удалить — окно «Память». Книга не меняется.</p>
                </div>
              );
            })()}
            {pending.name === "save_scenario" && (() => {
              const plan = pending.args as any;
              return (
                <div className="preview">
                  <p>Сохранить сценарий <strong>«{plan.name}»</strong> на этом компьютере:</p>
                  <ol>{(plan.steps ?? []).map((step: string, index: number) => <li key={index}>{step}</li>)}</ol>
                  <p className="undo-note">Запуск — кнопкой в окне «Память» или просьбой «выполни сценарий …». Каждое изменение при запуске — через свою карточку. Книга сейчас не меняется.</p>
                </div>
              );
            })()}
            {pending.name === "import_file_layout" && (() => {
              const plan = pending.args as any;
              return (
                <div className="preview">
                  <p>
                    Копия «<strong>{plan.fileName}</strong>» как в файле — на новый лист «<strong>{plan.sheetName}</strong>»:
                    {" "}{plan.rows} строк × {plan.columns} столбцов, страниц {plan.layout?.pages ?? 1}, объединений {plan.merges}, сторон рамок {plan.borderedEdges}.
                  </p>
                  {plan.preview?.length > 0 && <pre>{plan.preview.join("\n")}</pre>}
                  {(plan.layout?.warnings ?? []).map((text: string) => <p key={text} className="warn-note">{text}</p>)}
                  <p className="undo-note">Картинки, печати и цвета не переносятся. {plan.undoAvailable ? "Отмена удалит этот лист, если на нём ничего не меняли." : plan.undoNote}</p>
                </div>
              );
            })()}
            {pending.name === "import_file_table" && (() => {
              const plan = pending.args as any;
              return (
                <div className="preview">
                  <p>
                    Перенести «<strong>{plan.fileName}</strong>» ({plan.tableName}, строки {plan.sourceRows}) на{" "}
                    <strong>{plan.newSheet ? `новый лист «${plan.destSheet}»` : plan.destSheet}</strong>, {plan.destArea}: {plan.rows} × {plan.columns}. Место пустое — ничего не затрётся.
                  </p>
                  <div><strong>Первые строки</strong><pre>{plan.preview.join("\n")}</pre></div>
                  {plan.numbersFromText > 0 && <p>Однозначных чисел из текста файла: {plan.numbersFromText}.</p>}
                  {plan.keptAsText.count > 0 && (
                    <p className="warn-note">Останутся текстом ({plan.keptAsText.count}) — разделители или порядок даты неоднозначны: {plan.keptAsText.examples.map((text: string) => `«${text}»`).join(", ")}. Их переведёт в числа отдельная операция.</p>
                  )}
                  <p className="undo-note">Текст из файла записывается текстом, формулой не станет. После записи каждая ячейка сверяется с файлом. {plan.undoAvailable ? "Отмена будет доступна." : plan.undoNote}</p>
                </div>
              );
            })()}
            {pending.name === "set_page_layout" && (() => {
              const plan = pending.args as any;
              return (
                <div className="preview">
                  <p>Параметры печати листа {plan.target?.sheetName}:</p>
                  <pre>{plan.preview.join("\n")}</pre>
                  <p>Данные и оформление ячеек не меняются.</p>
                  <p className="undo-note">{plan.undoAvailable ? "После проверки будет доступна отмена." : plan.undoNote}</p>
                </div>
              );
            })()}
            {pending.name === "copy_sheet" && (() => {
              const plan = pending.args as any;
              return (
                <div className="preview">
                  <p>
                    Копия листа «<strong>{plan.sourceName}</strong>» → «<strong>{plan.newName ?? `${plan.sourceName} (2)`}</strong>», {plan.position === "end" ? "последним листом" : "сразу за исходным"}.
                    {plan.facts.used ? ` Данные ${plan.facts.used}` : " Лист пуст"}{plan.facts.formulas ? `, формул ${plan.facts.formulas}` : ""}{plan.facts.charts ? `, диаграмм ${plan.facts.charts}` : ""}{plan.facts.tables ? `, таблиц ${plan.facts.tables}` : ""}.
                  </p>
                  {plan.warnings.map((text: string) => <p key={text} className="warn-note">{text}</p>)}
                  <p className="undo-note">{plan.undoAvailable ? "После проверки будет доступна отмена, если копию не менять." : plan.undoNote}</p>
                </div>
              );
            })()}
            {pending.name === "rename_sheet" && (() => {
              const plan = pending.args as any;
              const list = (items: any[]) => <ul>{items.map((item) => <li key={item.sheet + item.cell}>{item.sheet}!{item.cell}: {item.formula}</li>)}</ul>;
              return (
                <div className="preview">
                  <p>Лист «{plan.oldName}» → «<strong>{plan.newName}</strong>». Ссылки на него Excel перепишет сам: формул с такими ссылками — {plan.referencing.length}.</p>
                  {plan.literal.length > 0 && (
                    <div className="warn-note">Имя листа внутри текста — Excel его не перепишет, формулы покажут #ССЫЛКА!:{list(plan.literal)}</div>
                  )}
                  {plan.textMentions.length > 0 && <div className="undo-note">Упоминания текстом (не изменятся):{list(plan.textMentions)}</div>}
                  {plan.unscannedSheets.length > 0 && <p className="warn-note">Листы {plan.unscannedSheets.join(", ")} слишком велики для обхода формул: про них ничего не проверено.</p>}
                  <p className="undo-note">{plan.undoNote}</p>
                </div>
              );
            })()}
            {pending.name === "delete_sheet" && (() => {
              const plan = pending.args as any;
              const broken = [...plan.referencing, ...plan.literal, ...plan.viaNames];
              return (
                <div className="preview">
                  <p>
                    Удалить лист «<strong>{plan.sheetName}</strong>»: {plan.usedAddress ? `данные ${plan.usedAddress}${plan.filledCells !== null ? `, непустых ячеек ${plan.filledCells}` : ""}` : "лист пуст"}
                    {plan.charts > 0 ? `, диаграмм ${plan.charts}` : ""}{plan.pivots > 0 ? `, сводных ${plan.pivots}` : ""}{plan.tables.length > 0 ? `, таблицы ${plan.tables.join(", ")}` : ""}.
                  </p>
                  {broken.length > 0 && (
                    <div className="warn-note">
                      Станут #ССЫЛКА!:
                      <ul>{broken.map((item: any) => <li key={item.sheet + item.cell}>{item.sheet}!{item.cell}: {item.formula}</li>)}</ul>
                      {plan.overflow > 0 && <p>…и ещё {plan.overflow}.</p>}
                    </div>
                  )}
                  {plan.brokenNames.length > 0 && <p className="warn-note">Сломаются именованные диапазоны: {plan.brokenNames.join(", ")}.</p>}
                  {plan.unscannedSheets.length > 0 && <p className="warn-note">Листы {plan.unscannedSheets.join(", ")} слишком велики для обхода формул: про них ничего не проверено.</p>}
                  <p className="warn-note">{plan.undoNote}</p>
                  {plan.backup
                    ? <p className="undo-note">Последняя резервная копия: {plan.backup.name}.</p>
                    : <p className="warn-note">Резервной копии в этом сеансе не создавалось.</p>}
                </div>
              );
            })()}
            {(pending.name === "insert_columns" || pending.name === "delete_columns") && (() => {
              const plan = pending.args as any;
              const deleting = pending.name === "delete_columns";
              const rows = (m: unknown[][]) => (m ?? []).map((r) => r.map((c) => (c === "" || c === null ? "∅" : String(c))).join(" · ")).join("\n");
              return (
                <div className="preview">
                  <p>
                    {deleting ? "Удалить" : "Вставить"} столбцы <strong>{plan.target?.sheetName}!{plan.columnsAddress}</strong> ({plan.count});
                    {deleting ? ` непустых ячеек в них: ${plan.filledCells}.` : " существующие сдвинутся вправо."}
                  </p>
                  {deleting && plan.preview.length > 0 && <div><strong>Что удалится</strong><pre>{rows(plan.preview)}{plan.previewTruncated ? "\n…" : ""}</pre></div>}
                  {plan.formulaRisks.length > 0 && (
                    <div className="warn-note">
                      {deleting ? "Формулы, которые сломаются или молча сузятся:" : "Формулы, которые не охватят новые столбцы:"}
                      <ul>{plan.formulaRisks.map((risk: any) => <li key={risk.sheet + risk.address + risk.reference}>{risk.sheet}!{risk.address}: {risk.formula} ({risk.kind === "broken" ? "станет #ССЫЛКА!" : risk.kind === "shrunk" ? "сузится" : "не охватит"})</li>)}</ul>
                      {plan.riskOverflow > 0 && <p>…и ещё {plan.riskOverflow}.</p>}
                    </div>
                  )}
                  {plan.tableFormulaSheets.length > 0 && <p className="warn-note">Формулы со ссылками на таблицы (листы {plan.tableFormulaSheets.join(", ")}) не разбирались.</p>}
                  {plan.unscannedSheets.length > 0 && <p className="warn-note">Листы {plan.unscannedSheets.join(", ")} слишком велики для обхода формул: про них ничего не проверено.</p>}
                  {plan.mergeWarning && <p className="warn-note">{plan.mergeWarning}</p>}
                  <p className="warn-note">{plan.undoNote}</p>
                </div>
              );
            })()}
            {pending.name === "group_rows_columns" && (() => {
              const plan = pending.args as any;
              const hiddenCount = plan.hiddenBefore.filter(Boolean).length;
              return (
                <div className="preview">
                  <p>
                    Сгруппировать {plan.band.axis === "rows" ? "строки" : "столбцы"} <strong>{plan.target?.sheetName}!{plan.band.address}</strong>
                    {plan.collapse ? " и оставить группу свёрнутой" : ""}. Данные не меняются.
                  </p>
                  <p className="undo-note">
                    Сейчас скрыто {hiddenCount} из {plan.hiddenBefore.length}. Панель проверит группу, на миг свернув её, и вернёт видимость
                    {plan.collapse ? " свёрнутой" : " прежней"}. Уровни прежних групп Excel не сообщает: если группа здесь уже была, новая добавится к ней уровнем.
                  </p>
                  <p className="undo-note">{plan.undoNote}</p>
                </div>
              );
            })()}
            {pending.name === "build_lbo_model" && (() => {
              const plan = pending.args as any;
              const r = plan.request ?? {};
              const x = r.assumptions ?? {};
              return (
                <div className="preview">
                  <p>
                    Новый лист <strong>{r.sheet}</strong>: модель LBO, вход в {r.entryYear}, владение {r.years} лет, {r.currency}, {r.units}.
                    Цена {x.entryMultiple}× EBITDA, долг {x.debtMultiple}× EBITDA, выход {x.exitMultiple}× EBITDA. Источник допущений: {r.source}.
                  </p>
                  <p>Расчёт панели: вложение ≈ {Math.round(plan.layout?.equity0 ?? 0).toLocaleString("ru-RU")}, MOIC {(plan.layout?.moic ?? 0).toFixed(2)}×, IRR {((plan.layout?.irr ?? 0) * 100).toFixed(1)} %. После записи каждое значение и три контрольных равенства сверяются.</p>
                  {plan.layout?.lowCoverage?.length > 0 && <p className="warn-note">EBITDA покрывает проценты меньше чем вдвое в годы: {plan.layout.lowCoverage.join(", ")}.</p>}
                  <p className="undo-note">{plan.simplifications}</p>
                  <p className="undo-note">{plan.undoNote}</p>
                </div>
              );
            })()}
            {pending.name === "build_dcf_model" && (() => {
              const plan = pending.args as any;
              const r = plan.request ?? {};
              const pct = (x: number) => `${(x * 100).toFixed(1)} %`;
              return (
                <div className="preview">
                  <p>
                    Новый лист <strong>{r.sheet}</strong>: оценка DCF на {r.years} г. ({r.firstYear}–{r.firstYear + r.years - 1}), {r.currency}, {r.units}.
                    WACC {pct(r.assumptions?.wacc ?? 0)}, рост после прогноза {pct(r.assumptions?.terminalGrowth ?? 0)}. Источник допущений: {r.source}.
                  </p>
                  <p>Расчёт панели: стоимость бизнеса ≈ {Math.round(plan.layout?.ev ?? 0).toLocaleString("ru-RU")}, из неё остаточная стоимость — {pct(plan.layout?.tvShare ?? 0)}. После записи каждое значение и таблица чувствительности сверяются.</p>
                  {plan.layout?.tvShare > 0.75 && <p className="warn-note">Больше трёх четвертей оценки — остаточная стоимость: результат держится на росте после прогноза и WACC.</p>}
                  <p className="undo-note">{plan.simplifications}</p>
                  <p className="undo-note">{plan.undoNote}</p>
                </div>
              );
            })()}
            {pending.name === "build_three_statement_model" && (() => {
              const plan = pending.args as any;
              const r = plan.request ?? {};
              return (
                <div className="preview">
                  <p>
                    Новый лист <strong>{r.sheet}</strong>: трёхотчётная модель на {r.years} г. ({r.firstYear}–{r.firstYear + r.years - 1}), {r.currency}, {r.units}.
                    Источник допущений: {r.source}.
                  </p>
                  <p>Отчёты — формулы от блока допущений; после записи каждое значение и баланс каждого года сверяются с расчётом панели.</p>
                  {plan.layout?.negativeCash?.length > 0 && <p className="warn-note">Деньги на конец отрицательны в годы: {plan.layout.negativeCash.join(", ")} — модели не хватит финансирования.</p>}
                  <p className="undo-note">{plan.simplifications}</p>
                  <p className="undo-note">{plan.undoNote}</p>
                </div>
              );
            })()}
            {pending.name === "add_multiples" && (() => {
              const plan = pending.args as any;
              return (
                <div className="preview">
                  <p>
                    Блок формул <strong>{plan.target?.sheetName}!{plan.destAddress}</strong> по таблице {plan.sourceAddress} ({plan.items} компаний):
                    EV и {plan.multiples.join(", ")} по каждой компании; медиана, среднее, 1-й и 3-й квартили, минимум и максимум по группе.
                  </p>
                  {plan.layout?.undefinedCells?.length > 0 && (
                    <p className="warn-note">Мультипликатор не имеет смысла (убыток, отрицательная EBITDA или нет числа), ячейки останутся пустыми и не войдут в медиану: {plan.layout.undefinedCells.join(", ")}.</p>
                  )}
                  {plan.skipped?.length > 0 && <p className="warn-note">{plan.skipped.join("; ")}.</p>}
                  <p className="undo-note">После записи каждое значение сверяется с расчётом панели. {plan.undoNote}</p>
                </div>
              );
            })()}
            {pending.name === "add_comparison" && (() => {
              const plan = pending.args as any;
              return (
                <div className="preview">
                  <p>
                    Блок формул <strong>{plan.target?.sheetName}!{plan.destAddress}</strong> по таблице {plan.sourceAddress} ({plan.items} объектов × {plan.periods} показателей):
                    среднее, медиана, минимум и максимум каждого показателя, отклонение каждого объекта от медианы и место (1 — наибольшее значение).
                  </p>
                  {plan.layout?.undefinedCells?.length > 0 && (
                    <p className="warn-note">Отклонение не определено (нет числа или медиана — ноль), ячейки останутся пустыми: {plan.layout.undefinedCells.join(", ")}.</p>
                  )}
                  <p className="undo-note">После записи каждое значение сверяется с расчётом панели. {plan.undoNote}</p>
                </div>
              );
            })()}
            {pending.name === "add_share_growth" && (() => {
              const plan = pending.args as any;
              return (
                <div className="preview">
                  <p>
                    Блок формул <strong>{plan.target?.sheetName}!{plan.destAddress}</strong> по таблице {plan.sourceAddress} ({plan.items} статей × {plan.periods} периодов):
                    доля каждой статьи в итоге периода, контроль «сумма долей = 100 %», рост к прошлому периоду.
                  </p>
                  {plan.layout?.undefinedCells?.length > 0 && (
                    <p className="warn-note">Рост не определён (в прошлом периоде ноль или пусто), ячейки останутся пустыми: {plan.layout.undefinedCells.join(", ")}.</p>
                  )}
                  <p className="undo-note">После записи каждое значение сверяется с расчётом панели. {plan.undoNote}</p>
                </div>
              );
            })()}
            {pending.name === "apply_color_convention" && (() => {
              const plan = pending.args as any;
              const swatch = (role: string, text: string) => plan.counts?.[role] > 0 && (
                <li key={role}><span style={{ color: plan.palette[role], fontWeight: 600 }}>{plan.palette[role]}</span> — {text}: {plan.counts[role]}</li>
              );
              return (
                <div className="preview">
                  <p>Цвет текста по роли ячеек на <strong>{plan.target?.sheetName}!{plan.resolvedAddress}</strong>:</p>
                  <ul>
                    {swatch("input", "входы (введённые числа)")}
                    {swatch("formula", "формулы на этом листе")}
                    {swatch("link", "ссылки на другой лист или книгу")}
                    {swatch("check", "контрольные ячейки")}
                  </ul>
                  <p className="undo-note">{plan.paletteNote} Подписи и пустые ячейки ({plan.skipped}) не трогаются.</p>
                  {plan.yearLabels?.length > 0 && <p className="undo-note">Годы в шапке считаются подписями: {plan.yearLabels.join(", ")}.</p>}
                  {plan.overwritten?.length > 0 && <p className="warn-note">Прежний цвет текста будет заменён: {plan.overwritten.join(", ")}.</p>}
                  {plan.conditionalNote && <p className="warn-note">{plan.conditionalNote}</p>}
                  <p className="undo-note">{plan.undoNote}</p>
                </div>
              );
            })()}
            {pending.name === "move_conditional_format" && (() => {
              const plan = pending.args as any;
              return (
                <div className="preview">
                  <p>
                    Правило <strong>{plan.ruleText}</strong> на {plan.target?.sheetName}!{plan.resolvedAddress} встанет {plan.to === "first" ? "выше всех правил области" : "ниже всех правил области"}.
                  </p>
                  <p>Сейчас по приоритету: {(plan.orderBefore ?? []).map((text: string, index: number) => `${index + 1}. ${text}`).join(" · ")}</p>
                  <p className="undo-note">{plan.undoNote}</p>
                </div>
              );
            })()}
            {pending.name === "convert_table_to_range" && (() => {
              const plan = pending.args as any;
              return (
                <div className="preview">
                  <p>
                    Таблица <strong>{plan.tableName}</strong> ({plan.target?.sheetName}!{plan.resolvedAddress}, {plan.rows} × {plan.columns}) станет обычным диапазоном. Данные остаются на месте.
                  </p>
                  <p className="warn-note">{plan.styleNote}</p>
                  {(plan.warnings ?? []).map((warning: string) => <p key={warning} className="warn-note">{warning}</p>)}
                  <p className="undo-note">{plan.undoNote}</p>
                </div>
              );
            })()}
            {pending.name === "set_data_validation" && (() => {
              const plan = pending.args as any;
              return (
                <div className="preview">
                  <p>
                    Правило ввода на <strong>{plan.target?.sheetName}!{plan.resolvedAddress}</strong>: {plan.ruleText}.
                    {plan.previousType === "None" ? " Правила здесь нет." : plan.previousType === "Inconsistent" ? " Сейчас на области разные правила — все заменятся." : ` Заменит прежнее (${plan.previousType}).`}
                  </p>
                  {plan.predictedInvalidCount > 0 && (
                    <p className="warn-note">
                      Уже введённые значения, которые правилу не соответствуют (оценка панели): {plan.predictedInvalidCount} — {plan.predictedInvalid.join(", ")}{plan.predictedInvalidCount > plan.predictedInvalid.length ? "…" : ""}. Правило их не исправит.
                    </p>
                  )}
                  <p className="undo-note">{plan.undoNote}</p>
                </div>
              );
            })()}
            {pending.name === "sort_range" && (() => {
              const plan = pending.args as any;
              const rows = (m: unknown[][]) => (m ?? []).map((r) => r.map((c) => (c === "" || c === null ? "∅" : String(c))).join(" · ")).join("\n");
              return (
                <div className="preview">
                  <p>
                    {plan.rows} строк × {plan.columns} столбцов; ключ — столбец {plan.column + 1}
                    {plan.keyHeader !== undefined ? ` «${plan.keyHeader}»` : ""}, {plan.ascending ? "по возрастанию" : "по убыванию"};
                    {" "}заголовки {plan.hasHeaders ? "остаются на месте" : "сортируются вместе с данными"}.
                  </p>
                  {plan.headerWarning && <p className="warn-note">{plan.headerWarning}</p>}
                  {plan.formulaWarning && <p className="warn-note">{plan.formulaWarning}</p>}
                  {plan.mergeWarning && <p className="warn-note">{plan.mergeWarning}</p>}
                  <div><strong>Первые строки сейчас</strong><pre>{rows(plan.previewBefore)}</pre></div>
                  <div><strong>Станут (по нашей оценке порядка Excel)</strong><pre>{rows(plan.previewAfter)}</pre></div>
                  <p className="undo-note">
                    {plan.undoAvailable ? "После проверки будет доступна отмена, если данные не изменятся." : plan.undoNote}
                  </p>
                </div>
              );
            })()}
            {pending.name === "apply_filter" && (() => {
              const plan = pending.args as any;
              return (
                <div className="preview">
                  <p>
                    Столбец {plan.column + 1}{plan.columnHeader !== undefined ? ` «${plan.columnHeader}»` : ""}, условие{" "}
                    <strong>{plan.criteriaText}</strong>. Данные не меняются, скрываются строки.
                  </p>
                  {plan.visibleRowsBefore !== null && <p>Сейчас видно строк: {plan.visibleRowsBefore} из {plan.rows}.</p>}
                  {plan.change === "replacesFilter" && (
                    <p className="warn-note">
                      На листе уже стоит фильтр на другой области {plan.before?.address} с условиями в {plan.before?.activeColumns} столбцах.
                      Новый фильтр заменит его целиком, прежние условия будут потеряны.
                    </p>
                  )}
                  {plan.change === "replacesColumn" && (
                    <p className="warn-note">В этом столбце уже есть условие фильтра — оно будет заменено. Условия других столбцов сохранятся.</p>
                  )}
                  {plan.change === "adds" && (
                    <p>Фильтр на этой области уже стоит; новое условие добавится к существующим, и строк может остаться меньше, чем по одному этому условию.</p>
                  )}
                  <p className="undo-note">Прежнюю комбинацию фильтров автоматически не вернуть; снять фильтр можно в Excel.</p>
                </div>
              );
            })()}
            {pending.name === "format_range" && (() => {
              const plan = pending.args as any;
              const label: Record<string, string> = {
                numberFormat: "числовой формат",
                bold: "полужирный",
                italic: "курсив",
                underline: "подчёркивание",
                fontColor: "цвет текста",
                fontSize: "размер шрифта",
                fontName: "шрифт",
                fillColor: "заливка",
                horizontalAlignment: "выравнивание по горизонтали",
                verticalAlignment: "выравнивание по вертикали",
                wrapText: "перенос по словам",
                columnWidth: "ширина столбцов, пт",
                rowHeight: "высота строк, пт",
                borders: "границы"
              };
              const edge: Record<string, string> = {
                EdgeTop: "верх",
                EdgeBottom: "низ",
                EdgeLeft: "лево",
                EdgeRight: "право",
                InsideHorizontal: "внутри гориз.",
                InsideVertical: "внутри верт."
              };
              const show = (value: unknown): string => {
                if (value === null) return "разное в области";
                if (value === undefined) return "не задано";
                if (value === true) return "да";
                if (value === false) return "нет";
                if (typeof value === "object") {
                  // Границы: каждая сторона отдельно, «None» — линии нет.
                  return Object.entries(value as Record<string, unknown>)
                    .map(([key, text]) => `${edge[key] ?? key}: ${text === "None" ? "нет" : text === null ? "разное" : String(text).replace(/\|/g, " ")}`)
                    .join("; ");
                }
                return String(value);
              };
              const autofitLabel: Record<string, string> = {
                columns: "ширина столбцов",
                rows: "высота строк",
                both: "ширина столбцов и высота строк"
              };
              return (
                <div className="preview">
                  <p>{plan.cellCount} ячеек; меняются только перечисленные свойства, остальное оформление не трогается.</p>
                  {plan.areas?.length > 1 && <p>Области ({plan.areas.length}): {plan.areas.join(", ")} — одной операцией, отмена вернёт все сразу.</p>}
                  {plan.mergeWarning && <p className="warn-note">{plan.mergeWarning}</p>}
                  {Object.keys(plan.expected ?? {}).map((key) => (
                    <div key={key}>
                      <strong>{label[key] ?? key}</strong>: {show(plan.before?.[key])} → {show(plan.expected?.[key])}
                      {key === "columnWidth" && plan.columnWidthChars && (
                        <> (в знаках: {show(plan.columnWidthChars.before)} → {show(plan.columnWidthChars.expected)})</>
                      )}
                    </div>
                  ))}
                  {plan.autofit && (
                    <div>
                      <strong>автоподбор</strong>: {autofitLabel[plan.autofit] ?? plan.autofit} — по содержимому
                      <p className="undo-note">{plan.autofitNote}</p>
                    </div>
                  )}
                  <p className="undo-note">
                    {plan.undoAvailable
                      ? "После проверки будет доступна отмена, если оформление не изменится."
                      : plan.undoNote ?? "Автоматическая отмена этой операции недоступна."}
                  </p>
                </div>
              );
            })()}
            {pending.name === "create_workbook_backup" && (
              <div className="preview">
                <p>
                  Сохранить <strong>полную копию книги</strong> — все листы и данные — в папку <code>backups</code> на этом компьютере.
                  Книга не меняется. Копия снимается из открытой книги, включая несохранённые правки.
                </p>
              </div>
            )}
            {pending.name === "create_sheet" && (() => {
              const plan = pending.args as any;
              return (
                <div className="preview">
                  <p>
                    Новый пустой лист <strong>{plan.name}</strong>, {plan.positionText}.
                  </p>
                  <p>Сейчас в книге листы: {(plan.sheetsBefore ?? []).join(", ")}.</p>
                  <p>Данные и другие листы не меняются.</p>
                  <p className="undo-note">{plan.undoNote}</p>
                </div>
              );
            })()}
            {pending.name === "freeze_panes" && (() => {
              const plan = pending.args as any;
              return (
                <div className="preview">
                  <p>
                    Закрепление на листе {plan.target?.sheetName}: <strong>{plan.beforeText}</strong> → <strong>{plan.expectedText}</strong>.
                  </p>
                  <p>Это настройка вида листа: данные и оформление ячеек не меняются.</p>
                  <p className="undo-note">
                    {plan.undoAvailable ? "После проверки будет доступна отмена." : plan.undoNote}
                  </p>
                </div>
              );
            })()}
            {pending.name === "add_conditional_format" && (() => {
              const plan = pending.args as any;
              return (
                <div className="preview">
                  <p>
                    Правило на {plan.cellCount} ячеек: <strong>{plan.ruleText}</strong>.
                  </p>
                  {plan.prediction && (
                    <p>
                      Подсветится примерно {plan.prediction.matches} из {plan.prediction.total}
                      {plan.prediction.sample?.length > 0 && <>: {plan.prediction.sample.join(", ")}{plan.prediction.matches > plan.prediction.sample.length ? "…" : ""}</>}
                      . <span className="undo-note">{plan.prediction.note}</span>
                    </p>
                  )}
                  {plan.request?.rule === "formula" && (
                    <p className="undo-note">Какие ячейки подсветит формула, считает только Excel: оценки у панели нет.</p>
                  )}
                  {plan.functionCheck && <p className="undo-note">{plan.functionCheck.note}</p>}
                  {plan.existingNote && <p className="warn-note">{plan.existingNote}</p>}
                  <p className="undo-note">
                    {plan.undoAvailable ? "После проверки будет доступна отмена: она удалит это правило." : plan.undoNote}
                  </p>
                </div>
              );
            })()}
            {pending.name === "create_table" && (() => {
              const plan = pending.args as any;
              return (
                <div className="preview">
                  <p>
                    Таблица Excel из {plan.rows} строк × {plan.columns} столбцов, стиль <strong>{plan.style}</strong>
                    {plan.name ? <>, имя <strong>{plan.name}</strong></> : ""}.
                  </p>
                  <p>Заголовки: {(plan.headers ?? []).map((h: unknown) => (h === "" || h === null ? "∅" : String(h))).join(" · ")}</p>
                  {plan.headerProblems?.length > 0 && (
                    <div className="warn-note">
                      <strong>Excel изменит заголовки:</strong>
                      {plan.headerProblems.map((text: string, index: number) => <div key={index}>{text}</div>)}
                    </div>
                  )}
                  {plan.manualFormattingWarning && <p className="warn-note">{plan.manualFormattingWarning}</p>}
                  {plan.autoFilterWarning && <p className="warn-note">{plan.autoFilterWarning}</p>}
                  {plan.mergeWarning && <p className="warn-note">{plan.mergeWarning}</p>}
                  <p className="warn-note">{plan.behaviourNote}</p>
                  <p className="undo-note">{plan.undoNote}</p>
                </div>
              );
            })()}
            {pending.name === "create_pivot_table" && (() => {
              const plan = pending.args as any;
              const agg: Record<string, string> = { sum: "сумма", count: "количество", average: "среднее", max: "максимум", min: "минимум" };
              return (
                <div className="preview">
                  <p>
                    Сводная по {plan.target?.sheetName}!{plan.sourceAddress} ({plan.sourceRows} строк данных) займёт{" "}
                    <strong>{plan.destSheet}!{plan.destArea}</strong> — {plan.newSheet ? `лист «${plan.destSheet}» будет создан этой же операцией; отмена уберёт сводную и пустой лист.` : "место свободно."}
                  </p>
                  <p>
                    Строки: {(plan.rowFields ?? []).join(" → ")}. Значения:{" "}
                    {(plan.valueFields ?? []).map((item: any) => `${item.field} (${agg[item.aggregation] ?? item.aggregation})`).join(", ")}.
                  </p>
                  <div>
                    <strong>Расчёт панели</strong>
                    <pre>{(plan.preview ?? []).join("\n")}</pre>
                  </div>
                  {(plan.expectation?.warnings ?? []).map((text: string, index: number) => <p key={index} className="warn-note">{text}</p>)}
                  <p className="undo-note">
                    {plan.undoAvailable ? "После построения итоги будут сверены с этим расчётом; отмена удалит сводную." : plan.undoNote}
                  </p>
                </div>
              );
            })()}
            {pending.name === "create_chart" && (() => {
              const plan = pending.args as any;
              const kinds: Record<string, string> = {
                ColumnClustered: "столбчатая",
                BarClustered: "линейчатая",
                Line: "график",
                Area: "с областями",
                Pie: "круговая",
                Doughnut: "кольцевая",
                XYScatter: "точечная"
              };
              const e = plan.expectation ?? {};
              return (
                <div className="preview">
                  <p>
                    Диаграмма <strong>{kinds[plan.chartType] ?? plan.chartType}</strong> по {plan.resolvedAddress}
                    {plan.title ? <>, заголовок «{plan.title}»</> : ""}; левый верхний угол — {plan.anchorCell}.
                  </p>
                  <p>
                    Рядов: {e.seriesNames?.length} ({e.seriesBy === "rows" ? "по строкам" : "по столбцам"}):{" "}
                    {(e.seriesNames ?? []).join(", ")}. Точек в ряду: {e.pointCount}.
                  </p>
                  {e.categories?.length > 0 && <p>Подписи: {e.categories.join(", ")}{e.pointCount > e.categories.length ? "…" : ""}</p>}
                  {plan.combo?.length > 0 && (
                    <p>
                      Комбинированная: {plan.combo.map((item: any) => `«${item.name}» — ${({ ColumnClustered: "столбцы", Line: "линия", LineMarkers: "линия с точками", Area: "область" } as Record<string, string>)[item.type] ?? item.type}${item.secondary ? " на второй оси" : ""}`).join("; ")}.
                    </p>
                  )}
                  {plan.colorLines?.length > 0 && <p>Цвета: {plan.colorLines.join("; ")}.</p>}
                  {!e.headerRow && <p className="undo-note">Шапки нет: имена рядам Excel даст сам.</p>}
                  {plan.chartsOnSheet > 0 && (
                    <p className="undo-note">
                      На листе уже {plan.chartsOnSheet} диаграмм: если новая попадёт на них, она опустится ниже.
                    </p>
                  )}
                  {(e.warnings ?? []).map((text: string, index: number) => <p key={index} className="warn-note">{text}</p>)}
                  {plan.anchorWarning && <p className="warn-note">{plan.anchorWarning}</p>}
                  <p className="undo-note">
                    {plan.undoAvailable ? "После построения будет доступна отмена: она удалит диаграмму." : plan.undoNote}
                  </p>
                </div>
              );
            })()}
            {pending.name === "format_chart" && (() => {
              const plan = pending.args as any;
              return (
                <div className="preview">
                  <p>Цвета диаграммы «<strong>{plan.chartName}</strong>» на листе {plan.target?.sheetName}:</p>
                  <ul>{plan.lines.map((line: string) => <li key={line}>{line}</li>)}</ul>
                  <p>Данные и остальное оформление диаграммы не меняются.</p>
                  {plan.hasFill && <p className="warn-note">Заливку эта версия Excel не даёт прочитать: прежний цвет вернуть отменой нельзя, а новый не сверяется.</p>}
                  <p className="undo-note">{plan.undoAvailable ? "После проверки будет доступна отмена: вернутся прежние цвета линий и маркеров." : plan.undoNote}</p>
                </div>
              );
            })()}
            {(pending.name === "insert_rows" || pending.name === "delete_rows") && (() => {
              const plan = pending.args as any;
              const deleting = pending.name === "delete_rows";
              const rows = (m: unknown[][]) =>
                (m ?? []).map((r) => r.map((c) => (c === "" || c === null ? "∅" : String(c))).join(" · ")).join("\n");
              const risk: Record<string, string> = {
                broken: "станет #ССЫЛКА!",
                shrunk: "диапазон уменьшится, итог изменится молча",
                missed: "не охватит новые строки"
              };
              return (
                <div className="preview">
                  <p>
                    {deleting ? "Удаление" : "Вставка"} строк <strong>{plan.rowsAddress}</strong> на листе{" "}
                    {plan.target?.sheetName}. {deleting
                      ? `Нижние строки поднимутся вверх; непустых ячеек в удаляемых строках: ${plan.filledCells}.`
                      : "Существующие строки сдвинутся вниз, адреса ниже точки вставки изменятся."}
                  </p>
                  {deleting && plan.preview?.length > 0 && (
                    <div>
                      <strong>Будет удалено{plan.previewTruncated ? " (показана часть)" : ""}</strong>
                      <pre>{rows(plan.preview)}</pre>
                    </div>
                  )}
                  {plan.formulaRisks?.length > 0 && (
                    <div className="warn-note">
                      <strong>Пострадают формулы книги:</strong>
                      {plan.formulaRisks.map((item: any, index: number) => (
                        <div key={index}>
                          {item.sheet}!{item.address}: <code>{item.formula}</code> — ссылка {item.reference} {risk[item.kind]}
                        </div>
                      ))}
                      {plan.riskOverflow > 0 && <div>…и ещё {plan.riskOverflow} таких ссылок.</div>}
                    </div>
                  )}
                  {plan.tableFormulaSheets?.length > 0 && (
                    <p className="warn-note">
                      Формулы со ссылками на таблицы (листы {plan.tableFormulaSheets.join(", ")}) не разбирались — что станет с ними, здесь не проверено.
                    </p>
                  )}
                  {plan.unscannedSheets?.length > 0 && (
                    <p className="warn-note">
                      Листы {plan.unscannedSheets.join(", ")} слишком велики для обхода формул: про них ничего не проверено.
                    </p>
                  )}
                  {plan.tableWarning && <p className="warn-note">{plan.tableWarning}</p>}
                  {plan.mergeWarning && <p className="warn-note">{plan.mergeWarning}</p>}
                  <p className="undo-note">{plan.undoNote}</p>
                  {!plan.backup && deleting && (
                    <p className="warn-note">
                      Резервной копии в этом сеансе не создавалось. Прежде чем подтверждать, имеет смысл попросить копию книги.
                    </p>
                  )}
                </div>
              );
            })()}
            {!["__read_sheets", "create_workbook_backup", "set_range_values", "set_ranges_values", "fill_range", "format_range", "sort_range", "apply_filter", "insert_rows", "delete_rows", "freeze_panes", "add_conditional_format", "create_table", "create_chart", "format_chart", "create_pivot_table", "create_sheet", "trim_text", "convert_values", "change_case", "remove_duplicates", "rename_sheet", "set_page_layout", "copy_sheet", "add_multiples", "remember_preference", "save_scenario", "import_file_layout", "import_file_table", "delete_sheet", "insert_columns", "delete_columns", "group_rows_columns", "set_data_validation", "convert_table_to_range", "move_conditional_format", "apply_color_convention", "add_share_growth", "add_comparison", "build_three_statement_model", "build_dcf_model", "build_lbo_model"].includes(pending.name) && (
              <pre>{JSON.stringify(pending.args, null, 2)}</pre>
            )}
            <div className="row">
              <button className="apply" onClick={() => decide(true)}>
                {pending.name === "__read_sheets" ? "Разрешить" : "Выполнить"}
              </button>
              <button onClick={() => decide(false)}>{pending.name === "__read_sheets" ? "Не разрешать" : "Отклонить"}</button>
            </div>
          </div>
        )}

        {streaming && (awaitingLead
          ? <div className="lead"><Logo size={24} className="avatar" /><div className="msg assistant"><Markdown text={streaming} /></div></div>
          : <div className="msg assistant"><Markdown text={streaming} /></div>)}
        {busy && !streaming && !pending && <div className={`thinking${awaitingLead ? " first" : ""}`}><Sparkle size={15} />Думает<span className="dots"><i /><i /><i /></span></div>}

        <div ref={logEnd} />
      </div>

      <div className="composer">
        {files.length > 0 && (
          <ul className="attached" aria-label="Прикреплённые файлы">
            {files.map((file) => (
              <li key={file.id} title={file.warnings.join("\n")}>
                <span>📎 {file.name}</span> <span className="hint">{describeFile(file)}</span>
                <button className="ghost" disabled={busy} aria-label={`Убрать ${file.name}`}
                  onClick={() => void removeFile(file.id).then(setFiles, (error) => setEntries((e) => [...e, { kind: "error", text: String(error?.message ?? error) }]))}>×</button>
              </li>
            ))}
          </ul>
        )}
        <input ref={fileInput} type="file" hidden accept=".csv,.tsv,.txt,.xlsx,.docx,.pdf"
          onChange={(event) => {
            const chosen = event.target.files?.[0];
            event.target.value = "";
            if (!chosen) return;
            setUploading(true);
            uploadFile(chosen)
              .then((file) => {
                setFiles((current) => [...current.filter((item) => item.id !== file.id), file]);
                setEntries((e) => [...e, { kind: "notice", text: `Файл «${file.name}» разобран на этом компьютере: ${describeFile(file)}.${file.warnings.length ? ` ${file.warnings.join(" ")}` : ""}` }]);
              })
              .catch((error) => setEntries((e) => [...e, { kind: "error", text: `Файл не прикреплён: ${error?.message ?? error}` }]))
              .finally(() => setUploading(false));
          }} />
        <div className="input-box">
          <textarea
            value={draft}
            onChange={(e) => setDraft(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter" && !e.shiftKey) {
                e.preventDefault();
                send();
              }
            }}
            placeholder="Что сделать с книгой?"
            disabled={busy}
            rows={2}
          />
          <div className="row">
            {!WEB_PANEL && (
              <button className="ghost icon-text" onClick={() => fileInput.current?.click()} disabled={busy || uploading} title="CSV, XLSX, DOCX, PDF или TXT до 20 МБ. Файл разбирается на этом компьютере.">
                <Icon.clip /><span>{uploading ? "Разбор…" : "Файл"}</span>
              </button>
            )}
            <span className="hint">{spendingTotal.tasks ? `Беседа: ${conversationTotalText(spendingTotal)}` : "Enter — отправить"}</span>
            <span className="spacer" />
            {taskRunning ? (
              <button
                className="send stop"
                title="Остановить"
                onClick={() => {
                  pending?.resolve(false);
                  setPending(null);
                  abort.current?.abort();
                }}
              >
                <Icon.stop /><span className="sr">Остановить</span>
              </button>
            ) : (
              <button className="send" title="Отправить (Enter)" onClick={() => send()} disabled={busy || !draft.trim() || !model}>
                <Icon.send /><span className="sr">Отправить</span>
              </button>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
