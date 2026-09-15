import { useEffect, useRef, useState } from "react";
import { fetchProviders, type ChatMessage, type ProviderInfo } from "./api/client";
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
  loadConversation,
  saveConversation,
  type PersistedEntry
} from "./conversationStore";

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

export default function Taskpane() {
  const [providers, setProviders] = useState<ProviderInfo[]>([]);
  const [apiStatus, setApiStatus] = useState<"checking" | "ready" | "error">("checking");
  const [apiError, setApiError] = useState("");
  const [provider, setProvider] = useState("");
  const [model, setModel] = useState("");
  const [entries, setEntries] = useState<Entry[]>([]);
  const [draft, setDraft] = useState("");
  const [busy, setBusy] = useState(false);
  const [streaming, setStreaming] = useState("");
  const [pending, setPending] = useState<Pending | null>(null);
  const [canUndo, setCanUndo] = useState(false);
  const [undoLabel, setUndoLabel] = useState<string | null>(null);
  const [undoAvailable, setUndoAvailable] = useState(false);
  const [monitorStatus, setMonitorStatus] = useState<"connecting" | "ready" | "unsupported" | "error">("connecting");
  const [analysisOnly, setAnalysisOnly] = useState(true);
  const [contextLabel, setContextLabel] = useState("Книга: проверка…");
  const [persistenceNote, setPersistenceNote] = useState("История: проверка привязки…");

  const history = useRef<ChatMessage[]>([]);
  const workbookBinding = useRef<{ key: string; url: string } | null>(null);
  const persistenceReady = useRef(false);
  const abort = useRef<AbortController | null>(null);
  const logEnd = useRef<HTMLDivElement>(null);

  useEffect(() => {
    void loadProviders();
    void refreshContext();
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
      setContextLabel(`${book} · ${context.activeSheet.name} · ${context.selectedAreas.join(", ")}`);
      const key = conversationIdentity(url);
      if (!key) {
        if (workbookBinding.current) {
          history.current = [];
          setEntries([]);
          setAnalysisOnly(true);
          permissionsReset = true;
        }
        workbookBinding.current = null;
        persistenceReady.current = true;
        setPersistenceNote("Несохранённая книга: беседа не восстанавливается автоматически.");
      } else if (workbookBinding.current?.key !== key) {
        persistenceReady.current = false;
        workbookBinding.current = { key, url };
        const restored = loadConversation(localStorage, key);
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
          setProvider(list[0].id);
          setModel(list[0].defaultModel);
        } else {
          setProvider("");
          setModel("");
          setApiError("Сервер работает, но ключи провайдеров не найдены в server/.env.");
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

  async function send() {
    const text = draft.trim();
    if (!text || busy || !provider || !model) return;

    setDraft("");
    setBusy(true);
    setStreaming("");

    const controller = new AbortController();
    abort.current = controller;

    try {
      const refreshed = await refreshContext();
      setEntries((e) => [...e, { kind: "user", text }]);
      history.current.push({ role: "user", content: text });
      await runAgent({
        provider,
        model,
        history: history.current,
        analysisOnly: refreshed?.permissionsReset ? true : analysisOnly,
        ...(refreshed ? { initialContext: refreshed.context } : {}),
        signal: controller.signal,
        hooks: {
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
      setBusy(false);
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

  async function undo() {
    try {
      const msg = await undoLast();
      setEntries((e) => [...e, { kind: "assistant", text: msg }]);
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

  return (
    <div className="pane">
      <div className="head">
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
          className="ghost"
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
          {!undoAvailable ? "Undo недоступен" : undoLabel ? `Отменить: ${undoLabel}` : "Отменить"}
        </button>
        {monitorStatus === "error" && (
          <button className="ghost" onClick={() => void connectUndoMonitor()} disabled={busy}>
            Повторить защиту undo
          </button>
        )}
        <button className="ghost" onClick={reset} disabled={busy}>
          Очистить
        </button>
      </div>

      <div className="context-bar">
        <span title={contextLabel}>{contextLabel}</span>
        <label>
          <input type="checkbox" checked={analysisOnly} onChange={(event) => setAnalysisOnly(event.target.checked)} disabled={busy} />
          Только анализ
        </label>
      </div>
      <div className="persistence-note">{persistenceNote}</div>

      {(apiStatus !== "ready" || apiError) && (
        <div className="api-status" role="status">
          {apiStatus === "checking" ? "Проверка локального API…" : apiError}
          {apiStatus !== "checking" && <button className="ghost" onClick={() => void loadProviders()}>Повторить подключение</button>}
        </div>
      )}

      <div className="log">
        {entries.length === 0 && (
          <div className="empty">
            <p>Опишите, что сделать с книгой. Модель сама прочитает нужные диапазоны.</p>
            <p>
              Например: <code>посчитай итоги по столбцу D и выдели их жирным</code>
            </p>
            <p>Записи и удаления запрашивают подтверждение перед выполнением.</p>
          </div>
        )}

        {entries.map((e, i) => {
          if (e.kind === "op") {
            const { event } = e;
            const addr = addressOf(event.args);
            return (
              <div key={`${event.id}-${i}`} className={`op ${event.status}`}>
                <span className="name">{event.name}</span>
                {addr && (
                  <>
                    {" "}
                    <span className="addr">{addr}</span>
                  </>
                )}
                {event.status === "rejected" && " — отклонено"}
                {event.status === "cancelled" && " — отменено"}
                {event.status === "uncertain" && " — проверьте книгу перед новой правкой"}
                {event.status === "done" && event.undoable === false && " — без автоматической отмены"}
                {event.undoNote && <div className="undo-note">{event.undoNote}</div>}
                {event.status === "error" && ` — ${event.result}`}
              </div>
            );
          }
          return (
            <div key={i} className={`msg ${e.kind}`}>
              {e.text}
            </div>
          );
        })}

        {pending && (
          <div className="confirm">
            <p>
              Разрешить <strong>{pending.name}</strong>
              {addressOf(pending.args) ? ` в ${addressOf(pending.args)}` : ""}? Операция изменит книгу.
            </p>
            {pending.name === "set_range_values" && (() => {
              const plan = pending.args as any;
              return (
                <div className="preview">
                  <p>
                    {plan.cellCount} ячеек; режим записи: <strong>{plan.isFormula ? "формулы" : "литеральные значения"}</strong>;
                    {" "}заменяемых формул: {plan.replacedFormulaCount}; проверка: {plan.errorScanAddress}
                  </p>
                  {plan.mergeWarning && <p className="warn-note">{plan.mergeWarning}</p>}
                  <div><strong>Было</strong><pre>{JSON.stringify(plan.before, null, 2)}</pre></div>
                  <div><strong>Станет</strong><pre>{JSON.stringify(plan.after, null, 2)}</pre></div>
                  <p className="undo-note">{plan.undoAvailable ? "После проверки будет доступна безопасная отмена, если диапазон не изменится." : "Автоматическая отмена сейчас недоступна."}</p>
                </div>
              );
            })()}
            {pending.name === "format_range" && (
              <p className="undo-note">Для больших диапазонов точная автоматическая отмена форматирования может быть недоступна.</p>
            )}
            {pending.name === "insert_rows" && (
              <p className="undo-note">Вставка строк структурная: собственного undo нет, а вся предыдущая история custom undo будет очищена.</p>
            )}
            {pending.name === "delete_rows" && (
              <p className="undo-note">Удаление строк необратимо для custom undo: собственного undo нет, а вся предыдущая история custom undo будет очищена.</p>
            )}
            {pending.name !== "set_range_values" && <pre>{JSON.stringify(pending.args, null, 2)}</pre>}
            <div className="row">
              <button className="apply" onClick={() => decide(true)}>
                Выполнить
              </button>
              <button onClick={() => decide(false)}>Отклонить</button>
            </div>
          </div>
        )}

        {streaming && <div className="msg assistant">{streaming}</div>}
        {busy && !streaming && !pending && <div className="thinking">Думает</div>}

        <div ref={logEnd} />
      </div>

      <div className="composer">
        <textarea
          value={draft}
          onChange={(e) => setDraft(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && !e.shiftKey) {
              e.preventDefault();
              void send();
            }
          }}
          placeholder="Что сделать с книгой?"
          disabled={busy}
        />
        <div className="row">
          <span className="hint">Enter — отправить, Shift+Enter — перенос</span>
          <span className="spacer" />
          {busy ? (
            <button
              className="send"
              onClick={() => {
                pending?.resolve(false);
                setPending(null);
                abort.current?.abort();
              }}
            >
              Остановить
            </button>
          ) : (
            <button className="send" onClick={() => void send()} disabled={!draft.trim() || !model}>
              Отправить
            </button>
          )}
        </div>
      </div>
    </div>
  );
}
