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
import { formatCapabilityLog, measureCapabilities, type CapabilityReport } from "../excel/capabilities";
import { createMetrics, formatBytes, formatMs, type MetricsSummary } from "../agent/metrics";
import {
  clearConversation,
  conversationKey,
  currentWorkbookUrl,
  loadConversation,
  RESTORED_CONVERSATION_NOTICE,
  saveConversation
} from "./persistence";

type Entry =
  | { kind: "user"; text: string }
  | { kind: "assistant"; text: string }
  | { kind: "error"; text: string }
  | { kind: "op"; event: ToolEvent };

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
  const addr = (a.address ?? a.sourceAddress ?? a.destAddress) as string | undefined;
  if (addr) return `${sheet}${addr}`;
  if (typeof a.startRow === "number") return `${sheet}строки ${a.startRow}–${a.startRow + Number(a.count ?? 1) - 1}`;
  return "";
}

export default function Taskpane() {
  const [providers, setProviders] = useState<ProviderInfo[]>([]);
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
  const [capabilities, setCapabilities] = useState<CapabilityReport | null>(null);
  const [showDiagnostics, setShowDiagnostics] = useState(false);
  const [metricsSummary, setMetricsSummary] = useState<MetricsSummary | null>(null);
  const [restored, setRestored] = useState(false);

  const history = useRef<ChatMessage[]>([]);
  // Лента нужна в finally блока отправки, где значение state уже устарело.
  const entriesRef = useRef<Entry[]>([]);
  const metrics = useRef(createMetrics());
  const storageKey = useRef(conversationKey(undefined));
  // Указание об устаревании данных отдаётся модели один раз, при первом
  // сообщении после восстановления беседы.
  const pendingNotices = useRef<string[]>([]);
  const abort = useRef<AbortController | null>(null);
  const logEnd = useRef<HTMLDivElement>(null);

  useEffect(() => {
    fetchProviders()
      .then((list) => {
        setProviders(list);
        if (list.length) {
          setProvider(list[0].id);
          setModel(list[0].defaultModel);
        } else {
          setEntries((e) => [
            ...e,
            { kind: "error", text: "Ни одного ключа не найдено. Заполните server/.env и перезапустите прокси." }
          ]);
        }
      })
      .catch((err) => setEntries((e) => [...e, { kind: "error", text: String(err.message ?? err) }]));
  }, []);

  // Замер возможностей установленного Excel и восстановление беседы.
  useEffect(() => {
    const report = measureCapabilities();
    setCapabilities(report);
    // В журнал уходят только версии и признаки поддержки: содержимого ячеек
    // и ключей в этой строке нет по построению.
    console.info(`[excel-ai] возможности: ${formatCapabilityLog(report)}`);

    const workbook = currentWorkbookUrl();
    const key = conversationKey(workbook);
    storageKey.current = key;

    const saved = loadConversation<Entry>(key);
    if (saved) {
      history.current = saved.history;
      setEntries(saved.entries);
      setRestored(true);
      pendingNotices.current = [RESTORED_CONVERSATION_NOTICE];
    }
  }, []);

  // Сохраняем после каждого изменения ленты: она меняется на любое событие —
  // сообщение пользователя, ответ модели, операцию с книгой.
  useEffect(() => {
    entriesRef.current = entries;
    if (!entries.length) return;
    saveConversation(storageKey.current, currentWorkbookUrl(), entries, history.current);
  }, [entries]);

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
    setEntries((e) => [...e, { kind: "user", text }]);
    history.current.push({ role: "user", content: text });
    setBusy(true);
    setStreaming("");

    const controller = new AbortController();
    abort.current = controller;

    try {
      const notices = pendingNotices.current;
      pendingNotices.current = [];

      await runAgent({
        provider,
        model,
        history: history.current,
        signal: controller.signal,
        ...(notices.length ? { notices } : {}),
        hooks: {
          onMetric: (m) => {
            metrics.current.record(m);
            setMetricsSummary(metrics.current.summary());
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
      if (err?.name !== "AbortError") {
        setEntries((e) => [...e, { kind: "error", text: err?.message ?? String(err) }]);
      }
    } finally {
      setBusy(false);
      setStreaming("");
      setPending(null);
      abort.current = null;
      refreshUndoState();
      saveConversation(storageKey.current, currentWorkbookUrl(), entriesRef.current, history.current);
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
    history.current = [];
    setEntries([]);
    setStreaming("");
    setRestored(false);
    pendingNotices.current = [];
    metrics.current.reset();
    setMetricsSummary(null);
    clearConversation(storageKey.current);
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
        <button
          className="ghost"
          onClick={() => setShowDiagnostics((v) => !v)}
          title="Возможности установленного Excel и измерения"
        >
          {showDiagnostics ? "Скрыть диагностику" : "Диагностика"}
        </button>
        <button className="ghost" onClick={reset} disabled={busy}>
          Очистить
        </button>
      </div>

      {showDiagnostics && (
        <div className="diag">
          {capabilities ? (
            <>
              <div className="diag-line">
                {capabilities.host}, {capabilities.platform}, Office {capabilities.officeVersion}
              </div>
              <div className="diag-line">
                Потолок ExcelApi: <strong>{capabilities.ceiling ?? "не определён"}</strong>
              </div>
              <ul className="diag-list">
                {capabilities.features.map((f) => (
                  <li key={f.id} className={f.available ? "yes" : "no"}>
                    <span>{f.available ? "✓" : "✕"}</span> {f.label}
                    {f.reason && <div className="diag-reason">{f.reason}</div>}
                  </li>
                ))}
              </ul>
            </>
          ) : (
            <div className="diag-line">Возможности ещё не измерены.</div>
          )}

          {metricsSummary && metricsSummary.count > 0 && (
            <div className="diag-line">
              Измерения: {metricsSummary.modelCount} запросов к модели, {metricsSummary.toolCount} операций,{" "}
              {formatMs(metricsSummary.totalMs)}, {formatBytes(metricsSummary.totalBytes)}
              {metricsSummary.slowest && ` · самая долгая: ${metricsSummary.slowest.name} ${formatMs(metricsSummary.slowest.ms)}`}
              {metricsSummary.largest && ` · самый крупный результат: ${metricsSummary.largest.name} ${formatBytes(metricsSummary.largest.bytes)}`}
            </div>
          )}
        </div>
      )}

      <div className="log">
        {restored && (
          <div className="msg restored">
            Беседа восстановлена. Данные книги могли измениться — агент перечитает нужные диапазоны перед выводами.
          </div>
        )}

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
                {event.status === "done" && event.undoable === false && " — без автоматической отмены"}
                {event.undoNote && <div className="undo-note">{event.undoNote}</div>}
                {event.status === "error" && ` — ${event.result}`}
                {event.status === "done" && typeof event.ms === "number" && (
                  <span className="meta">
                    {" "}
                    {formatMs(event.ms)}
                    {typeof event.bytes === "number" ? `, ${formatBytes(event.bytes)}` : ""}
                  </span>
                )}
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
            {pending.name === "format_range" && (
              <p className="undo-note">Для больших диапазонов точная автоматическая отмена форматирования может быть недоступна.</p>
            )}
            {pending.name === "insert_rows" && (
              <p className="undo-note">Вставка строк структурная: собственного undo нет, а вся предыдущая история custom undo будет очищена.</p>
            )}
            {pending.name === "delete_rows" && (
              <p className="undo-note">Удаление строк необратимо для custom undo: собственного undo нет, а вся предыдущая история custom undo будет очищена.</p>
            )}
            <pre>{JSON.stringify(pending.args, null, 2)}</pre>
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
