import { useRef, useState } from "react";
import {
  conversationFileName,
  conversationMarkdown,
  deleteAllConversations,
  deleteConversation,
  listConversations,
  readConversation,
  type ConversationsOverview,
  type StoredConversation
} from "./conversationStore";
import { copyText, exportConversationFile } from "./api/conversations";
import { WEB_PANEL } from "./panelMode";

/**
 * Окно «Беседы» (этап 10, 10.1): все беседы этого компьютера — список,
 * чтение, сохранение в файл или копирование, удаление. Беседа другой книги
 * открывается только для чтения: продолжать её здесь нельзя — история и
 * разрешения на чтение листов привязаны к своей книге.
 */

function storage(): Storage | null {
  try {
    return localStorage;
  } catch {
    return null;
  }
}

const date = (ms: number) => new Date(ms).toLocaleDateString("ru-RU", { day: "numeric", month: "long" });
const dateTime = (ms: number) => new Date(ms).toLocaleString("ru-RU", { day: "numeric", month: "long", hour: "2-digit", minute: "2-digit" });
const kb = (bytes: number) => `${Math.max(1, Math.round(bytes / 1024))} КБ`;

function plural(count: number, one: string, few: string, many: string) {
  const mod10 = count % 10;
  const mod100 = count % 100;
  const word = mod10 === 1 && mod100 !== 11 ? one : mod10 >= 2 && mod10 <= 4 && (mod100 < 12 || mod100 > 14) ? few : many;
  return `${count} ${word}`;
}

export default function ConversationsPanel({ currentKey, onClose, onDeletedCurrent }: {
  currentKey: string | null;
  onClose: () => void;
  /** Удалена беседа открытой книги — панель очищает ленту. */
  onDeletedCurrent: () => void;
}) {
  const store = storage();
  const [overview, setOverview] = useState<ConversationsOverview | null>(() => (store ? listConversations(store) : null));
  const [open, setOpen] = useState<StoredConversation | null>(null);
  const [confirm, setConfirm] = useState<string | null>(null);
  const [note, setNote] = useState("");
  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);
  /** Буфер обмена недоступен — текст показывается, чтобы скопировать вручную. */
  const [manualCopy, setManualCopy] = useState("");
  const copyArea = useRef<HTMLTextAreaElement>(null);

  const refresh = () => setOverview(store ? listConversations(store) : null);
  const ordered = overview
    ? [...overview.conversations].sort((a, b) => Number(b.workbookKey === currentKey) - Number(a.workbookKey === currentKey))
    : [];

  function show(key: string) {
    if (!store) return;
    setNote("");
    setError("");
    setManualCopy("");
    setOpen(readConversation(store, key));
  }

  function remove(key: string | "all") {
    if (!store) return;
    if (key === "all") deleteAllConversations(store);
    else deleteConversation(store, key);
    if (key === "all" || key === currentKey) onDeletedCurrent();
    setConfirm(null);
    setOpen(null);
    setNote(key === "all" ? "Все беседы удалены." : "Беседа удалена.");
    refresh();
  }

  async function save(conversation: StoredConversation) {
    setNote("");
    setError("");
    const text = conversationMarkdown(conversation);
    if (WEB_PANEL) {
      const result = await copyText(text, copyArea.current);
      setManualCopy(result === "copied" ? "" : text);
      setNote(result === "copied"
        ? "Беседа скопирована текстом — вставьте её в заметку или документ."
        : "Скопировать автоматически не вышло: выделите текст в поле ниже (⌘A, ⌘C или Ctrl+A, Ctrl+C).");
      return;
    }
    setSaving(true);
    try {
      const path = await exportConversationFile(conversationFileName(conversation), text);
      // У многих «Документы» перенесены в OneDrive — тогда файл уходит и в облако.
      setNote(`Сохранено: ${path}${/[\\/]OneDrive[^\\/]*[\\/]/i.test(path)
        ? ". Эта папка синхронизируется с OneDrive — файл попадёт и в облако Microsoft."
        : ""}`);
    } catch (err: any) {
      setError(String(err?.message ?? err));
    } finally {
      setSaving(false);
    }
  }

  if (!store) {
    return (
      <section className="keys" aria-label="Беседы">
        <div className="keys-head"><strong>Беседы</strong><span className="spacer" /><button className="ghost" onClick={onClose}>Готово</button></div>
        <div className="warn-note">Хранилище панели недоступно: беседы здесь не сохраняются.</div>
      </section>
    );
  }

  return (
    <section className="keys conversations" aria-label="Беседы">
      <div className="keys-head">
        <strong>{open ? "Беседа" : "Беседы"}</strong>
        <span className="spacer" />
        {open && <button className="ghost" onClick={() => { setOpen(null); setNote(""); }}>← К списку</button>}
        <button className="ghost" onClick={onClose}>Готово</button>
      </div>

      {!open && overview && (
        <>
          <p className="key-note">
            Хранятся только на этом компьютере, {overview.retentionDays} дней с последнего сообщения, не больше{" "}
            {overview.maxConversations} бесед. Занято {kb(overview.usedBytes)} из {kb(overview.limitBytes)}. В беседах могут
            быть данные ячеек.
          </p>
          {!ordered.length && <p className="key-note">Сохранённых бесед нет.</p>}
          <ul className="key-list conversation-list">
            {ordered.map((item) => (
              <li key={item.workbookKey}>
                <button className="conversation-open" onClick={() => show(item.workbookKey)} title={item.documentUrl || "Книга ещё не сохранялась"}>
                  <strong>{item.workbookName}</strong>
                  {item.workbookKey === currentKey && <span className="hint"> — эта книга</span>}
                  <span className="conversation-title">{item.title}</span>
                  <span className="hint">
                    {dateTime(item.updatedAt)} · {plural(item.messages, "сообщение", "сообщения", "сообщений")}
                    {item.actions ? ` · ${plural(item.actions, "действие", "действия", "действий")}` : ""} · удалится {date(item.expiresAt)}
                  </span>
                </button>
              </li>
            ))}
          </ul>
          {ordered.length > 0 && (confirm === "all" ? (
            <div className="warn-note">
              Удалить все беседы ({ordered.length})? Книги не изменятся.{" "}
              <button className="apply" onClick={() => remove("all")}>Удалить все</button>{" "}
              <button className="ghost" onClick={() => setConfirm(null)}>Нет</button>
            </div>
          ) : (
            <button className="ghost" onClick={() => setConfirm("all")}>Удалить все беседы</button>
          ))}
        </>
      )}

      {open && (
        <>
          <p className="key-note">
            <strong>{conversationFileName(open).replace(/ \d{4}-\d{2}-\d{2} \d{2}-\d{2}\.md$/, "")}</strong> · {dateTime(open.updatedAt)}
            {open.workbookKey !== currentKey && " · только чтение: продолжить эту беседу можно в её книге"}
          </p>
          <div className="row">
            <button className="apply" onClick={() => void save(open)} disabled={saving}>
              {WEB_PANEL ? "Скопировать текст" : saving ? "Сохранение…" : "Сохранить в файл"}
            </button>
            {confirm === open.workbookKey ? (
              <>
                <button className="apply" onClick={() => remove(open.workbookKey)}>Точно удалить</button>
                <button className="ghost" onClick={() => setConfirm(null)}>Нет</button>
              </>
            ) : (
              <button className="ghost" onClick={() => setConfirm(open.workbookKey)}>Удалить беседу</button>
            )}
          </div>
          <div className="conversation-read" aria-label="Текст беседы">
            {open.entries.map((entry, index) => {
              if (entry.kind === "op") return <div key={index} className={`op ${entry.event.status}`}><span className="name">{entry.event.name}</span></div>;
              const cls = entry.kind === "user" ? "msg user" : entry.kind === "assistant" ? "msg assistant" : entry.kind === "error" ? "msg error" : "msg notice";
              return <div key={index} className={cls}>{entry.text}</div>;
            })}
          </div>
        </>
      )}

      {note && <div className="key-note" role="status">{note}</div>}
      {error && <div className="key-error" role="alert">{error}</div>}
      {manualCopy && <textarea className="copy-visible" readOnly value={manualCopy} aria-label="Текст беседы для копирования" onFocus={(event) => event.currentTarget.select()} />}
      <textarea ref={copyArea} className="copy-area" readOnly aria-hidden="true" tabIndex={-1} />
    </section>
  );
}
