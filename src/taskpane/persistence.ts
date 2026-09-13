/**
 * Сохранение беседы между открытиями панели (решение Р0-3, минимум).
 *
 * Хранилище — localStorage панели. Оно не уезжает вместе с файлом книги, в
 * отличие от Office.context.document.settings и CustomXmlParts: те лежат внутри
 * .xlsx, и отправив книгу коллеге, пользователь отправил бы всю переписку с ИИ
 * вместе с прочитанными данными. Постоянное хранилище на сервере — Р3-3.
 *
 * Снимки данных книги здесь не сохраняются: при возврате к беседе правильнее
 * перечитать книгу, потому что данные могли измениться.
 */

import type { ChatMessage, ToolCall } from "./api/client";

export const STORAGE_PREFIX = "excel-ai-chat:";
/** Беседа крупнее предела не сохраняется целиком: старые записи отбрасываются. */
export const MAX_STORED_BYTES = 2_000_000;

export interface StoredConversation<TEntry> {
  version: 1;
  /** Книга, к которой относится беседа. Пусто — книга не определилась. */
  workbook: string;
  savedAt: number;
  entries: TEntry[];
  history: ChatMessage[];
}

function toolCallsOf(message: ChatMessage): ToolCall[] {
  return message.role === "assistant" && message.tool_calls?.length ? message.tool_calls : [];
}

/**
 * Беседа могла прерваться посреди шага: в истории остался ответ модели с
 * tool_calls без ответных сообщений роли tool. Провайдер такой запрос отклонит,
 * поэтому каждый незакрытый вызов закрывается синтетическим результатом, а
 * ответы-сироты без своего вызова отбрасываются.
 */
export function repairHistory(history: ChatMessage[]): { history: ChatMessage[]; repaired: number } {
  const out: ChatMessage[] = [];
  let repaired = 0;
  let pending: string[] = [];

  const closePending = () => {
    for (const id of pending) {
      out.push({
        role: "tool",
        tool_call_id: id,
        content: JSON.stringify({
          ok: false,
          error: "Результат неизвестен: беседа была прервана до завершения операции."
        })
      });
      repaired += 1;
    }
    pending = [];
  };

  for (const message of history) {
    if (message.role === "tool") {
      const index = pending.indexOf(message.tool_call_id);
      if (index === -1) {
        // Ответ без своего вызова: отправлять его нельзя.
        repaired += 1;
        continue;
      }
      pending.splice(index, 1);
      out.push(message);
      continue;
    }

    // Любое не-tool сообщение закрывает набор ожидающих вызовов.
    closePending();
    out.push(message);
    pending = toolCallsOf(message).map((c) => c.id);
  }

  closePending();
  return { history: out, repaired };
}

/** Ключ беседы. Разные книги не должны смешивать переписку. */
export function conversationKey(workbookUrl: string | undefined): string {
  const id = (workbookUrl ?? "").trim();
  return STORAGE_PREFIX + (id || "unknown-workbook");
}

export function currentWorkbookUrl(): string {
  try {
    return String(Office?.context?.document?.url ?? "");
  } catch {
    return "";
  }
}

/**
 * Урезает беседу до предела по байтам, отбрасывая самое старое. История и лента
 * режутся согласованно по числу отброшенных пользовательских ходов, иначе лента
 * и контекст модели разойдутся.
 */
export function trimToLimit<TEntry>(
  payload: StoredConversation<TEntry>,
  maxBytes: number = MAX_STORED_BYTES
): StoredConversation<TEntry> {
  let current = payload;
  while (JSON.stringify(current).length > maxBytes) {
    if (current.history.length <= 2 || current.entries.length <= 1) return current;
    const history = current.history.slice(1);
    // После среза начало истории может оказаться ответом-сиротой.
    const { history: safe } = repairHistory(history);
    current = { ...current, history: safe, entries: current.entries.slice(1) };
  }
  return current;
}

export function saveConversation<TEntry>(
  key: string,
  workbook: string,
  entries: TEntry[],
  history: ChatMessage[]
): boolean {
  if (!entries.length && !history.length) {
    clearConversation(key);
    return true;
  }
  const payload = trimToLimit<TEntry>({
    version: 1,
    workbook,
    savedAt: Date.now(),
    entries,
    history
  });
  try {
    localStorage.setItem(key, JSON.stringify(payload));
    return true;
  } catch {
    // Приватное окно, запрет на данные сайта, переполнение — сохранение
    // не обязано работать, но панель обязана работать без него.
    return false;
  }
}

export function loadConversation<TEntry>(key: string): StoredConversation<TEntry> | null {
  let raw: string | null = null;
  try {
    raw = localStorage.getItem(key);
  } catch {
    return null;
  }
  if (!raw) return null;

  try {
    const parsed = JSON.parse(raw) as StoredConversation<TEntry>;
    if (parsed?.version !== 1 || !Array.isArray(parsed.history) || !Array.isArray(parsed.entries)) {
      return null;
    }
    const { history } = repairHistory(parsed.history);
    return { ...parsed, history };
  } catch {
    return null;
  }
}

export function clearConversation(key: string): void {
  try {
    localStorage.removeItem(key);
  } catch {
    // Нечего делать: отсутствие хранилища не является ошибкой работы панели.
  }
}

/**
 * Разовое указание модели при возврате к сохранённой беседе. Без него агент
 * продолжит считать по цифрам, которых в книге уже нет.
 */
export const RESTORED_CONVERSATION_NOTICE =
  "Эта беседа восстановлена из сохранённой истории, возможно через долгое время. " +
  "Данные книги могли измениться: всё прочитанное ранее считай историческим состоянием. " +
  "Перед выводами и любыми правками перечитай нужные диапазоны заново.";
