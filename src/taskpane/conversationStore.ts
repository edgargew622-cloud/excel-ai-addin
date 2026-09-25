import type { ToolEvent } from "../agent/loop";
import { TOOL_BY_NAME } from "../excel/toolSchemas";
import type { ChatMessage } from "./api/client";

export type PersistedEntry =
  | { kind: "user"; text: string }
  | { kind: "assistant"; text: string }
  | { kind: "error"; text: string }
  /** Остановка пользователем. Отдельный вид, потому что это не ошибка и не
   * ответ модели, а её отсутствие: без такой записи остановка не оставляет
   * в ленте следа и неотличима от сбоя или ненажатой кнопки. */
  | { kind: "notice"; text: string }
  | { kind: "op"; event: ToolEvent };

export interface StoredConversation {
  version: 1;
  workbookKey: string;
  documentUrl: string;
  title: string;
  updatedAt: number;
  entries: PersistedEntry[];
  history: ChatMessage[];
}

const STORAGE_KEY = "excel-ai-addin.conversations.v1";
export const CONVERSATION_RETENTION_MS = 30 * 24 * 60 * 60_000;
export const MAX_CONVERSATION_STORAGE_BYTES = 2 * 1024 * 1024;
const MAX_CONVERSATIONS = 20;

const byteLength = (value: string) => new TextEncoder().encode(value).length;

function hash(value: string): string {
  let result = 0x811c9dc5;
  for (let index = 0; index < value.length; index++) {
    result ^= value.charCodeAt(index);
    result = Math.imul(result, 0x01000193);
  }
  return (result >>> 0).toString(16).padStart(8, "0");
}

const normalizeUrl = (documentUrl: string) => documentUrl.trim().replace(/\\/g, "/").toLocaleLowerCase();
const URL_KEY = "saved:url:";

/**
 * Ключ беседы — полный нормализованный адрес книги.
 *
 * Прежде ключом был 32-битный отпечаток адреса, и у разных книг он мог
 * совпасть: тогда одна книга получала историю и разрешение записи другой
 * (аудит 24 сентября 2026 года, SEC-05; пара совпадающих адресов — в тестах).
 */
export function conversationIdentity(documentUrl: string): string | null {
  const normalized = normalizeUrl(documentUrl);
  return normalized ? `${URL_KEY}${normalized}` : null;
}

function interruptedToolMessage(call: { id: string; name: string }): ChatMessage {
  const mutating = TOOL_BY_NAME.get(call.name)?.mutating === true;
  return {
    role: "tool",
    tool_call_id: call.id,
    content: JSON.stringify(mutating
      ? {
          ok: false,
          error: "Предыдущая сессия прервалась до фиксации результата записи. Не повторяйте её автоматически; сначала перечитайте цель.",
          executionState: "unknown"
        }
      : {
          ok: false,
          error: "Чтение прервалось при закрытии предыдущей сессии.",
          executionState: "not_started"
        })
  };
}

/** Removes private reasoning and repairs every assistant tool-call pair. */
export function repairConversationHistory(input: ChatMessage[]): ChatMessage[] {
  const repaired: ChatMessage[] = [];
  for (let index = 0; index < input.length; index++) {
    const message = input[index];
    if (message.role === "system" || message.role === "tool") continue;
    if (message.role !== "assistant" || !message.tool_calls?.length) {
      repaired.push(message.role === "assistant"
        ? { role: "assistant", content: message.content, provider: message.provider }
        : { ...message });
      continue;
    }

    repaired.push({
      role: "assistant",
      content: message.content,
      provider: message.provider,
      tool_calls: message.tool_calls.map((call) => ({ ...call }))
    });
    const responses = new Map<string, ChatMessage & { role: "tool" }>();
    let next = index + 1;
    while (next < input.length && input[next].role === "tool") {
      const response = input[next] as ChatMessage & { role: "tool" };
      if (!responses.has(response.tool_call_id)) responses.set(response.tool_call_id, { ...response });
      next += 1;
    }
    for (const call of message.tool_calls) repaired.push(responses.get(call.id) ?? interruptedToolMessage(call));
    index = next - 1;
  }
  return repaired;
}

function readAll(storage: Storage, now: number): StoredConversation[] {
  try {
    const parsed = JSON.parse(storage.getItem(STORAGE_KEY) ?? "[]");
    if (!Array.isArray(parsed)) return [];
    return parsed.filter((item): item is StoredConversation =>
      item?.version === 1 && typeof item.workbookKey === "string" &&
      typeof item.updatedAt === "number" && now - item.updatedAt <= CONVERSATION_RETENTION_MS &&
      Array.isArray(item.entries) && Array.isArray(item.history));
  } catch {
    return [];
  }
}

function writeBounded(storage: Storage, conversations: StoredConversation[]): boolean {
  const kept = conversations.sort((left, right) => right.updatedAt - left.updatedAt).slice(0, MAX_CONVERSATIONS);
  let serialized = JSON.stringify(kept);
  while (kept.length > 1 && byteLength(serialized) > MAX_CONVERSATION_STORAGE_BYTES) {
    kept.pop();
    serialized = JSON.stringify(kept);
  }
  if (byteLength(serialized) > MAX_CONVERSATION_STORAGE_BYTES) return false;
  try {
    storage.setItem(STORAGE_KEY, serialized);
    return true;
  } catch {
    return false;
  }
}

export function loadConversation(storage: Storage, workbookKey: string, now = Date.now()): StoredConversation | null {
  const all = readAll(storage, now);
  writeBounded(storage, all);
  let found = all.find((item) => item.workbookKey === workbookKey);
  // Беседа, сохранённая до исправления под коротким отпечатком, переносится
  // только если рядом сохранён тот же полный адрес: совпадения отпечатка мало.
  if (!found && workbookKey.startsWith(URL_KEY)) {
    const normalized = workbookKey.slice(URL_KEY.length);
    const legacyKey = `saved:${hash(normalized)}`;
    const legacy = all.find((item) => item.workbookKey === legacyKey && normalizeUrl(String(item.documentUrl ?? "")) === normalized);
    if (legacy) found = { ...legacy, workbookKey };
  }
  return found ?{ ...found, entries: [...found.entries], history: repairConversationHistory(found.history) } : null;
}

export function saveConversation(
  storage: Storage,
  value: Omit<StoredConversation, "version" | "updatedAt" | "history"> & { history: ChatMessage[] },
  now = Date.now()
): boolean {
  const history = repairConversationHistory(value.history);
  // Вместе с записью этой книги уходит и её прежняя копия под коротким ключом.
  const legacyKey = value.workbookKey.startsWith(URL_KEY) ? `saved:${hash(value.workbookKey.slice(URL_KEY.length))}` : null;
  const all = readAll(storage, now).filter((item) =>
    item.workbookKey !== value.workbookKey &&
    !(item.workbookKey === legacyKey && normalizeUrl(String(item.documentUrl ?? "")) === value.workbookKey.slice(URL_KEY.length)));
  if (!value.entries.length && !history.length) return writeBounded(storage, all);
  let conversation: StoredConversation = { ...value, version: 1, updatedAt: now, history };
  while (byteLength(JSON.stringify(conversation)) > MAX_CONVERSATION_STORAGE_BYTES) {
    const nextHistoryUser = conversation.history.findIndex((message, index) => index > 0 && message.role === "user");
    const nextEntryUser = conversation.entries.findIndex((entry, index) => index > 0 && entry.kind === "user");
    if (nextHistoryUser < 0 || nextEntryUser < 0) return false;
    const entries = conversation.entries.slice(nextEntryUser);
    conversation = {
      ...conversation,
      title: entries.find((entry) => entry.kind === "user")?.text.slice(0, 80) || conversation.title,
      entries,
      history: repairConversationHistory(conversation.history.slice(nextHistoryUser))
    };
  }
  return writeBounded(storage, [conversation, ...all]);
}

export function deleteConversation(storage: Storage, workbookKey: string, now = Date.now()) {
  writeBounded(storage, readAll(storage, now).filter((item) => item.workbookKey !== workbookKey));
}
