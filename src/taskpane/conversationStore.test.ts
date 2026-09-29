import test from "node:test";
import assert from "node:assert/strict";
import {
  MAX_CONVERSATION_STORAGE_BYTES,
  bindingDecision,
  conversationFileName,
  conversationIdentity,
  conversationMarkdown,
  deleteAllConversations,
  deleteConversation,
  listConversations,
  loadConversation,
  readConversation,
  repairConversationHistory,
  saveConversation
} from "./conversationStore";
import type { ChatMessage } from "./api/client";

class MemoryStorage implements Storage {
  private data = new Map<string, string>();
  get length() { return this.data.size; }
  clear() { this.data.clear(); }
  getItem(key: string) { return this.data.get(key) ?? null; }
  key(index: number) { return [...this.data.keys()][index] ?? null; }
  removeItem(key: string) { this.data.delete(key); }
  setItem(key: string, value: string) { this.data.set(key, value); }
}

test("saved workbook identity is stable while unsaved books are not auto-bound", () => {
  assert.equal(conversationIdentity(""), null);
  assert.equal(conversationIdentity("C:\\Books\\Plan.xlsx"), conversationIdentity("c:/books/plan.xlsx"));
});

test("history repair closes interrupted writes as unknown and strips reasoning", () => {
  const history: ChatMessage[] = [
    { role: "user", content: "write" },
    {
      role: "assistant",
      content: "",
      reasoning_content: "private",
      tool_calls: [{ id: "w1", name: "set_range_values", arguments: "{}" }]
    }
  ];
  const repaired = repairConversationHistory(history);
  assert.equal(repaired.length, 3);
  assert.equal("reasoning_content" in repaired[1], false);
  assert.match((repaired[2] as { content: string }).content, /"executionState":"unknown"/);
});

test("conversation is isolated by workbook, expires and can be deleted", () => {
  const storage = new MemoryStorage();
  const base = { documentUrl: "C:/Books/A.xlsx", title: "First", entries: [{ kind: "user" as const, text: "hello" }], history: [{ role: "user" as const, content: "hello" }] };
  assert.equal(saveConversation(storage, { ...base, workbookKey: "a" }, 1000), true);
  assert.equal(loadConversation(storage, "b", 1001), null);
  assert.equal(loadConversation(storage, "a", 1001)?.title, "First");
  deleteConversation(storage, "a", 1002);
  assert.equal(loadConversation(storage, "a", 1003), null);
});

test("oversized history drops complete old turns instead of leaving broken tool pairs", () => {
  const storage = new MemoryStorage();
  // Текст хранится и в видимой ленте, и в model history: два хода должны
  // превысить лимит, а один — гарантированно помещаться.
  const large = "я".repeat(Math.floor(MAX_CONVERSATION_STORAGE_BYTES / 7));
  const entries = [
    { kind: "user" as const, text: "old" },
    { kind: "assistant" as const, text: large },
    { kind: "user" as const, text: "new" },
    { kind: "assistant" as const, text: large }
  ];
  const history: ChatMessage[] = entries.map((entry) => ({
    role: entry.kind === "user" ? "user" as const : "assistant" as const,
    content: entry.text
  }));
  assert.equal(saveConversation(storage, {
    workbookKey: "large",
    documentUrl: "C:/large.xlsx",
    title: "old",
    entries,
    history
  }), true);
  const restored = loadConversation(storage, "large");
  assert.equal(restored?.entries[0].kind, "user");
  assert.equal(restored?.history[0].role, "user");
  assert.equal(restored?.title, "new");
});

// Аудит 24 сентября 2026 года (SEC-05): ключ истории был 32-битным отпечатком
// адреса книги. Эта пара разных книг даёт один и тот же отпечаток c95249db.
const COLLIDING = ["C:/Отчёты/Книга-1549599.xlsx", "C:/Отчёты/Книга-1712382.xlsx"];

test("two different workbooks never share a history, even when their short hashes collide", () => {
  const [first, second] = COLLIDING;
  assert.notEqual(conversationIdentity(first), conversationIdentity(second));
  const storage = new MemoryStorage();
  const key = conversationIdentity(first)!;
  saveConversation(storage, { workbookKey: key, documentUrl: first, title: "Секретная", entries: [{ kind: "user", text: "тайна" }], history: [{ role: "user", content: "тайна" }] }, 1000);
  assert.equal(loadConversation(storage, conversationIdentity(second)!, 1001), null);
  assert.equal(loadConversation(storage, key, 1001)?.title, "Секретная");
});

test("a history saved under the old short key is carried over only for the same workbook", () => {
  const [first, second] = COLLIDING;
  const storage = new MemoryStorage();
  // Так запись лежала до исправления: ключ — отпечаток, рядом — полный адрес.
  const legacy = [{ version: 1, workbookKey: "saved:c95249db", documentUrl: first, title: "Прежняя", updatedAt: 1000, entries: [{ kind: "user", text: "было" }], history: [{ role: "user", content: "было" }] }];
  storage.setItem("excel-ai-addin.conversations.v1", JSON.stringify(legacy));
  assert.equal(loadConversation(storage, conversationIdentity(second)!, 1001), null);
  assert.equal(loadConversation(storage, conversationIdentity(first)!, 1001)?.title, "Прежняя");
});

test("saving the workbook keeps the conversation: a new address inside an open panel is the same book", () => {
  const book1 = conversationIdentity("C:/Отчёты/Книга1.xlsx")!;
  const renamed = conversationIdentity("C:/Отчёты/Продажи.xlsx")!;
  // Панель только открылась — беседа ищется по адресу.
  assert.equal(bindingDecision(null, false, book1), "load");
  // Несохранённая книга впервые сохранена, пока панель открыта, — перенос.
  assert.equal(bindingDecision(null, true, book1), "migrate");
  // «Сохранить как» — тоже перенос.
  assert.equal(bindingDecision({ key: book1 }, true, renamed), "migrate");
  assert.equal(bindingDecision({ key: book1 }, true, book1), "keep");
});

const DAY = 24 * 60 * 60_000;

function seed(storage: Storage, key: string, url: string, title: string, at: number) {
  saveConversation(storage, {
    workbookKey: key,
    documentUrl: url,
    title,
    entries: [
      { kind: "user", text: title },
      { kind: "op", event: { id: "1", name: "set_range_values", args: { sheet: "Лист1", address: "A1:B2" }, status: "done" } },
      { kind: "assistant", text: "Готово." }
    ],
    history: [{ role: "user", content: title }]
  }, at);
}

test("the conversations window lists every workbook, newest first, with its expiry date", () => {
  const storage = new MemoryStorage();
  seed(storage, "doc:1", "C:/Отчёты/Продажи.xlsx", "сводная по городам", 1000);
  seed(storage, "doc:2", "https://d.docs.live.net/abc/Бюджет%202026.xlsx", "бюджет", 2000);
  seed(storage, "doc:3", "", "черновик", 3000);
  const overview = listConversations(storage, 3000);
  assert.deepEqual(overview.conversations.map((c) => c.workbookName), ["Книга без адреса", "Бюджет 2026.xlsx", "Продажи.xlsx"]);
  const sales = overview.conversations[2];
  assert.equal(sales.expiresAt, 1000 + 30 * DAY);
  assert.equal(sales.messages, 2, "сообщения пользователя и ответы агента");
  assert.equal(sales.actions, 1);
  assert.ok(overview.usedBytes > 0 && overview.usedBytes <= overview.limitBytes);
  assert.equal(overview.retentionDays, 30);
  assert.equal(overview.maxConversations, 20);
});

test("expired and damaged conversations are not listed, and a corrupt store does not break the window", () => {
  const storage = new MemoryStorage();
  seed(storage, "doc:old", "C:/a.xlsx", "старая", 1000);
  assert.equal(listConversations(storage, 1000 + 31 * DAY).conversations.length, 0);
  storage.setItem("excel-ai-addin.conversations.v1", "{не json");
  assert.deepEqual(listConversations(storage, 1000).conversations, []);
});

test("reading another workbook's conversation does not change the store", () => {
  const storage = new MemoryStorage();
  seed(storage, "doc:1", "C:/a.xlsx", "первая", 1000);
  const before = storage.getItem("excel-ai-addin.conversations.v1");
  assert.equal(readConversation(storage, "doc:1", 2000)?.title, "первая");
  assert.equal(readConversation(storage, "doc:нет", 2000), null);
  assert.equal(storage.getItem("excel-ai-addin.conversations.v1"), before);
});

test("one conversation or all of them can be deleted", () => {
  const storage = new MemoryStorage();
  seed(storage, "doc:1", "C:/a.xlsx", "первая", 1000);
  seed(storage, "doc:2", "C:/b.xlsx", "вторая", 2000);
  deleteConversation(storage, "doc:1", 2000);
  assert.deepEqual(listConversations(storage, 2000).conversations.map((c) => c.workbookKey), ["doc:2"]);
  deleteAllConversations(storage);
  assert.equal(listConversations(storage, 2000).conversations.length, 0);
});

test("a conversation becomes readable Markdown with the agent's actions and a safe file name", () => {
  const storage = new MemoryStorage();
  seed(storage, "doc:1", "/Users/ed/Отчёты/Продажи: итог.xlsx", "сводная по городам", new Date(2026, 8, 30, 14, 5).getTime());
  const conversation = readConversation(storage, "doc:1", new Date(2026, 8, 30, 15, 0).getTime())!;
  const text = conversationMarkdown(conversation);
  assert.match(text, /^# Беседа am\.AI — Продажи: итог\.xlsx/);
  assert.match(text, /## Вы\n\nсводная по городам/);
  assert.match(text, /\*\*Действия агента:\*\*\n- `set_range_values` Лист1!A1:B2 — выполнено/);
  assert.match(text, /## am\.AI\n\nГотово\./);
  assert.match(text, /Последнее сообщение: 30\.09\.2026 14:05/);
  assert.equal(conversationFileName(conversation), "Продажи_ итог 2026-09-30 14-05.md");
});
