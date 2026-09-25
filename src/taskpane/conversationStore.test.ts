import test from "node:test";
import assert from "node:assert/strict";
import {
  MAX_CONVERSATION_STORAGE_BYTES,
  conversationIdentity,
  deleteConversation,
  loadConversation,
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
