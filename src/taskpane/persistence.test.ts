import test from "node:test";
import assert from "node:assert/strict";
import {
  conversationKey,
  repairHistory,
  trimToLimit,
  STORAGE_PREFIX,
  type StoredConversation
} from "./persistence";
import type { ChatMessage } from "./api/client";

test("history interrupted mid tool call is closed so the provider accepts it", () => {
  const history: ChatMessage[] = [
    { role: "user", content: "посчитай итоги" },
    {
      role: "assistant",
      content: "",
      tool_calls: [
        { id: "a", name: "get_range_values", arguments: "{}" },
        { id: "b", name: "set_range_values", arguments: "{}" }
      ]
    },
    { role: "tool", tool_call_id: "a", content: '{"ok":true}' }
  ];

  const { history: fixed, repaired } = repairHistory(history);
  assert.equal(repaired, 1);
  assert.equal(fixed.length, 4);
  const last = fixed[3];
  assert.equal(last.role, "tool");
  assert.equal(last.role === "tool" && last.tool_call_id, "b");
  assert.match(String(last.role === "tool" && last.content), /прервана/i);
});

test("a trailing assistant turn with tool calls gets every reply", () => {
  const history: ChatMessage[] = [
    { role: "user", content: "удали строки" },
    {
      role: "assistant",
      content: "",
      tool_calls: [{ id: "x", name: "delete_rows", arguments: "{}" }]
    }
  ];
  const { history: fixed, repaired } = repairHistory(history);
  assert.equal(repaired, 1);
  assert.equal(fixed.length, 3);
  assert.equal(fixed[2].role, "tool");
});

test("orphan tool replies are dropped", () => {
  const history: ChatMessage[] = [
    { role: "user", content: "привет" },
    { role: "tool", tool_call_id: "ghost", content: '{"ok":true}' },
    { role: "assistant", content: "здравствуйте" }
  ];
  const { history: fixed, repaired } = repairHistory(history);
  assert.equal(repaired, 1);
  assert.equal(fixed.length, 2);
  assert.equal(fixed.some((m) => m.role === "tool"), false);
});

test("a complete history is left untouched", () => {
  const history: ChatMessage[] = [
    { role: "user", content: "посчитай" },
    { role: "assistant", content: "", tool_calls: [{ id: "a", name: "get_range_values", arguments: "{}" }] },
    { role: "tool", tool_call_id: "a", content: '{"ok":true}' },
    { role: "assistant", content: "готово" }
  ];
  const { history: fixed, repaired } = repairHistory(history);
  assert.equal(repaired, 0);
  assert.deepEqual(fixed, history);
});

test("repair is idempotent", () => {
  const history: ChatMessage[] = [
    { role: "user", content: "u" },
    { role: "assistant", content: "", tool_calls: [{ id: "a", name: "t", arguments: "{}" }] }
  ];
  const once = repairHistory(history).history;
  const twice = repairHistory(once);
  assert.equal(twice.repaired, 0);
  assert.deepEqual(twice.history, once);
});

test("different workbooks get different conversation keys", () => {
  const a = conversationKey("https://example/книга-1.xlsx");
  const b = conversationKey("https://example/книга-2.xlsx");
  assert.notEqual(a, b);
  assert.equal(a.startsWith(STORAGE_PREFIX), true);
  // Неопределившаяся книга не должна получать пустой ключ.
  assert.equal(conversationKey(undefined), `${STORAGE_PREFIX}unknown-workbook`);
  assert.equal(conversationKey("   "), `${STORAGE_PREFIX}unknown-workbook`);
});

test("oversized conversation drops the oldest and stays valid", () => {
  const big = "д".repeat(5000);
  const payload: StoredConversation<{ text: string }> = {
    version: 1,
    workbook: "книга.xlsx",
    savedAt: 0,
    entries: Array.from({ length: 20 }, (_, i) => ({ text: `${i}` })),
    history: Array.from({ length: 20 }, (_, i) =>
      i % 2 === 0
        ? ({ role: "user", content: big } as ChatMessage)
        : ({ role: "assistant", content: big } as ChatMessage)
    )
  };

  const trimmed = trimToLimit(payload, 20_000);
  assert.equal(JSON.stringify(trimmed).length <= 20_000, true);
  assert.equal(trimmed.history.length < payload.history.length, true);
  // Срез не должен оставить ответ-сироту в начале истории.
  assert.equal(repairHistory(trimmed.history).repaired, 0);
});

test("trimming stops instead of emptying the conversation", () => {
  const payload: StoredConversation<{ text: string }> = {
    version: 1,
    workbook: "книга.xlsx",
    savedAt: 0,
    entries: [{ text: "один" }],
    history: [{ role: "user", content: "х".repeat(10_000) }]
  };
  const trimmed = trimToLimit(payload, 100);
  assert.equal(trimmed.history.length, 1);
});
