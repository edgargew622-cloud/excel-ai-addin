import test from "node:test";
import assert from "node:assert/strict";
import { cachesByMarkers, withPromptCache } from "./promptCache.js";

const conversation = [
  { role: "system", content: "Правила агента" },
  { role: "system", content: "Контекст задачи" },
  { role: "user", content: "Сделай копию" },
  { role: "assistant", content: "", tool_calls: [{ id: "c1", type: "function", function: { name: "x", arguments: "{}" } }] },
  { role: "tool", tool_call_id: "c1", content: "{\"ok\":true}" }
];

test("Claude gets cache markers on the agent rules and on the end of the conversation", () => {
  const out = withPromptCache(conversation, "openrouter", "anthropic/claude-sonnet-5.5") as any[];
  assert.deepEqual(out[0].content, [{ type: "text", text: "Правила агента", cache_control: { type: "ephemeral" } }]);
  assert.equal(out[1].content, "Контекст задачи", "контекст задачи меняется — не помечается");
  assert.deepEqual(out[4].content, [{ type: "text", text: "{\"ok\":true}", cache_control: { type: "ephemeral" } }]);
  assert.equal(out[4].tool_call_id, "c1");
  assert.equal((conversation[0] as any).content, "Правила агента", "исходные сообщения не меняются");
});

test("other models are left as they are: their providers cache by themselves", () => {
  for (const [provider, model] of [["deepseek", "deepseek-flash"], ["openai", "gpt-6.1-sol"], ["openrouter", "google/gemini-3.8-flash"]]) {
    assert.equal(cachesByMarkers(provider, model), false);
    assert.deepEqual(withPromptCache(conversation, provider, model), conversation);
  }
});
