import test from "node:test";
import assert from "node:assert/strict";
import { serializeMessages, type InternalMessage } from "./protocol.js";

test("serializes internal tool_calls to OpenAI wire format", () => {
  const input: InternalMessage[] = [
    {
      role: "assistant",
      content: "",
      provider: "openai",
      tool_calls: [{ id: "call_1", name: "get_range_values", arguments: '{"sheet":"Лист1","address":"A1:B2"}' }]
    }
  ];
  assert.deepEqual(serializeMessages(input, "openai"), [
    {
      role: "assistant",
      content: "",
      tool_calls: [
        {
          id: "call_1",
          type: "function",
          function: { name: "get_range_values", arguments: '{"sheet":"Лист1","address":"A1:B2"}' }
        }
      ]
    }
  ]);
});

test("reasoning_content is replayed only to DeepSeek", () => {
  const input: InternalMessage[] = [
    { role: "assistant", content: "", provider: "deepseek", reasoning_content: "reason", tool_calls: [] }
  ];
  assert.equal((serializeMessages(input, "deepseek")[0] as any).reasoning_content, "reason");
  assert.equal("reasoning_content" in (serializeMessages(input, "openai")[0] as any), false);
});
