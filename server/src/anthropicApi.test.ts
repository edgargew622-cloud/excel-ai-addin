import test from "node:test";
import assert from "node:assert/strict";
import { AnthropicTranslator, buildAnthropicParams, toolUseId } from "./anthropicApi.js";
import type { InternalMessage } from "./protocol.js";

const tool = { type: "function" as const, function: { name: "get_range_values", description: "Прочитать", parameters: { type: "object", properties: { address: { type: "string" } } } } };

test("system messages go to the system field, the rules block is cached", () => {
  const messages: InternalMessage[] = [
    { role: "system", content: "правила" },
    { role: "system", content: "контекст задачи" },
    { role: "user", content: "привет" }
  ];
  const params = buildAnthropicParams("claude-opus-5-5", messages, [tool]);
  assert.equal(params.system.length, 2);
  assert.deepEqual(params.system[0], { type: "text", text: "правила", cache_control: { type: "ephemeral" } });
  assert.equal(params.system[1].text, "контекст задачи");
  assert.deepEqual(params.messages, [{ role: "user", content: "привет" }]);
  assert.equal(params.tools[0].name, "get_range_values");
  assert.equal((params as any).thinking.block_binding.prefix_mismatch_behavior, "drop_block");
  assert.equal((params as any).fallbacks, "default");
  assert.equal((buildAnthropicParams("claude-haiku-5-5", messages, []) as any).fallbacks, undefined, "у Haiku резерва нет");
  const legacy = buildAnthropicParams("claude-haiku-4-5", messages, []) as any;
  assert.equal(legacy.thinking, undefined, "Haiku 4.5 не знает адаптивных размышлений");
  assert.equal(legacy.output_config, undefined, "и effort");
  assert.equal((buildAnthropicParams("claude-sonnet-4-6", messages, []) as any).thinking.type, "adaptive");
});

test("tool calls become tool_use, results are grouped into one user message, thinking goes first", () => {
  const messages: InternalMessage[] = [
    { role: "user", content: "прочитай" },
    {
      role: "assistant", content: "Читаю.", provider: "anthropic",
      thinking_blocks: [{ type: "thinking", thinking: "", signature: "sig" }],
      tool_calls: [{ id: "toolu_1", name: "get_range_values", arguments: "{\"address\":\"A1\"}" }, { id: "call:2", name: "get_range_values", arguments: "не json" }]
    },
    { role: "tool", tool_call_id: "toolu_1", content: "{\"ok\":true}" },
    { role: "tool", tool_call_id: "call:2", content: "" }
  ];
  const params = buildAnthropicParams("claude-sonnet-5-5", messages, [tool]);
  const assistant = params.messages[1] as any;
  assert.deepEqual(assistant.content.map((b: any) => b.type), ["thinking", "text", "tool_use", "tool_use"]);
  assert.deepEqual(assistant.content[2].input, { address: "A1" });
  assert.deepEqual(assistant.content[3].input, {}, "неразборчивые аргументы — пустой объект, а не отказ всего запроса");
  assert.equal(assistant.content[3].id, "call_2");
  const results = params.messages[2] as any;
  assert.equal(results.role, "user");
  assert.deepEqual(results.content.map((b: any) => b.tool_use_id), ["toolu_1", "call_2"]);
  assert.equal(results.content[1].content, "(пусто)");
});

test("thinking of another provider is not sent to Claude", () => {
  const messages: InternalMessage[] = [
    { role: "user", content: "x" },
    { role: "assistant", content: "ответ", provider: "deepseek", reasoning_content: "мысли", thinking_blocks: [{ type: "thinking" }] } as any
  ];
  const params = buildAnthropicParams("claude-opus-5-5", messages, []);
  assert.deepEqual((params.messages[1] as any).content, [{ type: "text", text: "ответ" }]);
  assert.equal(toolUseId("call_00_ab-C"), "call_00_ab-C");
});

const parse = (pieces: string[]) => pieces.filter((p) => p.startsWith("data: {")).map((p) => JSON.parse(p.slice(6)));

test("the stream becomes Chat Completions chunks with tool calls, thinking and usage", () => {
  const t = new AnthropicTranslator();
  const out = [
    ...t.event({ type: "message_start", message: { usage: { input_tokens: 100, cache_read_input_tokens: 900, cache_creation_input_tokens: 50, output_tokens: 1 } } }),
    ...t.event({ type: "content_block_start", index: 0, content_block: { type: "thinking", thinking: "", signature: "" } }),
    ...t.event({ type: "content_block_delta", index: 0, delta: { type: "signature_delta", signature: "SIG" } }),
    ...t.event({ type: "content_block_stop", index: 0 }),
    ...t.event({ type: "content_block_start", index: 1, content_block: { type: "text", text: "" } }),
    ...t.event({ type: "content_block_delta", index: 1, delta: { type: "text_delta", text: "Читаю" } }),
    ...t.event({ type: "content_block_start", index: 2, content_block: { type: "tool_use", id: "toolu_9", name: "get_range_values", input: {} } }),
    ...t.event({ type: "content_block_delta", index: 2, delta: { type: "input_json_delta", partial_json: "{\"address\":" } }),
    ...t.event({ type: "content_block_delta", index: 2, delta: { type: "input_json_delta", partial_json: "\"A1\"}" } }),
    ...t.event({ type: "message_delta", delta: { stop_reason: "tool_use" }, usage: { output_tokens: 40 } }),
    ...t.event({ type: "message_stop" })
  ];
  const chunks = parse(out);
  assert.deepEqual(chunks.find((c) => c.choices?.[0]?.delta?.thinking_block)?.choices[0].delta.thinking_block, { type: "thinking", thinking: "", signature: "SIG" });
  assert.equal(chunks.find((c) => c.choices?.[0]?.delta?.content)?.choices[0].delta.content, "Читаю");
  const calls = chunks.flatMap((c) => c.choices?.[0]?.delta?.tool_calls ?? []);
  assert.equal(calls[0].id, "toolu_9");
  assert.equal(calls.map((c: any) => c.function.arguments).join(""), "{\"address\":\"A1\"}");
  assert.equal(chunks.find((c) => c.choices?.[0]?.finish_reason)?.choices[0].finish_reason, "tool_calls");
  assert.deepEqual(chunks.find((c) => c.usage)?.usage, { prompt_tokens: 1050, completion_tokens: 40, total_tokens: 1090, prompt_tokens_details: { cached_tokens: 900 } });
  assert.equal(out.at(-1), "data: [DONE]\n\n");
  assert.equal(t.done, true);
});

test("a refusal is explained to the user and ends the step", () => {
  const t = new AnthropicTranslator();
  t.event({ type: "message_delta", delta: { stop_reason: "refusal", stop_details: { category: "cyber" } }, usage: { output_tokens: 0 } });
  const chunks = parse(t.event({ type: "message_stop" }));
  assert.match(chunks[0].choices[0].delta.content, /отказался.*cyber/);
  assert.equal(chunks[1].choices[0].finish_reason, "stop");
});
