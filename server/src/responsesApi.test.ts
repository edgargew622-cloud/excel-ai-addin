import test from "node:test";
import assert from "node:assert/strict";
import { ResponsesTranslator, buildResponsesBody, toResponsesInput, translateResponsesChunk } from "./responsesApi.js";
import type { InternalMessage } from "./protocol.js";

/** Собирает то, что увидела бы панель, из событий /v1/responses. */
function panelView(events: Record<string, unknown>[]) {
  const translator = new ResponsesTranslator();
  const pieces: string[] = [];
  for (const event of events) pieces.push(...translateResponsesChunk(translator, `data: ${JSON.stringify(event)}`));
  let content = "";
  let finishReason = "";
  const calls = new Map<number, { id: string; name: string; arguments: string }>();
  let done = false;
  for (const piece of pieces) {
    const payload = piece.slice(5).trim();
    if (payload === "[DONE]") { done = true; continue; }
    const chunk = JSON.parse(payload);
    const choice = chunk.choices?.[0];
    if (!choice) continue;
    if (choice.finish_reason) finishReason = choice.finish_reason;
    if (typeof choice.delta?.content === "string") content += choice.delta.content;
    for (const call of choice.delta?.tool_calls ?? []) {
      const acc = calls.get(call.index) ?? { id: "", name: "", arguments: "" };
      if (call.id) acc.id = call.id;
      if (call.function?.name) acc.name = call.function.name;
      if (typeof call.function?.arguments === "string") acc.arguments += call.function.arguments;
      calls.set(call.index, acc);
    }
  }
  return { content, finishReason, done, calls: [...calls.entries()].sort((a, b) => a[0] - b[0]).map(([, c]) => c) };
}

test("a text answer with a tool call becomes the shape the panel already reads", () => {
  // Форма событий снята с живого ответа gpt-6-astra.
  const view = panelView([
    { type: "response.created" },
    { type: "response.output_item.added", item: { id: "msg_1", type: "message", role: "assistant" }, output_index: 0 },
    { type: "response.output_text.delta", delta: "готов", item_id: "msg_1", output_index: 0 },
    { type: "response.output_text.delta", delta: "о", item_id: "msg_1", output_index: 0 },
    { type: "response.output_item.added", item: { id: "fc_1", type: "function_call", call_id: "call_abc", name: "ping", arguments: "" }, output_index: 1 },
    { type: "response.function_call_arguments.delta", delta: '{"a"', item_id: "fc_1", output_index: 1 },
    { type: "response.function_call_arguments.delta", delta: ":1}", item_id: "fc_1", output_index: 1 },
    { type: "response.completed" }
  ]);

  assert.equal(view.content, "готово");
  assert.deepEqual(view.calls, [{ id: "call_abc", name: "ping", arguments: '{"a":1}' }]);
  // Панель отвергает вызовы инструментов без этого finish_reason.
  assert.equal(view.finishReason, "tool_calls");
  assert.equal(view.done, true);
});

test("tool calls are numbered from zero regardless of their output_index", () => {
  const view = panelView([
    { type: "response.output_item.added", item: { id: "msg_1", type: "message", role: "assistant" }, output_index: 0 },
    { type: "response.output_item.added", item: { id: "fc_a", type: "function_call", call_id: "call_a", name: "first" }, output_index: 3 },
    { type: "response.output_item.added", item: { id: "fc_b", type: "function_call", call_id: "call_b", name: "second" }, output_index: 7 },
    { type: "response.function_call_arguments.delta", delta: "{}", item_id: "fc_b", output_index: 7 },
    { type: "response.function_call_arguments.delta", delta: "{}", item_id: "fc_a", output_index: 3 },
    { type: "response.completed" }
  ]);
  // Аргументы приходят вперемешку и должны попасть каждый в свой вызов.
  assert.deepEqual(view.calls, [
    { id: "call_a", name: "first", arguments: "{}" },
    { id: "call_b", name: "second", arguments: "{}" }
  ]);
});

test("an answer without tools finishes as stop", () => {
  const view = panelView([
    { type: "response.output_text.delta", delta: "ответ", item_id: "msg_1" },
    { type: "response.completed" }
  ]);
  assert.equal(view.finishReason, "stop");
  assert.equal(view.content, "ответ");
});

test("a truncated stream never looks like a finished step", () => {
  const translator = new ResponsesTranslator();
  const pieces = [
    ...translateResponsesChunk(translator, 'data: {"type":"response.output_text.delta","delta":"поло"}'),
    ...translator.finalizeIfUnfinished()
  ];
  const last = JSON.parse(pieces[pieces.length - 1].slice(5).trim());
  // Ни finish_reason, ни [DONE]: панель обязана отвергнуть такой ответ.
  assert.ok(last.error?.message, "оборванный поток сообщает об ошибке");
  assert.equal(pieces.some((piece) => piece.includes("finish_reason")), false);
  assert.equal(pieces.some((piece) => piece.includes("[DONE]")), false);
});

test("hitting the token limit is not reported as a clean stop", () => {
  const view = panelView([
    { type: "response.output_text.delta", delta: "начало", item_id: "msg_1" },
    { type: "response.incomplete" }
  ]);
  assert.equal(view.finishReason, "length");
});

test("tool calls and their results are linked by call_id in the request", () => {
  const messages: InternalMessage[] = [
    { role: "system", content: "правила" },
    { role: "user", content: "посчитай" },
    { role: "assistant", content: "считаю", tool_calls: [{ id: "call_1", name: "get_range_values", arguments: '{"address":"A1"}' }] },
    { role: "tool", tool_call_id: "call_1", content: '{"ok":true}' },
    { role: "assistant", content: "готово" }
  ];
  assert.deepEqual(toResponsesInput(messages), [
    { role: "system", content: "правила" },
    { role: "user", content: "посчитай" },
    { role: "assistant", content: "считаю" },
    { type: "function_call", call_id: "call_1", name: "get_range_values", arguments: '{"address":"A1"}' },
    { type: "function_call_output", call_id: "call_1", output: '{"ok":true}' },
    { role: "assistant", content: "готово" }
  ]);
});

test("tools are described flat, as the responses API expects", () => {
  const body = buildResponsesBody("gpt-6-astra", [{ role: "user", content: "привет" }], [
    { type: "function", function: { name: "ping", description: "тест", parameters: { type: "object" } } }
  ]);
  assert.deepEqual(body.tools, [{ type: "function", name: "ping", description: "тест", parameters: { type: "object" } }]);
  assert.equal(body.stream, true);
  assert.equal(body.model, "gpt-6-astra");
});
