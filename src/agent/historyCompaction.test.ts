import test from "node:test";
import assert from "node:assert/strict";
import { compactOldResults, fitRequest } from "./historyCompaction";
import type { ChatMessage } from "../taskpane/api/client";

const big = (label: string) => `${label}:` + "x".repeat(20_000);
/** Задача: просьба, вызов инструмента, его большой результат, ответ. */
function turn(n: number): ChatMessage[] {
  return [
    { role: "user", content: `просьба ${n}` },
    { role: "assistant", content: "", tool_calls: [{ id: `c${n}`, type: "function", function: { name: "read_web_page", arguments: "{}" } }] as any },
    { role: "tool", tool_call_id: `c${n}`, content: big(`страница ${n}`) },
    { role: "assistant", content: `ответ ${n}` }
  ];
}

test("old tool results are cut to a short head with a note; the last two tasks stay whole", () => {
  const history = [...turn(1), ...turn(2), ...turn(3), ...turn(4)];
  assert.equal(compactOldResults(history), 2);
  const tools = history.filter((m) => m.role === "tool") as { content: string }[];
  assert.ok(tools[0].content.length < 800 && /сокращён панелью/.test(tools[0].content));
  assert.ok(tools[1].content.length < 800);
  assert.equal(tools[2].content.length, big("страница 3").length);
  assert.equal(tools[3].content.length, big("страница 4").length);
  // Повторно сжимать нечего.
  assert.equal(compactOldResults(history), 0);
});

test("a request over the limit drops the earliest tasks whole, keeps tool pairs and tells the model", () => {
  const messages: ChatMessage[] = [{ role: "system", content: "правила" }, ...turn(1), ...turn(2), ...turn(3)];
  const fits = (list: ChatMessage[]) => JSON.stringify(list).length < 50_000;
  const fitted = fitRequest(messages, fits);
  assert.equal(fitted.droppedTurns, 1);
  assert.ok(fits(fitted.messages));
  assert.equal(fitted.messages[0].content, "правила");
  assert.match(String(fitted.messages[1].content), /Ранние задачи этой беседы \(1\) опущены/);
  assert.equal((fitted.messages[2] as any).content, "просьба 2");
  // Каждому вызову — свой ответ инструмента.
  const calls = fitted.messages.flatMap((m: any) => m.tool_calls?.map((c: any) => c.id) ?? []);
  const answers = fitted.messages.filter((m) => m.role === "tool").map((m: any) => m.tool_call_id);
  assert.deepEqual(calls, answers);
  assert.equal(messages.length, 13, "исходный список не меняется");
});

test("when the current task alone is too big its early results are shortened, the last four stay", () => {
  const current: ChatMessage[] = [{ role: "user", content: "одна большая задача" }];
  for (let i = 0; i < 8; i++) {
    current.push({ role: "assistant", content: "", tool_calls: [{ id: `k${i}`, type: "function", function: { name: "get_range_values", arguments: "{}" } }] as any });
    current.push({ role: "tool", tool_call_id: `k${i}`, content: big(`часть ${i}`) });
  }
  const fitted = fitRequest(current, (list) => JSON.stringify(list).length < 110_000);
  assert.ok(fitted.shortenedCurrent > 0);
  const tools = fitted.messages.filter((m) => m.role === "tool") as { content: string }[];
  assert.ok(tools.slice(-4).every((m) => m.content.length > 20_000));
});
