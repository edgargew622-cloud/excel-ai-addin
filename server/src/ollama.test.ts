import test from "node:test";
import assert from "node:assert/strict";
import { detectOllamaModels, ollamaRoot, OllamaWatcher } from "./ollama.js";

function fakeOllama(models: Record<string, string[] | undefined>, calls: string[] = []) {
  return async (url: string, init?: RequestInit): Promise<Response> => {
    calls.push(url);
    if (url.endsWith("/api/tags")) {
      return new Response(JSON.stringify({ models: Object.keys(models).map((name) => ({ name })) }));
    }
    if (url.endsWith("/api/show")) {
      const name = JSON.parse(String(init?.body)).model;
      const capabilities = models[name];
      return new Response(JSON.stringify(capabilities ? { capabilities } : {}));
    }
    return new Response("", { status: 404 });
  };
}

test("root address is derived from the OpenAI-compatible one", () => {
  assert.equal(ollamaRoot("http://127.0.0.1:11434/v1"), "http://127.0.0.1:11434");
  assert.equal(ollamaRoot("http://127.0.0.1:11434/v1/"), "http://127.0.0.1:11434");
});

test("only models that can call tools are offered to the agent", async () => {
  const found = await detectOllamaModels("http://127.0.0.1:11434/v1", fakeOllama({
    "qwen3:8b": ["completion", "tools", "thinking"],
    "nomic-embed-text:latest": ["embedding"],
    "gemma3:4b": ["completion", "vision"],
    "old-model:1b": undefined
  }));
  assert.equal(found.running, true);
  assert.deepEqual(found.models, ["old-model:1b", "qwen3:8b"], "старая Ollama без возможностей — модель не отсеиваем");
});

test("no Ollama on this computer is not an error: the provider just is not listed", async () => {
  const found = await detectOllamaModels("http://127.0.0.1:11434/v1", async () => { throw new TypeError("fetch failed"); });
  assert.deepEqual(found, { models: [], running: false });
});

test("frequent panel requests do not hammer Ollama", async () => {
  const calls: string[] = [];
  let now = 1000;
  const seen: string[][] = [];
  const watcher = new OllamaWatcher(() => "http://127.0.0.1:11434/v1", (found) => seen.push(found.models),
    fakeOllama({ "qwen3:8b": ["tools"] }, calls), 5000, () => now);
  await Promise.all([watcher.refresh(), watcher.refresh()]);
  await watcher.refresh();
  assert.equal(calls.filter((url) => url.endsWith("/api/tags")).length, 1);
  now += 6000;
  await watcher.refresh();
  assert.equal(calls.filter((url) => url.endsWith("/api/tags")).length, 2);
  assert.deepEqual(seen.at(-1), ["qwen3:8b"]);
});
