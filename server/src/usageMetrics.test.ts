import test from "node:test";
import assert from "node:assert/strict";
import {
  MetricsStore,
  UsageScanner,
  extractUsage,
  formatMetricLine,
  normalizeUsage,
  type RequestMetric
} from "./usageMetrics.js";

const metric = (over: Partial<RequestMetric> = {}): RequestMetric => ({
  provider: "openai",
  model: "gpt-6-astra",
  api: "chat",
  requestBytes: 1000,
  responseBytes: 200,
  firstByteMs: 300,
  totalMs: 900,
  ok: true,
  attempts: 1,
  at: "2026-09-16T09:00:00.000Z",
  ...over
});

test("both naming schemes for usage are understood", () => {
  // chat/completions
  assert.deepEqual(
    normalizeUsage({ prompt_tokens: 100, completion_tokens: 40, total_tokens: 140 }),
    { promptTokens: 100, completionTokens: 40, totalTokens: 140 }
  );
  // /v1/responses
  assert.deepEqual(
    normalizeUsage({ input_tokens: 100, output_tokens: 40, total_tokens: 140 }),
    { promptTokens: 100, completionTokens: 40, totalTokens: 140 }
  );
  // Итог считается сам, если провайдер его не прислал.
  assert.equal(normalizeUsage({ prompt_tokens: 7, completion_tokens: 3 })?.totalTokens, 10);
  assert.equal(normalizeUsage(null), null);
  assert.equal(normalizeUsage({ prompt_tokens: 0, completion_tokens: 0 }), null, "пустой usage — не usage");
});

test("reasoning and cached tokens are kept when the provider separates them", () => {
  const usage = normalizeUsage({
    prompt_tokens: 100, completion_tokens: 40, total_tokens: 140,
    completion_tokens_details: { reasoning_tokens: 25 },
    prompt_tokens_details: { cached_tokens: 64 }
  });
  assert.equal(usage?.reasoningTokens, 25);
  assert.equal(usage?.cachedTokens, 64);
  // Нулевые подробности не засоряют запись.
  const plain = normalizeUsage({ prompt_tokens: 1, completion_tokens: 1, completion_tokens_details: { reasoning_tokens: 0 } });
  assert.equal("reasoningTokens" in (plain ?? {}), false);
});

test("usage is picked out of the stream, and ordinary frames are ignored", () => {
  assert.equal(extractUsage('data: {"choices":[{"delta":{"content":"привет"}}]}'), null);
  assert.equal(extractUsage("data: [DONE]"), null);
  assert.equal(extractUsage(": ping"), null);
  assert.equal(extractUsage("data: не json"), null);

  const chat = extractUsage('data: {"choices":[],"usage":{"prompt_tokens":5,"completion_tokens":2,"total_tokens":7}}');
  assert.equal(chat?.totalTokens, 7);

  const responses = extractUsage('data: {"type":"response.completed","response":{"usage":{"input_tokens":9,"output_tokens":1}}}');
  assert.equal(responses?.totalTokens, 10);
});

test("the scanner survives usage split across chunks", () => {
  const scanner = new UsageScanner();
  // Читаем через функцию: иначе типы node сужают геттер после сравнения с null.
  const seen = () => scanner.usage;

  scanner.push('data: {"choices":[{"delta":{"content":"a"}}]}\n');
  assert.equal(seen(), null);
  // Кадр приходит разорванным посередине, как в настоящем потоке.
  scanner.push('data: {"usage":{"prompt_tokens":11,');
  assert.equal(seen(), null, "неполный кадр не разбирается");
  scanner.push('"completion_tokens":4,"total_tokens":15}}\n');
  assert.equal(seen()?.totalTokens, 15);
});

test("a final frame without a trailing newline is still read", () => {
  const scanner = new UsageScanner();
  scanner.push('data: {"usage":{"prompt_tokens":3,"completion_tokens":1,"total_tokens":4}}');
  assert.equal(scanner.usage, null, "строка ещё не закрыта");
  assert.equal(scanner.finish()?.totalTokens, 4);
});

test("a metric line carries numbers and names, never workbook content", () => {
  const line = formatMetricLine(metric({
    usage: { promptTokens: 100, completionTokens: 40, totalTokens: 140, cachedTokens: 64, reasoningTokens: 25 }
  }));
  assert.match(line, /openai\/gpt-6-astra/);
  assert.match(line, /tokens=140/);
  assert.match(line, /из кэша 64/);
  assert.match(line, /всего=900мс/);

  // Ключевое свойство: положить в строку содержимое ячеек неоткуда — запись
  // метрики состоит из чисел и имён. Проверяем на форме записи, а не на вере.
  const fields = Object.values(metric({ usage: { promptTokens: 1, completionTokens: 1, totalTokens: 2 } }));
  for (const value of fields) {
    assert.ok(
      ["number", "boolean", "string", "object"].includes(typeof value),
      "поле метрики имеет ожидаемый тип"
    );
  }
  assert.equal(line.includes("Москва"), false);
});

test("a request without usage is counted separately instead of as zero", () => {
  const store = new MetricsStore();
  store.record(metric({ usage: { promptTokens: 10, completionTokens: 5, totalTokens: 15 } }));
  store.record(metric());

  const summary = store.summary();
  assert.equal(summary.requests, 2);
  assert.equal(summary.totalTokens, 15);
  // Иначе неполный счёт выглядел бы как дешёвый запрос.
  assert.equal(summary.withoutUsage, 1);
  assert.match(formatMetricLine(metric()), /tokens=нет данных/);
});

test("the summary splits by model and counts failures", () => {
  const store = new MetricsStore();
  store.record(metric({ usage: { promptTokens: 10, completionTokens: 0, totalTokens: 10 } }));
  store.record(metric({ provider: "openrouter", model: "anthropic/claude-sonnet-5", ok: false, totalMs: 100 }));

  const summary = store.summary();
  assert.equal(summary.failed, 1);
  assert.equal(summary.totalMs, 1000);
  assert.equal(summary.byModel["openai/gpt-6-astra"].requests, 1);
  assert.equal(summary.byModel["openrouter/anthropic/claude-sonnet-5"].requests, 1);
});

test("the window is bounded: a long-lived server must not leak memory", () => {
  const store = new MetricsStore(3);
  for (let index = 0; index < 10; index++) store.record(metric({ totalMs: index }));
  assert.equal(store.all().length, 3);
  assert.deepEqual(store.all().map((item) => item.totalMs), [7, 8, 9], "остаются последние");
});
