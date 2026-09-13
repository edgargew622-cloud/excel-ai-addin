import test from "node:test";
import assert from "node:assert/strict";
import {
  byteLength,
  createMetrics,
  formatBytes,
  formatMetricLog,
  formatMs,
  summarize,
  type OpMetric
} from "./metrics";

const metric = (over: Partial<OpMetric> = {}): OpMetric => ({
  kind: "tool",
  name: "get_range_values",
  ms: 100,
  bytes: 500,
  ok: true,
  ...over
});

test("summary separates model requests from tool calls", () => {
  const s = summarize([
    metric({ kind: "model", name: "deepseek/deepseek-flash", ms: 2000, bytes: 3000 }),
    metric({ ms: 50, bytes: 100 }),
    metric({ ms: 300, bytes: 90_000 })
  ]);
  assert.equal(s.count, 3);
  assert.equal(s.modelCount, 1);
  assert.equal(s.toolCount, 2);
  assert.equal(s.totalMs, 2350);
  assert.equal(s.totalBytes, 93_100);
  assert.equal(s.slowest?.ms, 2000);
  assert.equal(s.largest?.bytes, 90_000);
});

test("empty summary has no slowest or largest", () => {
  const s = summarize([]);
  assert.equal(s.count, 0);
  assert.equal(s.slowest, null);
  assert.equal(s.largest, null);
});

test("store keeps a bounded window", () => {
  const store = createMetrics(3);
  for (let i = 0; i < 10; i++) store.record(metric({ ms: i }));
  assert.equal(store.all().length, 3);
  // Остаются последние, а не первые.
  assert.deepEqual(store.all().map((m) => m.ms), [7, 8, 9]);
  store.reset();
  assert.equal(store.all().length, 0);
});

test("byte length counts UTF-8 bytes, not characters", () => {
  // Кириллица в UTF-8 занимает два байта: считать символы значило бы
  // недооценивать объём русских книг вдвое.
  assert.equal(byteLength("abc"), 3);
  assert.equal(byteLength("абв"), 6);
});

test("metric log line cannot carry workbook content", () => {
  const line = formatMetricLog(metric({ name: "set_range_values", ms: 12.7, bytes: 64 }));
  assert.equal(line, "kind=tool name=set_range_values ms=13 bytes=64 ok=true");
  // Структура метрики не имеет поля под значения ячеек, и это её главное
  // свойство для технического журнала.
  assert.equal(Object.keys(metric()).sort().join(","), "bytes,kind,ms,name,ok");
});

test("human formatting of sizes and durations", () => {
  assert.equal(formatBytes(512), "512 Б");
  assert.equal(formatBytes(2048), "2.0 КБ");
  assert.equal(formatBytes(3 * 1024 * 1024), "3.0 МБ");
  assert.equal(formatMs(450), "450 мс");
  assert.equal(formatMs(2500), "2.5 с");
});
