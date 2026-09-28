import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { MEMORY_LIMITS, MemoryError, MemoryStore } from "./memoryStore.js";

const fresh = () => join(mkdtempSync(join(tmpdir(), "memory-")), "memory.json");

test("preferences and scenarios are saved to disk and read back after a restart", async () => {
  const file = fresh();
  const store = new MemoryStore(file);
  store.load();
  await store.addPreference({ category: "numbers", text: "Суммы — с разделителем тысяч, без копеек" });
  await store.saveScenario({ name: "Месячный отчёт", steps: ["Убери лишние пробелы в A:A", "Построй сводную по городам"] });
  const again = new MemoryStore(file);
  again.load();
  const state = again.snapshot();
  assert.equal(state.preferences[0].text, "Суммы — с разделителем тысяч, без копеек");
  assert.deepEqual(state.scenarios[0].steps, ["Убери лишние пробелы в A:A", "Построй сводную по городам"]);
});

test("the same preference is not stored twice, a scenario with the same name is replaced", async () => {
  const store = new MemoryStore(fresh());
  await store.addPreference({ category: "headers", text: "Заголовки жирные" });
  await store.addPreference({ category: "headers", text: "заголовки жирные" });
  assert.equal(store.snapshot().preferences.length, 1);
  await store.saveScenario({ name: "Отчёт", steps: ["шаг 1"] });
  await store.saveScenario({ name: "отчёт", steps: ["шаг 1", "шаг 2"] });
  assert.equal(store.snapshot().scenarios.length, 1);
  assert.deepEqual(store.snapshot().scenarios[0].steps, ["шаг 1", "шаг 2"]);
});

test("limits and bad input are refused by the server itself", async () => {
  const store = new MemoryStore(fresh());
  await assert.rejects(() => store.addPreference({ category: "secrets", text: "x" }), MemoryError);
  await assert.rejects(() => store.addPreference({ category: "other", text: "  " }), /Пустое/);
  await assert.rejects(() => store.addPreference({ category: "other", text: "я".repeat(MEMORY_LIMITS.preferenceText + 1) }), /короче/);
  await assert.rejects(() => store.saveScenario({ name: "С", steps: [] }), /хотя бы один шаг/);
  await assert.rejects(() => store.saveScenario({ name: "С", steps: Array.from({ length: MEMORY_LIMITS.steps + 1 }, () => "шаг") }), /разбейте/);
  // Управляющие символы и переводы строк не проходят в текст, который увидит модель.
  const item = await store.addPreference({ category: "other", text: "строка\nвторая\u0007" });
  assert.equal(item.text, "строка вторая");
});

test("a broken memory file is not overwritten", async () => {
  const file = fresh();
  writeFileSync(file, "{ это не json", "utf8");
  const store = new MemoryStore(file);
  store.load();
  assert.match(String(store.loadError), /не прочиталась/);
  await assert.rejects(() => store.addPreference({ category: "other", text: "что-то" }), /испорчен/);
  assert.equal(readFileSync(file, "utf8"), "{ это не json");
});

test("removing an item that is not there is an error, not a silent success", async () => {
  const store = new MemoryStore(fresh());
  await assert.rejects(() => store.removePreference("нет-такого"), /нет/);
  await assert.rejects(() => store.removeScenario("нет-такого"), /нет/);
});
