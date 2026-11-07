import test from "node:test";
import assert from "node:assert/strict";
import { appendDictation, SpeechInput, SPEECH_FALLBACK, type SpeechEngine, type SpeechPhase } from "./speechInput";

function setup(available = true) {
  const text: string[] = [], phases: SpeechPhase[] = [], messages: string[] = [];
  let stops = 0, aborts = 0, starts = 0;
  const engine: SpeechEngine = {
    lang: "", continuous: false, interimResults: false,
    onstart: null, onend: null, onerror: null, onresult: null,
    start() { starts++; }, stop() { stops++; }, abort() { aborts++; }
  };
  const input = new SpeechInput({ phase: (p) => phases.push(p), transcript: (t) => text.push(t), message: (m) => messages.push(m) }, () => available ? engine : null);
  const result = (entries: Array<[string, boolean]>, index = 0) => engine.onresult?.({
    resultIndex: index, results: entries.map(([transcript, isFinal]) => ({ 0: { transcript }, isFinal }))
  });
  return { input, engine, text, phases, messages, result, calls: () => ({ stops, aborts, starts }) };
}

test("unsupported recognition gives an honest fallback and never enters listening state", () => {
  const s = setup(false);
  s.input.start();
  assert.equal(s.input.active, false);
  assert.deepEqual(s.phases, []);
  assert.equal(s.messages.at(-1), SPEECH_FALLBACK);
});

test("only final results enter the draft, repeated events are deduplicated by result index", () => {
  const s = setup();
  s.input.start(); s.input.start();
  assert.equal(s.calls().starts, 1);
  assert.equal(s.engine.lang, "ru-RU");
  s.engine.onstart?.();
  s.result([["черновик", false]]);
  assert.deepEqual(s.text, []);
  s.result([["Первая фраза", true]]);
  s.result([["Первая фраза", true], ["Первая фраза", true]]);
  assert.deepEqual(s.text, ["Первая фраза", "Первая фраза"]);
  s.input.cancel();
});

test("stop waits for final recognition and the end event before allowing send", () => {
  const s = setup();
  s.input.start(); s.engine.onstart?.(); s.input.stop(); s.input.stop();
  assert.equal(s.calls().stops, 1);
  assert.equal(s.input.active, true);
  assert.equal(s.phases.at(-1), "stopping");
  s.result([["Готовый запрос", true]]);
  s.engine.onend?.();
  assert.deepEqual(s.text, ["Готовый запрос"]);
  assert.equal(s.input.active, false);
  assert.equal(s.phases.at(-1), "idle");
});

test("cancel and a later session ignore stale callbacks, including permission arriving after cancel", () => {
  const s = setup();
  s.input.start();
  const oldResult = s.engine.onresult, oldStart = s.engine.onstart;
  s.input.stop();
  assert.equal(s.calls().aborts, 1);
  oldStart?.();
  oldResult?.({ resultIndex: 0, results: [{ isFinal: true, 0: { transcript: "не добавлять" } }] });
  assert.deepEqual(s.text, []);
  assert.equal(s.input.active, false);
});

test("permission/network errors release the microphone without a stuck recording indicator", () => {
  for (const error of ["not-allowed", "network", "audio-capture"]) {
    const s = setup();
    s.input.start(); s.engine.onerror?.({ error });
    assert.equal(s.input.active, false);
    assert.equal(s.calls().aborts, 1);
    assert.equal(s.phases.at(-1), "idle");
    assert.ok(s.messages.at(-1));
    assert.deepEqual(s.text, []);
  }
});

test("synchronous start failure is recoverable and never opens a second implicit session", () => {
  const s = setup();
  s.engine.start = () => { throw new Error("unavailable"); };
  s.input.start();
  assert.equal(s.input.active, false);
  assert.equal(s.phases.at(-1), "idle");
  assert.equal(s.calls().aborts, 1);
});

test("dictation appends to the latest typed draft without replacing it or duplicating whitespace", () => {
  assert.equal(appendDictation("Сделай таблицу", "  продаж "), "Сделай таблицу продаж");
  assert.equal(appendDictation("Сделай\n", "таблицу"), "Сделай\nтаблицу");
  assert.equal(appendDictation("мой текст", "  "), "мой текст");
  assert.equal(appendDictation("", "Привет"), "Привет");
});
