import test from "node:test";
import assert from "node:assert/strict";
import {
  action,
  clear,
  depth,
  getStructuralRevision,
  guardedContentUndo,
  invalidateAfterStructuralChange,
  isCustomUndoAvailable,
  push,
  setUndoMonitorReady,
  undoLast
} from "./undo";

function enableUndo() {
  clear();
  setUndoMonitorReady(true);
}

test("structural revision increments even when undo stack is empty", () => {
  enableUndo();
  const before = getStructuralRevision();
  assert.equal(invalidateAfterStructuralChange(), 0);
  assert.equal(getStructuralRevision(), before + 1);
});

test("in-flight undo is not returned to stack after structural invalidation", async () => {
  enableUndo();
  let release!: () => void;
  const gate = new Promise<void>((resolve) => { release = resolve; });

  assert.equal(push(action("write A2", async () => {
    await gate;
  })), true);

  const pending = undoLast();
  assert.equal(depth(), 0, "action is removed while undo is in flight");

  invalidateAfterStructuralChange();
  release();

  await assert.rejects(pending, /Структура книги изменилась/i);
  assert.equal(depth(), 0, "stale action must not be re-added after structural change");
});

test("monitor failure disables creation and execution of custom undo", async () => {
  enableUndo();
  let executed = false;
  assert.equal(push(action("safe write", async () => { executed = true; })), true);
  assert.equal(depth(), 1);

  const removed = setUndoMonitorReady(false);
  assert.equal(removed, 1);
  assert.equal(isCustomUndoAvailable(), false);
  assert.equal(depth(), 0);
  assert.equal(push(action("must not be stored", async () => {})), false);
  assert.equal(depth(), 0);

  await assert.rejects(() => undoLast(), /monitor|монитор/i);
  assert.equal(executed, false);
});

test("re-enabling monitor starts a fresh undo era", async () => {
  enableUndo();
  assert.equal(push(action("old action", async () => {})), true);
  setUndoMonitorReady(false);
  assert.equal(depth(), 0);

  setUndoMonitorReady(true);
  assert.equal(isCustomUndoAvailable(), true);
  assert.equal(depth(), 0, "old actions must not be restored");

  let executed = false;
  assert.equal(push(action("new action", async () => { executed = true; })), true);
  assert.match(await undoLast(), /new action/i);
  assert.equal(executed, true);
});

test("monitor failure invalidates an undo already in flight", async () => {
  enableUndo();
  let release!: () => void;
  const gate = new Promise<void>((resolve) => { release = resolve; });

  assert.equal(push(action("write B2", async () => {
    await gate;
  })), true);

  const pending = undoLast();
  setUndoMonitorReady(false);
  release();

  await assert.rejects(pending, /монитор/i);
  assert.equal(depth(), 0);
});

test("content undo refuses to overwrite a newer manual edit", async (t) => {
  enableUndo();
  let writes = 0;
  const range = {
    formulas: [[777]],
    load: () => undefined
  } as { formulas: unknown[][]; load: () => void };
  const ctx = {
    workbook: { worksheets: { getItem: () => ({ getRange: () => range }) } },
    sync: async () => undefined
  };
  const previousExcel = (globalThis as any).Excel;
  Object.defineProperty(range, "formulas", {
    configurable: true,
    get: () => [[777]],
    set: () => { writes += 1; }
  });
  (globalThis as any).Excel = { run: async (fn: (context: unknown) => Promise<unknown>) => fn(ctx) };
  t.after(() => {
    setUndoMonitorReady(false);
    (globalThis as any).Excel = previousExcel;
  });

  const undo = guardedContentUndo(
    "write",
    { sheet: "Sheet1", address: "A1", formulas: [[1]] },
    { sheet: "Sheet1", address: "A1", formulas: [[2]] }
  );
  await assert.rejects(() => undo.undo(), /более свежие изменения/i);
  assert.equal(writes, 0);
});
