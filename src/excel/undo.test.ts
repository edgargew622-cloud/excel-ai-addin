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

test("undo keeps literal text as text instead of re-entering it", async () => {
  const { restorableFormulas } = await import("./undo");
  const snapshot = {
    sheet: "Продажи",
    address: "A2:D2",
    // Текстовая дата, текст с нулями, формула с текстовым результатом, число.
    formulas: [["2026-09-04", "00123", '="x"', 4]],
    values: [["2026-09-04", "00123", "x", 4]],
    valueTypes: [["String", "String", "String", "Double"]]
  };
  assert.deepEqual(restorableFormulas(snapshot), [["'2026-09-04", "'00123", '="x"', 4]]);
});

test("old snapshots without value types restore as before", async () => {
  const { restorableFormulas } = await import("./undo");
  assert.deepEqual(restorableFormulas({ sheet: "Лист", address: "A1", formulas: [["2026-09-04"]] }), [["2026-09-04"]]);
});

test("empty cells stay empty on restore", async () => {
  const { restorableFormulas } = await import("./undo");
  const restored = restorableFormulas({
    sheet: "Лист", address: "A1:B1",
    formulas: [["", "текст"]], values: [["", "текст"]], valueTypes: [["Empty", "String"]]
  });
  assert.deepEqual(restored, [["", "'текст"]]);
});

/** Макет столбца таблицы Excel: запись всей области с одной формулой в
 * иначе пустом столбце протягивает её на все строки, как вычисляемый столбец.
 * Поячеечная запись пустоты протяжку не вызывает — так ведёт себя Excel. */
function tableColumn(options: { stubborn?: boolean } = {}) {
  let cells: unknown[] = ["", "", "", ""];
  const range: any = {
    load: () => undefined,
    get formulas() { return cells.map((cell) => [cell]); },
    set formulas(matrix: unknown[][]) {
      const incoming = matrix.map((row) => row[0]);
      const formula = incoming.find((cell) => typeof cell === "string" && cell.startsWith("="));
      const othersEmpty = incoming.filter((cell) => cell !== formula).every((cell) => cell === "");
      cells = formula && othersEmpty ? incoming.map(() => formula) : incoming;
    },
    getCell: (r: number) => ({
      load: () => undefined,
      address: `Продажи!G${r + 2}`,
      set formulas(matrix: unknown[][]) { if (!options.stubborn) cells[r] = matrix[0][0]; }
    })
  };
  const ctx: any = { sync: async () => undefined };
  return { range, ctx, cells: () => cells };
}

test("a formula spread by a table column is repaired cell by cell", async () => {
  const { repairMismatchedCells } = await import("./undo");
  const column = tableColumn();
  const intended = [["=1/0"], [""], [""], [""]];
  column.range.formulas = intended;
  assert.deepEqual(column.cells(), ["=1/0", "=1/0", "=1/0", "=1/0"], "макет воспроизводит протяжку");

  const remaining = await repairMismatchedCells(column.ctx, column.range, {
    property: "formulas", expected: intended, toWrite: intended
  });
  assert.deepEqual(remaining, []);
  assert.deepEqual(column.cells(), ["=1/0", "", "", ""]);
});

test("cells that still differ after repair are named, not hidden", async () => {
  const { repairMismatchedCells } = await import("./undo");
  const column = tableColumn({ stubborn: true });
  const intended = [["=1/0"], [""], [""], [""]];
  column.range.formulas = intended;

  const remaining = await repairMismatchedCells(column.ctx, column.range, {
    property: "formulas", expected: intended, toWrite: intended
  });
  assert.deepEqual(remaining, ["Продажи!G3", "Продажи!G4", "Продажи!G5"]);
});

test("a matching range needs no repair and writes nothing", async () => {
  const { repairMismatchedCells } = await import("./undo");
  let writes = 0;
  const range: any = {
    load: () => undefined,
    formulas: [["a"], ["b"]],
    getCell: () => { writes += 1; return {}; }
  };
  const remaining = await repairMismatchedCells({ sync: async () => undefined } as any, range, {
    property: "formulas", expected: [["a"], ["b"]], toWrite: [["a"], ["b"]]
  });
  assert.deepEqual(remaining, []);
  assert.equal(writes, 0);
});
