import test from "node:test";
import assert from "node:assert/strict";
import { isStructuralChangeType } from "./workbookEvents";
import {
  action,
  clear,
  depth,
  invalidateAfterStructuralChange,
  push,
  setUndoMonitorReady
} from "./undo";
import { bumpWorkbookRevision, getWorkbookRevision, resetWorkbookRevision } from "./workbookRevision";

test("recognizes structural Excel changes that can shift A1 addresses", () => {
  for (const type of [
    "RowInserted",
    "RowDeleted",
    "ColumnInserted",
    "ColumnDeleted",
    "CellInserted",
    "CellDeleted"
  ]) {
    assert.equal(isStructuralChangeType(type), true, type);
  }
  assert.equal(isStructuralChangeType("RangeEdited"), false);
  assert.equal(isStructuralChangeType("Unknown"), false);
});

test("structural invalidation clears stale undo addresses", () => {
  clear();
  setUndoMonitorReady(true);
  assert.equal(push(action("write A2", async () => {})), true);
  assert.equal(push(action("format B4", async () => {})), true);
  assert.equal(depth(), 2);
  assert.equal(invalidateAfterStructuralChange(), 2);
  assert.equal(depth(), 0);
});

test("structural revision invalidates content and formatting snapshots together", () => {
  resetWorkbookRevision();
  bumpWorkbookRevision("content");
  bumpWorkbookRevision("format");
  const before = getWorkbookRevision();
  const after = bumpWorkbookRevision("structure");
  assert.equal(after.sequence, before.sequence + 1);
  assert.equal(after.structure, before.structure + 1);
  assert.equal(after.content, before.content + 1);
  assert.equal(after.format, before.format + 1);
});

test("a new empty sheet does not wipe the undo history", () => {
  // Проверка 21 сентября 2026 года: создание листа стирало собственную же
  // отмену. Новый лист пуст и адреса на других листах не двигает, поэтому
  // прежние отмены остаются в силе; удаление листа — по-прежнему чистит.
  clear();
  setUndoMonitorReady(true);
  try {
    push(action("запись значений Данные!A1", async () => undefined));
    assert.equal(depth(), 1);
    bumpWorkbookRevision("structure");
    assert.equal(depth(), 1, "добавление листа историю не трогает");
    assert.equal(invalidateAfterStructuralChange(), 1, "удаление листа — чистит");
    assert.equal(depth(), 0);
  } finally {
    clear();
    setUndoMonitorReady(false);
  }
});

test("renaming a sheet keeps the undo history; deleting one clears it — through the real handlers", async () => {
  // Этап 7, 7.3.1: переименование на адреса не влияет, история отмены
  // должна остаться. Обработчики берутся из самой регистрации монитора.
  const { ensureStructuralChangeMonitor } = await import("./workbookEvents");
  const handlers: Record<string, (event?: unknown) => Promise<void>> = {};
  const event = (name: string) => ({ add: (handler: any) => { handlers[name] = handler; } });
  (globalThis as any).Office = { context: { requirements: { isSetSupported: () => true } } };
  (globalThis as any).Excel = {
    run: async (fn: any) => fn({
      workbook: {
        worksheets: {
          onChanged: event("changed"), onAdded: event("added"), onDeleted: event("deleted"), onCalculated: event("calculated"),
          onFormatChanged: event("format"), onProtectionChanged: event("protection"), onNameChanged: event("name")
        }
      },
      sync: async () => undefined
    })
  };
  assert.equal(await ensureStructuralChangeMonitor(), true);
  try {
    push(action("запись значений Продажи!A1", async () => undefined));
    assert.equal(depth(), 1);
    await handlers.name();
    assert.equal(depth(), 1, "переименование историю не трогает");
    await handlers.added();
    assert.equal(depth(), 1, "добавление листа — тоже");
    await handlers.deleted();
    assert.equal(depth(), 0, "удаление листа — чистит");
  } finally {
    clear();
    setUndoMonitorReady(false);
  }
});
