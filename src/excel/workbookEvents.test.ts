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
