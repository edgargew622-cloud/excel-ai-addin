import test from "node:test";
import assert from "node:assert/strict";
import { needsApproval, needsReadApproval, overwrittenCells, parseApprovalMode, OVERWRITE_LIMIT } from "./approvalMode";

const write = (filled: number) => ({ kind: "set_range_values", undoAvailable: true, before: Array.from({ length: filled }, () => [1]) });

test("manual asks for everything, auto only for deletions", () => {
  for (const name of ["format_range", "set_range_values", "create_chart", "insert_rows"]) {
    assert.equal(needsApproval("manual", name, null, {}), true, name);
    assert.equal(needsApproval("auto", name, { undoAvailable: false }, {}), false, name);
  }
  for (const name of ["delete_rows", "delete_columns", "delete_sheet"]) assert.equal(needsApproval("auto", name, null, {}), true, name);
  assert.equal(needsApproval("auto", "remove_duplicates", null, {}), true, "на месте — необратимо");
  assert.equal(needsApproval("auto", "remove_duplicates", null, { destSheet: "Без дублей" }), false, "на другой лист — источник цел");
});

test("semi runs what can be undone and asks for the rest", () => {
  assert.equal(needsApproval("semi", "set_range_values", write(3), {}), false);
  assert.equal(needsApproval("semi", "set_range_values", write(OVERWRITE_LIMIT + 1), {}), true, "большая перезапись");
  assert.equal(needsApproval("semi", "set_range_values", { ...write(1), undoAvailable: false }, {}), true, "без отмены");
  assert.equal(needsApproval("semi", "format_range", { undoAvailable: false }, {}), false, "оформление данных не трогает");
  for (const name of ["insert_rows", "insert_columns", "create_table", "delete_chart"]) assert.equal(needsApproval("semi", name, null, {}), true, name);
  assert.equal(needsApproval("semi", "create_pivot_table", { undoAvailable: true }, {}), false);
});

test("overwritten cells count filled cells of every write in a group", () => {
  assert.equal(overwrittenCells({ kind: "set_ranges_values", items: [write(2), { before: [["", null, "x"]] }] }), 3);
  assert.equal(overwrittenCells({ kind: "fill_range", beforeFormulas: [["=A1", ""], [5, ""]] }), 2);
  assert.equal(overwrittenCells(null), 0);
});

test("reading another sheet is asked only in manual mode; unknown stored mode falls back to semi", () => {
  assert.equal(needsReadApproval("manual"), true);
  assert.equal(needsReadApproval("semi"), false);
  assert.equal(parseApprovalMode("что-то"), "semi");
  assert.equal(parseApprovalMode("auto"), "auto");
});
