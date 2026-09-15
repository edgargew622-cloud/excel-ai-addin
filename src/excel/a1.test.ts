import test from "node:test";
import assert from "node:assert/strict";
import { assertRangeReference, cellCount, contains, intersects, isDefinedName, parseA1Rect } from "./a1";

test("A1 parser enforces Excel bounds and absolute references", () => {
  assert.deepEqual(parseA1Rect("$A$1:$XFD$1048576")?.kind, "cells");
  assert.equal(parseA1Rect("XFE1"), null);
  assert.equal(parseA1Rect("A1048577"), null);
  assert.equal(parseA1Rect("B2:A1"), null);
});

test("full rows and columns are valid scopes but expose their real size", () => {
  assert.equal(cellCount(parseA1Rect("H:H")!), 1_048_576);
  assert.equal(cellCount(parseA1Rect("1:2")!), 32_768);
});

test("containment and intersection use rectangular coordinates", () => {
  const outer = parseA1Rect("B2:D10")!;
  assert.equal(contains(outer, parseA1Rect("C3:D4")!), true);
  assert.equal(intersects(outer, parseA1Rect("D10:F12")!), true);
  assert.equal(intersects(outer, parseA1Rect("E10:F12")!), false);
});

test("defined names are distinct from cell references", () => {
  assert.equal(isDefinedName("Sales_Total"), true);
  assert.equal(isDefinedName("A1"), false);
  assert.equal(assertRangeReference("Sales_Total"), "Sales_Total");
});
