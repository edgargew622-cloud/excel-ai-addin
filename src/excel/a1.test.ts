import test from "node:test";
import assert from "node:assert/strict";
import { assertRangeReference, cellCount, contains, intersects, isDefinedName, parseA1Rect, splitSheetPrefix } from "./a1";

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

test("a sheet name inside the address is accepted when it is the same sheet", () => {
  assert.deepEqual(splitSheetPrefix("Продажи!D1:D7", "Продажи"), { address: "D1:D7", sheet: "Продажи" });
  // Слабые модели повторяют имя дважды.
  assert.deepEqual(splitSheetPrefix("М_лист!М_лист!A1:G7", "м_лист"), { address: "A1:G7", sheet: "м_лист" });
  assert.deepEqual(splitSheetPrefix("'Мой лист'!B2", "Мой лист"), { address: "B2", sheet: "Мой лист" });
  assert.deepEqual(splitSheetPrefix("'Отчёт ''24'!A1", "Отчёт '24"), { address: "A1", sheet: "Отчёт '24" });
});

test("without a sheet argument the sheet is taken from the address", () => {
  assert.deepEqual(splitSheetPrefix("Итоги!A1"), { address: "A1", sheet: "Итоги" });
  assert.deepEqual(splitSheetPrefix("A1:B2"), { address: "A1:B2", sheet: undefined });
  assert.deepEqual(splitSheetPrefix("Sales_Total", "Лист1"), { address: "Sales_Total", sheet: "Лист1" });
});

test("another sheet in the address is never silently redirected", () => {
  assert.deepEqual(splitSheetPrefix("Итоги!A1", "Продажи"), { address: "Итоги!A1", sheet: "Продажи" });
  assert.deepEqual(splitSheetPrefix("Итоги!Продажи!A1"), { address: "Итоги!Продажи!A1", sheet: undefined });
  assert.throws(() => assertRangeReference("Итоги!A1"), /Имя листа передавайте отдельно в sheet/);
});
