import test from "node:test";
import assert from "node:assert/strict";
import { columnIndex, columnLetters } from "./formulaFill";

test("column letters and indexes convert both ways", () => {
  assert.equal(columnLetters(1), "A");
  assert.equal(columnLetters(26), "Z");
  assert.equal(columnLetters(27), "AA");
  assert.equal(columnLetters(16384), "XFD", "последний столбец листа");
  for (const letters of ["A", "Z", "AA", "XFD"]) {
    assert.equal(columnLetters(columnIndex(letters)), letters);
  }
});
