import test from "node:test";
import assert from "node:assert/strict";
import { columnIndex, columnLetters, fillFormulaMatrix, shiftFormula } from "./formulaFill";

test("column letters and indexes convert both ways", () => {
  assert.equal(columnLetters(1), "A");
  assert.equal(columnLetters(26), "Z");
  assert.equal(columnLetters(27), "AA");
  assert.equal(columnLetters(16384), "XFD", "последний столбец листа");
  for (const letters of ["A", "Z", "AA", "XFD"]) {
    assert.equal(columnLetters(columnIndex(letters)), letters);
  }
});

test("relative references move with the cell, like a fill by the corner", () => {
  assert.equal(shiftFormula("=D2*E2", 1, 0), "=D3*E3");
  assert.equal(shiftFormula("=D2*E2", 4, 0), "=D6*E6");
  assert.equal(shiftFormula("=A1", 0, 2), "=C1");
  assert.equal(shiftFormula("=SUM(A1:A5)", 2, 0), "=SUM(A3:A7)");
});

test("dollars pin the part they stand in front of", () => {
  assert.equal(shiftFormula("=$D2*E$2", 3, 1), "=$D5*F$2");
  assert.equal(shiftFormula("=$D$2", 5, 5), "=$D$2");
  assert.equal(shiftFormula("=Справочник!$B$2*A1", 1, 0), "=Справочник!$B$2*A2");
});

test("text and function names are left alone", () => {
  // В кавычках может стоять что угодно, похожее на ссылку.
  assert.equal(shiftFormula('=IF(A1>0,"B2 штук","")', 1, 0), '=IF(A2>0,"B2 штук","")');
  // LOG10( выглядит как ссылка LOG10, но это имя функции.
  assert.equal(shiftFormula("=LOG10(A1)", 1, 0), "=LOG10(A2)");
  // Не формула — не трогаем вовсе.
  assert.equal(shiftFormula("просто текст", 3, 3), "просто текст");
});

test("a reference pushed off the sheet becomes an error, as in Excel", () => {
  assert.equal(shiftFormula("=A1", -1, 0), "=#REF!");
  assert.equal(shiftFormula("=A1", 0, -1), "=#REF!");
  assert.equal(shiftFormula("=$A$1", -5, -5), "=$A$1", "закреплённая ссылка не уезжает");
});

test("the matrix repeats the formula down and across the area", () => {
  assert.deepEqual(fillFormulaMatrix("=D2*E2", 5, 1).flat(), [
    "=D2*E2", "=D3*E3", "=D4*E4", "=D5*E5", "=D6*E6"
  ]);
  assert.deepEqual(fillFormulaMatrix("=A1", 2, 3), [
    ["=A1", "=B1", "=C1"],
    ["=A2", "=B2", "=C2"]
  ]);
});
