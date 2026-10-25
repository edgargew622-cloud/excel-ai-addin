import test from "node:test";
import assert from "node:assert/strict";
import { percentCellsUnderHeaders } from "./percentHeaders";

test("shares under a «%» header get found, typed numbers and formatted cells are left alone", () => {
  // «Книга1567»: второй блок листа «Анализ» остался с долями 0,0789… в формате «Общий».
  const values = [
    ["Статья", "Доля, %", "Изменение", "Вручную, %"],
    ["Долг", 0.26, 100, 12],
    ["Капитал", 0.66, 200, 30],
    ["Итого", 1, 300, ""]
  ];
  const formulas = [
    ["Статья", "Доля, %", "Изменение", "Вручную, %"],
    ["Долг", "=B5/B8", "=D2-C2", 12],
    ["Капитал", "=B6/B8", "=D3-C3", 30],
    ["Итого", "=SUM(B2:B3)", "=D4-C4", ""]
  ];
  const formats = [
    ["General", "General", "General", "General"],
    ["General", "General", "General", "General"],
    ["General", "0.0%", "General", "General"],
    ["General", "General", "General", "General"]
  ];
  assert.deepEqual(percentCellsUnderHeaders(values, formulas, formats), [[1, 1], [3, 1]]);
});

test("the column stops at the first text or empty cell", () => {
  const values = [["Темп, %"], [0.1], ["Итого"], [0.2]];
  const formulas = [["Темп, %"], ["=A9"], ["Итого"], ["=A10"]];
  const formats = [["General"], ["General"], ["General"], ["General"]];
  assert.deepEqual(percentCellsUnderHeaders(values, formulas, formats), [[1, 0]]);
});
