import test from "node:test";
import assert from "node:assert/strict";
import { auditSheets, constantsIn, type AuditSheet } from "./audit";
import { columnIndex } from "./formulaFill";

/** Относительные ссылки A1 → R1C1, как их отдаёт Excel: `=B2*1.1` в C2 → `=RC[-1]*1.1`. */
function toR1C1(formula: unknown, row: number, column: number): unknown {
  if (typeof formula !== "string" || !formula.startsWith("=")) return formula;
  return formula.replace(/(?<![A-Za-z"])([A-Z]{1,3})(\d+)(?![\d(])/g, (_, letters: string, digits: string) => {
    const dr = Number(digits) - row;
    const dc = columnIndex(letters) - column;
    return `R${dr ? `[${dr}]` : ""}C${dc ? `[${dc}]` : ""}`;
  });
}

/** Лист из формул A1 с началом в A1; значения и типы — как посчитал бы Excel. */
function sheetOf(name: string, formulas: unknown[][], values: unknown[][], types: string[][]): AuditSheet {
  return {
    name, rowIndex: 0, columnIndex: 0, formulas, values, valueTypes: types,
    formulasR1C1: formulas.map((row, r) => row.map((formula, c) => toR1C1(formula, r + 1, c + 1)))
  };
}

test("numbers written inside formulas are found, cell addresses are not", () => {
  assert.deepEqual(constantsIn("=B2*1.05"), [1.05]);
  assert.deepEqual(constantsIn("=$B$2+C12"), []);
  assert.deepEqual(constantsIn("=SUM(Лист1!A1:A10)*12"), [12]);
  assert.deepEqual(constantsIn("=LOG10(B2)+1"), []);
  assert.deepEqual(constantsIn('="2025 год"&B2'), []);
  assert.deepEqual(constantsIn("=SUM(2:2)"), []);
});

// Эталон с намеренными ошибками: число вместо формулы (D3), формула не как
// у соседей (D4), деление на пустую ячейку (B5) и её следствие (B7),
// потерянная ссылка (B8), несходящаяся проверка (C9), INDIRECT (B10).
const MODEL = sheetOf("Модель", [
  ["Показатель", 2025, 2026, 2027, 2028],
  ["Выручка", 1000, "=B2*1.1", "=C2*1.1", "=D2*1.1"],
  ["Затраты", 600, "=B3*1.05", 700, "=D3*1.05"],
  ["Прибыль", "=B2-B3", "=C2-C3", "=D2-C3", "=E2-E3"],
  ["Маржа", "=B4/B6", "=C4/C2", "=D4/D2", "=E4/E2"],
  ["", "", "", "", ""],
  ["Удвоенная", "=B5*2", "", "", ""],
  ["Потеря", "=#REF!+1", "", "", ""],
  ["Проверка", "=B4-(B2-B3)", "=C4-(C2-C3)+5", "", ""],
  ["Динамика", '=INDIRECT("B2")', "", "", ""]
], [
  ["Показатель", 2025, 2026, 2027, 2028],
  ["Выручка", 1000, 1100, 1210, 1331],
  ["Затраты", 600, 630, 700, 735],
  ["Прибыль", 400, 470, 580, 596],
  ["Маржа", "#ДЕЛ/0!", 0.43, 0.48, 0.45],
  ["", "", "", "", ""],
  ["Удвоенная", "#ДЕЛ/0!", "", "", ""],
  ["Потеря", "#ССЫЛКА!", "", "", ""],
  ["Проверка", 0, 5, "", ""],
  ["Динамика", 1000, "", "", ""]
], [
  ["String", "Double", "Double", "Double", "Double"],
  ["String", "Double", "Double", "Double", "Double"],
  ["String", "Double", "Double", "Double", "Double"],
  ["String", "Double", "Double", "Double", "Double"],
  ["String", "Error", "Double", "Double", "Double"],
  ["Empty", "Empty", "Empty", "Empty", "Empty"],
  ["String", "Error", "Empty", "Empty", "Empty"],
  ["String", "Error", "Empty", "Empty", "Empty"],
  ["String", "Double", "Double", "Empty", "Empty"],
  ["String", "Double", "Empty", "Empty", "Empty"]
]);

const where = (items: { cell: string; reason: string }[], pattern: RegExp) => items.filter((item) => pattern.test(item.reason)).map((item) => item.cell);

test("proven findings: the root error, the lost reference and a failing check — not the consequence", () => {
  const report = auditSheets([MODEL], { checks: [{ sheet: "Модель", address: "B9:C9" }] });
  assert.deepEqual(where(report.proven, /деление на ноль/), ["B5"]);
  assert.deepEqual(where(report.proven, /потерянная ссылка/), ["B8"]);
  assert.deepEqual(where(report.proven, /контрольное равенство/), ["C9"]);
  assert.equal(report.totals.errorsRoot, 1);
  assert.equal(report.totals.errorsConsequence, 1);
  assert.ok(!report.proven.some((item) => item.cell === "B7"));
});

test("suspicions: a hardcoded number and an odd formula in a row, an empty input, numbers inside formulas", () => {
  const report = auditSheets([MODEL]);
  assert.deepEqual(where(report.suspicions, /число вместо формулы/), ["D3"]);
  assert.deepEqual(where(report.suspicions, /не такая, как у соседей/), ["D4"]);
  assert.deepEqual(where(report.suspicions, /пустую ячейку B6/), ["B5"]);
  assert.ok(where(report.suspicions, /число 1\.1 внутри формулы/).includes("C2"));
});

test("unverified: computed addresses; the report never calls them errors", () => {
  const report = auditSheets([MODEL], { unscanned: ["Огромный"] });
  assert.deepEqual(where(report.unverified, /INDIRECT/), ["B10"]);
  assert.deepEqual(report.unverified.filter((item) => /слишком большой/.test(item.reason)).map((item) => item.sheet), ["Огромный"]);
  assert.ok(!report.proven.some((item) => item.cell === "B10"));
});

test("a clean model gives no proven findings", () => {
  const clean = sheetOf("Чисто", [["", 2025, 2026], ["Выручка", 1000, "=B2*Рост"], ["Итог", "=B2", "=C2"]], [["", 2025, 2026], ["Выручка", 1000, 1100], ["Итог", 1000, 1100]], [["Empty", "Double", "Double"], ["String", "Double", "Double"], ["String", "Double", "Double"]]);
  const report = auditSheets([clean]);
  assert.equal(report.totals.proven, 0);
  assert.equal(report.totals.formulas, 3);
});
