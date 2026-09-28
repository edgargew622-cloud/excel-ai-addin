import test from "node:test";
import assert from "node:assert/strict";
import { findCycles } from "./cycles";
import { auditSheets } from "./audit";

/*
 * Книга из замера 28 сентября 2026 года (Office 2021). Excel во всех этих
 * ячейках показал 0 (G1 — 1, посчитанную до замыкания цикла) с типом
 * «Double»: ни одной ошибки. Имя «Ставка» Excel отдал как «=З84!$D$1».
 */
const MAIN = {
  name: "З84",
  rowIndex: 0,
  columnIndex: 0,
  formulas: [
    ["=B1+1", "=A1*2", "=C1+1", "=E1", "=Ставка*2", "=SUM(A1:B1)", "='З84б'!A1+1", "=SUM(H:H)"],
    ["", "", "", "", "", "", "", 5]
  ]
};
const OTHER = { name: "З84б", rowIndex: 0, columnIndex: 0, formulas: [["='З84'!G1*2"]] };
const NAMES = [{ name: "Ставка", formula: "=З84!$D$1" }];

test("every kind of cycle from the measured workbook is found, and a formula that only reads a cycle is not in it", () => {
  const cycles = findCycles([MAIN, OTHER], NAMES).map((cycle) => cycle.cells.join(" → ")).sort();
  assert.deepEqual(cycles, [
    "З84!A1 → З84!B1 → З84!A1",
    "З84!C1 → З84!C1",
    "З84!D1 → З84!E1 → З84!D1",
    "З84!G1 → З84б!A1 → З84!G1",
    "З84!H1 → З84!H1"
  ]);
});

test("without the name the cycle through it is invisible — names are what make it visible", () => {
  const cycles = findCycles([MAIN, OTHER]).map((cycle) => cycle.cells[0]);
  assert.ok(!cycles.includes("З84!D1"));
});

test("a long chain is shown in order, and a sheet-level name wins over a workbook one", () => {
  // B2 → C2 → D2 → «Итог» (на листе — B3) → B3 → B2. Имя книги «Итог» смотрит
  // на пустую Z99: если бы оно победило, цикла бы не было.
  const chain = { name: "Модель", rowIndex: 1, columnIndex: 1, formulas: [["=C2", "=D2", "=Итог"], ["=B2", "", ""]] };
  const cycles = findCycles([chain], [{ name: "Итог", formula: "=Модель!$Z$99" }, { name: "Итог", formula: "=Модель!$B$3", scope: "Модель" }]);
  assert.equal(cycles.length, 1);
  assert.deepEqual(cycles[0].cells, ["Модель!B2", "Модель!C2", "Модель!D2", "Модель!B3", "Модель!B2"]);
  assert.equal(cycles[0].size, 4);
});

test("the audit reports cycles as proven even though Excel shows no error there", () => {
  const values = MAIN.formulas.map((row) => row.map((formula) => (formula === "" ? "" : typeof formula === "number" ? formula : 0)));
  const types = values.map((row) => row.map((value) => (value === "" ? "Empty" : "Double")));
  const report = auditSheets([
    { ...MAIN, formulasR1C1: MAIN.formulas, values, valueTypes: types },
    { ...OTHER, formulasR1C1: OTHER.formulas, values: [[0]], valueTypes: [["Double"]] }
  ], { names: NAMES, iterative: false });
  assert.equal(report.totals.cycles, 5);
  const cycle = report.proven.find((item) => item.cell === "A1" && /циклическая/.test(item.reason))!;
  assert.match(cycle.reason, /З84!A1 → З84!B1 → З84!A1 \(ячеек в цикле: 2\)/);
  assert.match(cycle.reason, /не показывает здесь ошибку/);
  assert.equal(cycle.content, "=B1+1");

  const iterative = auditSheets([{ ...MAIN, formulasR1C1: MAIN.formulas, values, valueTypes: types }], { names: NAMES, iterative: true });
  assert.ok(iterative.proven.some((item) => /Итеративные вычисления книги включены/.test(item.reason)));
});
