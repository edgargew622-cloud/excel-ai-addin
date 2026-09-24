import test from "node:test";
import assert from "node:assert/strict";
import { buildLbo, parseLboRequest, type LboAssumptions } from "./lbo";
import { evaluateGrid } from "./formulaEval.testkit";
import { modelMismatches } from "./modelSheet";
import { PLANNED_TOOLS } from "./plans";

const ASSUMPTIONS: LboAssumptions = {
  ebitda0: 100, entryMultiple: 8, debtMultiple: 5, fees: 20, ebitdaGrowth: 0.08, daPct: 0.2, capexPct: 0.25,
  nwcPct: 0.3, taxRate: 0.2, interestRate: 0.09, cashSweep: 1, exitMultiple: 8
};
const REQUEST = { sheet: "LBO", currency: "руб.", units: "млн", source: "слова пользователя", entryYear: 2025, years: 5, assumptions: ASSUMPTIONS };

test("no model without every assumption; debt may not pay for the whole deal", () => {
  const { exitMultiple: _exit, ...partial } = ASSUMPTIONS;
  assert.throws(() => parseLboRequest({ ...REQUEST, assumptions: partial }), /Не хватает: .*exitMultiple/);
  assert.throws(() => parseLboRequest({ ...REQUEST, assumptions: { ...ASSUMPTIONS, debtMultiple: 9 } }), /покрывает всю цену/);
  assert.throws(() => parseLboRequest({ ...REQUEST, assumptions: { ...ASSUMPTIONS, cashSweep: 100 } }), /от 0 до 1/);
  assert.throws(() => parseLboRequest({ ...REQUEST, years: 2 }), /от 3 до 10/);
});

test("every formula of the model gives the panel's own numbers, and all three checks are zero", () => {
  const layout = buildLbo(parseLboRequest(REQUEST));
  const values = evaluateGrid(layout.rows) as (number | string)[][];
  assert.deepEqual(modelMismatches(layout.rows, values), []);
  assert.deepEqual(layout.checkRows.map((row) => values[row - 1][1]), [0, 0, 0]);
});

test("entry, first year and returns match a hand calculation", () => {
  const layout = buildLbo(parseLboRequest(REQUEST));
  const row = (label: string) => layout.rows.find((cells) => cells[0].expected === label)!;
  // Вход: цена 800, долг 500, расходы 20 → вложение 320.
  assert.equal(layout.equity0, 320);
  // Год 1: EBITDA 108; амортизация 21,6; проценты 45; прибыль до налога 41,4; налог 8,28; прибыль 33,12;
  // капвложения 27; прирост оборотного капитала 2,4; поток 33,12 + 21,6 − 27 − 2,4 = 25,32 — весь на погашение.
  assert.ok(Math.abs((row("Свободный поток до погашения долга")[2].expected as number) - 25.32) < 1e-9);
  assert.ok(Math.abs((row("Долг на конец года")[2].expected as number) - 474.68) < 1e-9);
  assert.ok(Math.abs(layout.irr - (layout.moic ** (1 / 5) - 1)) < 1e-12);
  assert.equal(row("Деньги на конец года")[6].expected, 0);
});

test("a heavy debt load is named by the interest coverage", () => {
  const layout = buildLbo(parseLboRequest({ ...REQUEST, assumptions: { ...ASSUMPTIONS, debtMultiple: 7.5, interestRate: 0.15 } }));
  assert.ok(layout.lowCoverage.includes("2026"));
  assert.ok(layout.debtLeft > 0);
});

test("the model goes through the plan registry", () => {
  assert.ok(PLANNED_TOOLS.includes("build_lbo_model"));
});
