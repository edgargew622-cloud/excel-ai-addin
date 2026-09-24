import test from "node:test";
import assert from "node:assert/strict";
import { buildModel, INPUTS, modelMismatches, parseModelRequest, type Assumptions } from "./threeStatement";
import { PLANNED_TOOLS } from "./plans";
import { evaluateGrid } from "./formulaEval.testkit";

const ASSUMPTIONS: Assumptions = {
  revenue0: 1000, growth: 0.1, cogsPct: 0.6, opexPct: 0.2, daPct: 0.05, capexPct: 0.06,
  receivablesPct: 0.1, inventoryPct: 0.08, payablesPct: 0.07, taxRate: 0.2, interestRate: 0.1,
  repayment: 50, payout: 0.3,
  cash0: 100, ppe0: 500, receivables0: 100, inventory0: 80, payables0: 70, debt0: 300, equity0: 410
};
const REQUEST = { sheet: "Модель", currency: "руб.", units: "тыс.", source: "слова пользователя", firstYear: 2026, years: 5, assumptions: ASSUMPTIONS };

test("a model is not built while an assumption is missing, and defaults are never filled in", () => {
  const { taxRate: _tax, ...partial } = ASSUMPTIONS;
  assert.throws(() => parseModelRequest({ ...REQUEST, assumptions: partial }), /Не хватает: .*taxRate/);
  assert.throws(() => parseModelRequest({ ...REQUEST, currency: "" }), /валюта/);
  assert.throws(() => parseModelRequest({ ...REQUEST, assumptions: { ...ASSUMPTIONS, cogsPct: 60 } }), /от 0 до 1/);
});

test("an opening balance that does not balance is refused with the difference", () => {
  assert.throws(() => parseModelRequest({ ...REQUEST, assumptions: { ...ASSUMPTIONS, cash0: 150 } }), /разница 50/);
});

test("every formula of the model gives exactly what the panel calculated", () => {
  const layout = buildModel(parseModelRequest(REQUEST));
  assert.deepEqual(modelMismatches(layout, evaluateGrid(layout.rows) as (number | string)[][]), []);
});

test("the balance holds in every year, and the first year matches a hand calculation", () => {
  const layout = buildModel(parseModelRequest(REQUEST));
  const row = (label: string) => layout.rows.find((cells) => cells[0].expected === label)!.slice(1).map((cell) => cell.expected as number);
  const assets = row("Итого активы");
  const liabilities = row("Итого обязательства и капитал");
  assets.forEach((value, t) => assert.ok(Math.abs(value - liabilities[t]) < 1e-9, `год ${t}: ${value} против ${liabilities[t]}`));
  // Год 1 вручную: выручка 1100; EBIT = 1100 × (1 − 0,6 − 0,2 − 0,05) = 165; проценты 30; налог 27; прибыль 108.
  assert.ok(Math.abs(row("Выручка")[1] - 1100) < 1e-9);
  assert.ok(Math.abs(row("Операционная прибыль (EBIT)")[1] - 165) < 1e-9);
  assert.ok(Math.abs(row("Чистая прибыль")[1] - 108) < 1e-9);
  assert.equal(row("Долг")[5], 50);
  assert.deepEqual(layout.rows[layout.checkRow - 1].slice(1).map((cell) => cell.expected), [0, 0, 0, 0, 0, 0]);
});

test("the sheet keeps currency, units, period, source and every assumption as an input cell", () => {
  const layout = buildModel(parseModelRequest(REQUEST));
  const text = layout.rows.slice(0, 4).map((row) => `${row[0].expected}: ${row[1].expected}`);
  assert.deepEqual(text, ["Трёхотчётная модель: ", "Валюта и единицы: руб., тыс.", "Период: 2025 — факт, 2026–2030 — прогноз", "Источник допущений: слова пользователя"]);
  assert.equal(layout.rows.filter((row) => row[1].input).length, INPUTS.length);
});

test("years without enough cash are named, not hidden", () => {
  const layout = buildModel(parseModelRequest({ ...REQUEST, assumptions: { ...ASSUMPTIONS, capexPct: 0.4, cash0: 10, equity0: 320 } }));
  assert.ok(layout.negativeCash.length > 0);
  assert.equal(layout.negativeCash[0], "2026");
});

test("the model goes through the plan registry", () => {
  assert.ok(PLANNED_TOOLS.includes("build_three_statement_model"));
});
