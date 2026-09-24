import test from "node:test";
import assert from "node:assert/strict";
import { buildDcf, parseDcfRequest, type DcfAssumptions } from "./dcf";
import { evaluateGrid } from "./formulaEval.testkit";
import { modelMismatches } from "./modelSheet";
import { PLANNED_TOOLS } from "./plans";

const ASSUMPTIONS: DcfAssumptions = {
  revenue0: 1000, growth: 0.1, ebitMargin: 0.2, taxRate: 0.2, daPct: 0.04, capexPct: 0.05, nwcPct: 0.1,
  wacc: 0.12, terminalGrowth: 0.03, netDebt: 200, shares: 100
};
const REQUEST = { sheet: "DCF", currency: "руб.", units: "млн", source: "слова пользователя", firstYear: 2026, years: 5, assumptions: ASSUMPTIONS };

test("no valuation without every assumption, and growth after the forecast must stay below WACC", () => {
  const { wacc: _wacc, ...partial } = ASSUMPTIONS;
  assert.throws(() => parseDcfRequest({ ...REQUEST, assumptions: partial }), /Не хватает: .*wacc/);
  assert.throws(() => parseDcfRequest({ ...REQUEST, assumptions: { ...ASSUMPTIONS, terminalGrowth: 0.12 } }), /меньше WACC/);
  assert.throws(() => parseDcfRequest({ ...REQUEST, assumptions: { ...ASSUMPTIONS, wacc: 12 } }), /от 0 до 1/);
  assert.throws(() => parseDcfRequest({ ...REQUEST, years: 2 }), /от 3 до 10/);
  // Число акций необязательно: без него нет строки «на акцию», но оценка строится.
  const { shares: _shares, ...noShares } = ASSUMPTIONS;
  assert.equal(buildDcf(parseDcfRequest({ ...REQUEST, assumptions: noShares })).perShare, null);
});

test("every formula of the valuation, the sensitivity table included, gives the panel's own numbers", () => {
  const layout = buildDcf(parseDcfRequest(REQUEST));
  assert.deepEqual(modelMismatches(layout.rows, evaluateGrid(layout.rows) as (number | string)[][]), []);
});

test("the first year and the valuation match a hand calculation; the sensitivity centre is the EV", () => {
  const layout = buildDcf(parseDcfRequest(REQUEST));
  const row = (label: string) => layout.rows.find((cells) => cells[0].expected === label)!;
  // 2026: выручка 1100; EBIT 220; налог −44; амортизация 44; капвложения −55; оборотный капитал 110 − 100 → −10; поток 155.
  assert.ok(Math.abs((row("Свободный денежный поток")[2].expected as number) - 155) < 1e-9);
  const pvSum = [1, 2, 3, 4, 5].reduce((sum, t) => sum + (row("Свободный денежный поток")[1 + t].expected as number) / 1.12 ** t, 0);
  const fcf5 = row("Свободный денежный поток")[6].expected as number;
  const ev = pvSum + fcf5 * 1.03 / 0.09 / 1.12 ** 5;
  assert.ok(Math.abs(layout.ev - ev) < 1e-6);
  assert.ok(Math.abs(layout.equity - (ev - 200)) < 1e-6);
  assert.ok(Math.abs((layout.perShare as number) - (ev - 200) / 100) < 1e-8);
  assert.equal(layout.rows[layout.checkRow - 1][1].expected, 0);
  assert.ok(layout.tvShare > 0.5 && layout.tvShare < 1);
});

test("the sheet keeps currency, units, period and source", () => {
  const layout = buildDcf(parseDcfRequest(REQUEST));
  assert.deepEqual(layout.rows.slice(1, 4).map((row) => row[1].expected), ["руб., млн", "2025 — факт, 2026–2030 — прогноз; дисконтирование на конец года", "слова пользователя"]);
});

test("the valuation goes through the plan registry", () => {
  assert.ok(PLANNED_TOOLS.includes("build_dcf_model"));
});
