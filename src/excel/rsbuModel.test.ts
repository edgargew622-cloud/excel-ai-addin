import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { accountOf, BALANCE, computeForm, detectOsv, osvTotals, readOsv, RESULTS, unmappedAccounts } from "./rsbuModel";
import { PLANNED_TOOLS } from "./plans";

// Синтетическая ОСВ двойной записью — scripts/create-osv.py; ожидаемые строки посчитаны там же по проводкам.
const fixture = JSON.parse(readFileSync(join(process.cwd(), "src", "excel", "fixtures-osv.json"), "utf8"));

test("an OSV saved from 1C is found by its header: account, opening, turnover, closing, debit and credit", () => {
  const cols = detectOsv(fixture.grid)!;
  assert.deepEqual([cols.account, cols.startDebit, cols.startCredit, cols.turnDebit, cols.turnCredit, cols.endDebit, cols.endCredit, cols.firstRow], [0, 2, 3, 4, 5, 6, 7, 5]);
  const rows = readOsv(fixture.grid, cols);
  assert.ok(rows.some((row) => row.account === "60.02"));
  assert.ok(!rows.some((row) => row.account === "Итого"));
  const totals = osvTotals(rows);
  assert.equal(totals.endDebit, totals.endCredit);
  assert.equal(accountOf(1), "01");
  assert.equal(accountOf("10.01"), "10.01");
  assert.equal(accountOf("Итого"), null);
});

test("form 1 lines match the postings: settlements split by side, 84 + 99 as retained earnings, assets = liabilities on both dates", () => {
  const rows = readOsv(fixture.grid, detectOsv(fixture.grid)!);
  const lines = computeForm(BALANCE, rows, true);
  const get = (code: string) => lines.find((item) => item.line.code === code)!;
  for (const [code, value] of Object.entries(fixture.expected.end)) assert.equal(get(code).now, value, `${code} на конец`);
  for (const [code, value] of Object.entries(fixture.expected.start)) assert.equal(get(code).before, value, `${code} на начало`);
  assert.equal(get("1600").now, get("1700").now);
  assert.equal(get("1600").before, get("1700").before);
  assert.deepEqual(unmappedAccounts(rows, [BALANCE, RESULTS]), []);
});

test("form 2 from 90/91/99 turnovers: revenue net of VAT, net profit agrees with account 99", () => {
  const rows = readOsv(fixture.grid, detectOsv(fixture.grid)!);
  const lines = computeForm(RESULTS, rows, false);
  const get = (code: string) => lines.find((item) => item.line.code === code)!;
  for (const [code, value] of Object.entries(fixture.expected.results)) assert.equal(get(code).now, value, code);
  assert.ok(PLANNED_TOOLS.includes("build_rsbu_reports"));
});

test("an account with a balance outside the mapping is named, not lost", () => {
  const grid = fixture.grid.map((row: unknown[]) => [...row]);
  grid.splice(6, 0, ["45", "Товары отгруженные", 0, 0, 0, 0, 1000, 0], ["86", "Целевое финансирование", 0, 0, 0, 0, 0, 1000]);
  const rows = readOsv(grid, detectOsv(grid)!);
  assert.deepEqual(unmappedAccounts(rows, [BALANCE, RESULTS]), ["86"]);
});
