import test from "node:test";
import assert from "node:assert/strict";
import { groupStats, multiplesLayout } from "./multiples";
import { PLANNED_TOOLS } from "./plans";

// Таблица на листе с A1: компания, капитализация, чистый долг, EBITDA, выручка, чистая прибыль.
const TABLE = [
  ["Компания", "Капитализация", "Чистый долг", "EBITDA", "Выручка", "Чистая прибыль"],
  ["Альфа", 1000, 200, 100, 500, 50],
  ["Бета", 500, -50, -10, 400, 25],
  ["Гамма", 800, 100, 90, "", -5]
];
const COLUMNS = { marketCap: 2, netDebt: 3, ebitda: 4, revenue: 5, netIncome: 6 };

test("add_multiples goes through the plan registry", () => {
  assert.ok(PLANNED_TOOLS.includes("add_multiples"));
});

test("EV and every multiple per company, with losses and negative EBITDA left out rather than shown as nonsense", () => {
  const { layout, multiples } = multiplesLayout(TABLE, 1, 1, COLUMNS, 6, 1);
  assert.deepEqual(multiples, ["evEbitda", "evRevenue", "pe"]);
  const expected = layout.rows.map((row) => row.map((cell) => cell.expected));
  assert.deepEqual(expected[1], ["Компания", "EV", "EV/EBITDA", "EV/Выручка", "P/E"]);
  assert.deepEqual(expected[2], ["Альфа", 1200, 12, 2.4, 20]);
  // Бета: чистые деньги (долг −50) уменьшают EV; EBITDA отрицательная — EV/EBITDA не имеет смысла.
  assert.deepEqual(expected[3], ["Бета", 450, "", 1.125, 20]);
  // Гамма: нет выручки, убыток — ни EV/Выручка, ни P/E.
  assert.deepEqual(expected[4], ["Гамма", 900, 10, "", ""]);
  assert.deepEqual(layout.undefinedCells, ["C9", "D10", "E10"]);
  // Медиана — только по осмысленным значениям, как MEDIAN по ячейкам с пустыми строками.
  assert.deepEqual(expected[6], ["Медиана", 900, 11, 1.7625, 20]);
  // Квартили (10.4) — КВАРТИЛЬ.ВКЛ: позиция (n−1)·p с интерполяцией; пустые не входят.
  assert.deepEqual(expected[8], ["1-й квартиль", 675, 10.5, 1.44375, 20]);
  assert.deepEqual(expected[9], ["3-й квартиль", 1050, 11.5, 2.08125, 20]);
  assert.deepEqual(expected[10], ["Минимум", 450, 10, 1.125, 20]);
  assert.deepEqual(expected[11], ["Максимум", 1200, 12, 2.4, 20]);
});

test("group statistics are found by row name, not position — the median is the median", () => {
  const { layout } = multiplesLayout(TABLE, 1, 1, COLUMNS, 6, 1);
  const values = layout.rows.map((row) => row.map((cell) => cell.expected));
  const stats = groupStats(layout, values, ["EV/EBITDA", "EV/Выручка", "P/E"]);
  // Живая проверка 30.09.2026: «медиана» бралась 4-й строкой с конца и после квартилей стала 1-м квартилем.
  assert.deepEqual(stats["Медиана"], { EV: 900, "EV/EBITDA": 11, "EV/Выручка": 1.7625, "P/E": 20 });
  assert.deepEqual(stats["1-й квартиль"], { EV: 675, "EV/EBITDA": 10.5, "EV/Выручка": 1.44375, "P/E": 20 });
  assert.deepEqual(Object.keys(stats), ["Медиана", "Среднее", "1-й квартиль", "3-й квартиль", "Минимум", "Максимум"]);
});

test("the formulas point at the table cells and guard the denominator the same way the panel does", () => {
  const { layout } = multiplesLayout(TABLE, 1, 1, COLUMNS, 6, 1);
  const alpha = layout.rows[2].map((cell) => cell.formula);
  assert.equal(alpha[0], "=$A2");
  assert.equal(alpha[1], '=IF(AND(ISNUMBER(B2),ISNUMBER(C2)),B2+C2,"")');
  assert.equal(alpha[2], '=IF(AND(ISNUMBER(B8),ISNUMBER(D2)),IF(D2>0,B8/D2,""),"")', "EV/EBITDA — от EV в самом блоке");
  assert.equal(alpha[4], '=IF(AND(ISNUMBER(B2),ISNUMBER(F2)),IF(F2>0,B2/F2,""),"")', "P/E — капитализация к прибыли");
  assert.equal(layout.rows[6][2].formula, '=IF(COUNT(C$8:C$10)=0,"",MEDIAN(C$8:C$10))');
  assert.equal(layout.rows[8][2].formula, '=IF(COUNT(C$8:C$10)=0,"",QUARTILE.INC(C$8:C$10,1))');
  assert.equal(layout.rows[9][4].formula, '=IF(COUNT(E$8:E$10)=0,"",QUARTILE.INC(E$8:E$10,3))');
  assert.equal(layout.rowFormats[8], '0.0"x"');
  assert.equal(layout.rowFormats[2], '0.0"x"');
  assert.equal(layout.columnFormats?.[1], "#,##0", "EV — деньги, а не «x»");
});

test("a ready EV column is used as is, and only the multiples that have columns are built", () => {
  const table = [["Компания", "EV", "Выручка"], ["Альфа", 1200, 600]];
  const { layout, multiples } = multiplesLayout(table, 1, 1, { enterpriseValue: 2, revenue: 3 }, 4, 1);
  assert.deepEqual(multiples, ["evRevenue"]);
  assert.deepEqual(layout.rows[2].map((cell) => cell.expected), ["Альфа", 1200, 2]);
  assert.match(String(layout.rows[0][0].expected), /EV — из таблицы/);
});
