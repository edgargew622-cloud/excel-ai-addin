import test from "node:test";
import assert from "node:assert/strict";
import { dateGroupFormula, dateGroupLabel, expectPivot, isDateFormat, pivotMismatches } from "./pivotModel";

/*
 * Раскладки ниже прочитаны из настоящего Excel (Office 2021) 27 сентября
 * 2026 года — `pivot.layout.getRange().values` для сводных по этим же
 * данным. Расчёт панели обязан сойтись с ними без единого расхождения.
 */

const sales = [
  ["Город", "Товар", "Сумма", "Кол"],
  ["Москва", "Чай", 100, 1],
  ["Москва", "Кофе", 200, 2],
  ["Казань", "Чай", 50, 3],
  ["Казань", "Кофе", 70, 4],
  ["Омск", "Чай", 30, 5],
  ["Москва", "Чай", 10, 6],
  ["Омск", "Сок", 5, 7],
  ["Казань", "Сок", 8, 8]
];
const sum = (field: string) => ({ field, aggregation: "sum" as const });

const measuredOneValue = [
  ["Сумма по полю Сумма", "Товар", "", "", ""],
  ["Город", "Кофе", "Сок", "Чай", "Общий итог"],
  ["Казань", 70, 8, 50, 128],
  ["Москва", 200, "", 110, 310],
  ["Омск", "", 5, 30, 35],
  ["Общий итог", 270, 13, 190, 473]
];

const measuredTwoValues = [
  ["", "Товар", "Значения", "", "", "", "", "", ""],
  ["", "Кофе", "", "Сок", "", "Чай", "", "Итог Сумма по полю Сумма", "Итог Сумма по полю Кол"],
  ["Город", "Сумма по полю Сумма", "Сумма по полю Кол", "Сумма по полю Сумма", "Сумма по полю Кол", "Сумма по полю Сумма", "Сумма по полю Кол", "", ""],
  ["Казань", 70, 4, 8, 8, 50, 3, 128, 15],
  ["Москва", 200, 2, "", "", 110, 7, 310, 9],
  ["Омск", "", "", 5, 7, 30, 5, 35, 12],
  ["Общий итог", 270, 6, 13, 15, 190, 15, 473, 36]
];

test("a column field matches the layout Excel really built, one value field", () => {
  const expected = expectPivot(sales, ["Город"], [sum("Сумма")], { columnField: "Товар" });
  assert.equal(expected.headerRows, 2);
  assert.equal(expected.height, 6);
  assert.equal(expected.width, 5);
  assert.deepEqual(pivotMismatches(expected, measuredOneValue), []);
});

test("a column field with two value fields: three header rows, a group of columns per item", () => {
  const expected = expectPivot(sales, ["Город"], [sum("Сумма"), sum("Кол")], { columnField: "Товар" });
  assert.equal(expected.headerRows, 3);
  assert.equal(expected.width, 9);
  assert.deepEqual(pivotMismatches(expected, measuredTwoValues), []);
});

test("a number where Excel should leave the cell empty is a mismatch, and so is a wrong one", () => {
  const expected = expectPivot(sales, ["Город"], [sum("Сумма")], { columnField: "Товар" });
  const zeroInsteadOfEmpty = measuredOneValue.map((row) => [...row]);
  zeroInsteadOfEmpty[3][2] = 0;
  assert.match(pivotMismatches(expected, zeroInsteadOfEmpty).join("; "), /Москва \(Сок\): 0 вместо пустой ячейки/);
  const wrong = measuredOneValue.map((row) => [...row]);
  wrong[2][3] = 51;
  assert.match(pivotMismatches(expected, wrong).join("; "), /Казань \(Чай\): 51 вместо 50/);
});

test("column items are matched by name, not by position", () => {
  const expected = expectPivot(sales, ["Город"], [sum("Сумма")], { columnField: "Товар" });
  // Тот же смысл, другой порядок столбцов — сверка не должна зависеть от сортировки.
  const reordered = measuredOneValue.map((row) => [row[0], row[3], row[2], row[1], row[4]]);
  reordered[0] = ["Сумма по полю Сумма", "", "", "Товар", ""];
  assert.deepEqual(pivotMismatches(expected, reordered), []);
  const renamed = measuredOneValue.map((row) => [...row]);
  renamed[1][2] = "Сок ";
  assert.match(pivotMismatches(expected, renamed).join("; "), /лишний столбец «Сок »/);
});

test("top 2 cities by sum: the total counts only what is left, as Excel showed (438, not 473)", () => {
  const expected = expectPivot(sales, ["Город"], [sum("Сумма")], { filters: [{ field: "Город", top: 2, by: 0 }] });
  assert.deepEqual(expected.grandTotals, [438]);
  assert.deepEqual(expected.filteredOut, [{ field: "Город", items: ["Омск"] }]);
  assert.equal(expected.fullHeight, 5, "до фильтра сводная выше — место проверяется и под неё");
  const measured = [["Город", "Сумма по полю Сумма"], ["Казань", 128], ["Москва", 310], ["Общий итог", 438]];
  assert.deepEqual(pivotMismatches(expected, measured), []);
});

test("keeping chosen items matches Excel's manual filter", () => {
  const expected = expectPivot(sales, ["Город"], [sum("Сумма")], { filters: [{ field: "Город", include: ["москва", "Омск"] }] });
  const measured = [["Город", "Сумма по полю Сумма"], ["Москва", 310], ["Омск", 35], ["Общий итог", 345]];
  assert.deepEqual(pivotMismatches(expected, measured), []);
});

test("bottom N and a tie on the boundary are named", () => {
  const tied = [["Город", "Сумма"], ["А", 10], ["Б", 10], ["В", 30]];
  const expected = expectPivot(tied, ["Город"], [sum("Сумма")], { filters: [{ field: "Город", bottom: 1, by: 0 }] });
  assert.match(expected.warnings.join(" "), /равные значения/);
});

test("sorting by value descending is checked in the built pivot", () => {
  const expected = expectPivot(sales, ["Город"], [sum("Сумма")], { sort: { field: "Город", by: 0, order: "desc" } });
  const measured = [["Город", "Сумма по полю Сумма"], ["Москва", 310], ["Казань", 128], ["Омск", 35], ["Общий итог", 473]];
  assert.deepEqual(pivotMismatches(expected, measured), []);
  const alphabetical = [["Город", "Сумма по полю Сумма"], ["Казань", 128], ["Москва", 310], ["Омск", 35], ["Общий итог", 473]];
  assert.match(pivotMismatches(expected, alphabetical).join("; "), /порядок по убыванию нарушен: «Казань» \(128\) стоит перед «Москва» \(310\)/);
});

test("two row fields and a column field, with subtotals, match Excel", () => {
  const dates = [
    ["Дата", "Город", "Сумма", "Месяц", "Квартал"],
    [46037, "Москва", 100, "2026-01", "2026 К1"],
    [46056, "Москва", 200, "2026-02", "2026 К1"],
    [46073, "Казань", 50, "2026-02", "2026 К1"],
    [46122, "Казань", 70, "2026-04", "2026 К2"],
    [46204, "Омск", 30, "2026-07", "2026 К3"]
  ];
  const measured = [
    ["Сумма по полю Сумма", "", "Город", "", "", ""],
    ["Квартал", "Месяц", "Казань", "Москва", "Омск", "Общий итог"],
    ["2026 К1", "2026-01", "", 100, "", 100],
    ["", "2026-02", 50, 200, "", 250],
    ["2026 К1 Итог", "", 50, 300, "", 350],
    ["2026 К2", "2026-04", 70, "", "", 70],
    ["2026 К2 Итог", "", 70, "", "", 70],
    ["2026 К3", "2026-07", "", "", 30, 30],
    ["2026 К3 Итог", "", "", "", 30, 30],
    ["Общий итог", "", 120, 300, 30, 450]
  ];
  const expected = expectPivot(dates, ["Квартал", "Месяц"], [sum("Сумма")], { columnField: "Город" });
  assert.deepEqual(pivotMismatches(expected, measured), []);
});

test("date group labels are what Excel's helper formulas gave in the measurement", () => {
  // Замер 27.09.2026: 15.01.2026 → «2026-01», «2026 К1»; 01.07.2026 → «2026-07», «2026 К3».
  assert.equal(dateGroupLabel(46037, "month"), "2026-01");
  assert.equal(dateGroupLabel(46037, "quarter"), "2026 К1");
  assert.equal(dateGroupLabel(46204, "month"), "2026-07");
  assert.equal(dateGroupLabel(46204, "quarter"), "2026 К3");
  assert.equal(dateGroupLabel(46204.75, "year"), 2026, "время суток не сдвигает дату");
  assert.equal(dateGroupFormula("A2", "month"), '=IF(A2="","",YEAR(A2)&"-"&TEXT(MONTH(A2),"00"))');
});

test("a date format is told apart from General, text and plain numbers", () => {
  assert.ok(isDateFormat("dd.mm.yyyy"));
  assert.ok(isDateFormat("[$-419]d mmmm yyyy;@"));
  assert.ok(isDateFormat("ДД.ММ.ГГГГ"));
  assert.ok(!isDateFormat("General"));
  assert.ok(!isDateFormat("Основной"));
  assert.ok(!isDateFormat("0.00"));
  assert.ok(!isDateFormat('#,##0 "руб."'));
});
