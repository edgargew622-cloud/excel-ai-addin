import test from "node:test";
import assert from "node:assert/strict";
import { analyzeGrid, linearFit, pearson, quantileInc, round, sampleStd } from "./dataAnalysis";
import { sheetsReadBy } from "../agent/readScope";

test("statistics match Excel functions on the same numbers", () => {
  const sorted = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
  // КВАРТИЛЬ.ВКЛ(1..10; 1) = 3,25; МЕДИАНА = 5,5; КВАРТИЛЬ.ВКЛ(…; 3) = 7,75.
  assert.equal(quantileInc(sorted, 0.25), 3.25);
  assert.equal(quantileInc(sorted, 0.5), 5.5);
  assert.equal(quantileInc(sorted, 0.75), 7.75);
  // СТАНДОТКЛОН.В(2;4;4;4;5;5;7;9) = 2,13809…
  assert.equal(round(sampleStd([2, 4, 4, 4, 5, 5, 7, 9]), 4), 2.138);
  // КОРРЕЛ и НАКЛОН на прямой y = 3x + 1 — 1 и 3; на обратной — −1.
  assert.equal(pearson([1, 2, 3, 4], [4, 7, 10, 13]), 1);
  assert.equal(pearson([1, 2, 3, 4], [13, 10, 7, 4]), -1);
  const fit = linearFit([1, 2, 3, 4], [4, 7, 10, 13]);
  assert.deepEqual([fit.slope, fit.intercept, fit.r2], [3, 1, 1]);
  assert.ok(Number.isNaN(pearson([1, 1, 1], [2, 3, 4])), "постоянный ряд — корреляции нет");
});

// Даты как в Excel: серийные номера и формат даты; 1 января 2026 = 46023.
const DAY0 = 46023;
function salesGrid() {
  const values: unknown[][] = [["Дата", "Регион", "Выручка", "Реклама", "Возвраты"]];
  const numberFormat: unknown[][] = [["General", "General", "General", "General", "General"]];
  for (let i = 0; i < 24; i++) {
    const revenue = i === 17 ? 900 : 100 + i * 5; // рост и один выброс
    values.push([DAY0 + i * 7, i % 2 ? "Север" : "Юг", revenue, 10 + i * 0.5, i % 3]);
    numberFormat.push(["dd.mm.yyyy", "General", "#,##0", "#,##0", "0"]);
  }
  return { values, numberFormat };
}

test("trend by dates, outlier with its address, correlations and skipped text columns", () => {
  const { values, numberFormat } = salesGrid();
  const result = analyzeGrid({ values, numberFormat, hasHeaders: true, origin: { rowIndex: 0, columnIndex: 0 } });
  assert.equal(result.rows, 24);
  assert.match(result.xAxis, /даты столбца A \(«Дата»\)/);
  const revenue = result.columns.find((column) => column.name === "Выручка")!;
  assert.equal(revenue.column, "C");
  assert.equal(revenue.count, 24);
  assert.equal((revenue.trend as any).direction, "рост", "выброс не прячет рост");
  assert.equal((revenue.trend as any).excludedOutliers, 1);
  assert.equal((revenue.trend as any).r2, 1, "без выброса точки лежат на прямой");
  assert.equal((revenue.trend as any).slopeUnit, "в день");
  assert.equal((revenue.trend as any).from, "2026-01-01");
  assert.ok((revenue.trend as any).slopePerMonth > 0);
  // Выброс — 18-я строка данных, то есть строка 19 листа: C19 = 900.
  assert.deepEqual(revenue.outliers.cells, [{ address: "C19", value: 900 }]);
  assert.equal(revenue.outliers.count, 1);

  // Пирсон как КОРРЕЛ: один выброс опускает связь до слабой; без него она почти полная.
  const pair = result.correlations!.pairs.find((item) => item.a === "Выручка" && item.b === "Реклама")!;
  assert.equal(pair.r, 0.382);
  assert.equal(pair.strength, "слабая, прямая");
  assert.equal(pair.n, 24);
  assert.equal(pair.rWithoutOutliers, 1);
  assert.equal(pair.strengthWithoutOutliers, "сильная, прямая");
  assert.deepEqual(result.skipped.map((item) => [item.column, item.reason]), [["B", "в основном текст"]]);
  assert.ok(result.notes.some((note) => /не причин/.test(note)));
});

test("without dates the trend follows row order; a flat noisy column has no clear trend", () => {
  const values: unknown[][] = [["Шаг", "Шум"], ...[5, 9, 4, 8, 5, 9, 4, 8, 5, 9].map((v, i) => [i + 1, v])];
  const result = analyzeGrid({ values, hasHeaders: true, origin: { rowIndex: 4, columnIndex: 2 } });
  assert.equal(result.xAxis, "порядок строк");
  const noise = result.columns.find((column) => column.name === "Шум")!;
  assert.equal(noise.column, "D", "буквы — по положению области на листе");
  assert.equal((noise.trend as any).direction, "явного тренда нет");
  assert.equal(result.analyzedRows, "6:15");
});

test("chosen columns, the trend axis and refusals", () => {
  const { values, numberFormat } = salesGrid();
  const origin = { rowIndex: 0, columnIndex: 0 };
  const only = analyzeGrid({ values, numberFormat, hasHeaders: true, origin, columns: ["Выручка", "d"] });
  assert.deepEqual(only.columns.map((column) => column.name), ["Выручка", "Реклама"]);
  // Живая проверка 30.09.2026: невыбранный числовой столбец назывался «мало чисел», и модель так и сказала.
  assert.deepEqual(only.skipped.map((item) => [item.name, item.reason]), [["Регион", "в основном текст"], ["Возвраты", "не выбран в columns"]]);
  const byAds = analyzeGrid({ values, numberFormat, hasHeaders: true, origin, columns: ["Выручка"], xColumn: "Реклама" });
  assert.match(byAds.xAxis, /значения столбца D/);
  assert.equal(byAds.correlations, null, "один столбец — корреляций нет");
  assert.throws(() => analyzeGrid({ values, numberFormat, hasHeaders: true, origin, columns: ["Прибыль"] }), /Столбца «Прибыль» нет/);
  assert.throws(() => analyzeGrid({ values, numberFormat, hasHeaders: true, origin, columns: ["Регион"] }), /недостаточно чисел/);
  assert.throws(() => analyzeGrid({ values: [["А"], [1], [2]], hasHeaders: true, origin }), /хотя бы три строки/);
  assert.throws(() => analyzeGrid({ values: [["А"], ["x"], ["y"], ["z"]], hasHeaders: true, origin }), /нет числовых столбцов/);
});

test("analysis reads only the sheet it names, like any other read", async () => {
  const io = { sheetOfAddress: async (sheet: string) => sheet, allSheets: async () => ["Лист1", "Секрет"] } as any;
  assert.deepEqual(await sheetsReadBy("analyze_range", { sheet: "Продажи", address: "A1:F500" }, io), ["Продажи", "Продажи"]);
});
