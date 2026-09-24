import test from "node:test";
import assert from "node:assert/strict";
import { classifyNumberText, cultureDateOrder, dateByOrder, isDateFormat, parseDateText, parseNumberText, profileData } from "./dataProfile";

const RU = { decimal: ",", group: " " };
const letters = (index: number) => String.fromCharCode(64 + index);

test("numbers written as text are read by the workbook's own separators", () => {
  // Замер 24 сентября 2026 года: ru-RU — запятая и пробел между тысячами.
  assert.equal(parseNumberText("1 200", ",", " "), 1200);
  assert.equal(parseNumberText("1\u00a0200,50", ",", " "), 1200.5, "неразрывный пробел — тот же разделитель");
  assert.equal(parseNumberText("-3", ",", " "), -3);
  assert.equal(parseNumberText(" 42 ", ",", " "), 42);
  assert.equal(parseNumberText("12 34", ",", " "), null, "тысячи только группами по три");
  for (const text of ["1e3", "12%", "1 200 ₽", "", "abc"]) assert.equal(parseNumberText(text, ",", " "), null, text);
});

test("what cannot be decided is not decided", () => {
  assert.deepEqual(classifyNumberText("1 200,5", RU), { kind: "number", value: 1200.5 });
  assert.deepEqual(classifyNumberText("1.5", RU), { kind: "foreignNumber", value: 1.5 });
  // Полтора или тысяча пятьсот — без ответа пользователя не решить.
  assert.deepEqual(classifyNumberText("1,500", RU), { kind: "ambiguousNumber" });
  assert.deepEqual(classifyNumberText("2.300", RU), { kind: "ambiguousNumber" });
  // Ведущие нули — код, а не число: преобразование их бы уничтожило.
  assert.deepEqual(classifyNumberText("007", RU), { kind: "leadingZeros" });
  assert.equal(classifyNumberText("Москва", RU), null);
});

test("a date is taken as certain only when the numbers themselves say so", () => {
  assert.deepEqual(parseDateText("2026-02-01"), { date: { year: 2026, month: 2, day: 1 }, orders: ["YMD"] });
  assert.deepEqual(parseDateText("25.02.2026"), { date: { year: 2026, month: 2, day: 25 }, orders: ["DMY"] });
  assert.deepEqual(parseDateText("02/25/2026"), { date: { year: 2026, month: 2, day: 25 }, orders: ["MDY"] });
  // 1 февраля или 2 января.
  assert.deepEqual(parseDateText("01.02.2026"), { orders: ["DMY", "MDY"] });
  assert.equal(parseDateText("31.02.2026"), null, "такого дня нет");
  assert.equal(parseDateText("01.02.26"), null, "двузначный год не говорит о веке");
  assert.deepEqual(dateByOrder("01.02.2026", "DMY"), { year: 2026, month: 2, day: 1 });
  assert.deepEqual(dateByOrder("01.02.2026", "MDY"), { year: 2026, month: 1, day: 2 });
  assert.equal(dateByOrder("25.02.2026", "MDY"), null, "названный порядок противоречит данным");
});

test("the workbook's date order and date formats are read, not assumed", () => {
  assert.equal(cultureDateOrder("ДД.ММ.ГГГГ"), "DMY");
  assert.equal(cultureDateOrder("M/d/yyyy"), "MDY");
  assert.equal(isDateFormat("dd.mm.yyyy"), true);
  assert.equal(isDateFormat("General"), false);
  assert.equal(isDateFormat("#,##0.00"), false);
  assert.equal(isDateFormat('0.00" дней"'), false, "текст в кавычках — не код даты");
});

test("the profile counts what gets in the way of calculating, with a few examples", () => {
  const values = [
    ["Город", "Сумма", "Дата", "Код"],
    [" Москва", 900, "01.02.2026", "007"],
    ["Москва", "1 200", "25.02.2026", "12"],
    ["Казань  Север", "1,500", 46054, "12"],
    ["Москва\u00a0", "1.5", "", ""],
    ["", "", "", ""],
    ["Москва", "1 200", "25.02.2026", "12"]
  ];
  const types = values.map((row) => row.map((value) => (value === "" ? "Empty" : typeof value === "number" ? "Double" : "String")));
  const formats = values.map((row, r) => row.map((_, c) => (r === 3 && c === 2 ? "dd.mm.yyyy" : "General")));
  const profile = profileData({
    values, formulas: values, valueTypes: types, numberFormat: formats, hasHeaders: true,
    origin: { rowIndex: 0, columnIndex: 0 }, address: "A1:D7",
    culture: { ...RU, dateOrder: "DMY" }, columnName: letters
  });
  const [city, sum, date, code] = profile.columns;
  assert.equal(profile.rows, 6);
  assert.equal(profile.emptyRows, 1);
  assert.equal(city.edgeSpaces.count, 2);
  assert.deepEqual(city.edgeSpaces.examples, ["A2: « Москва»", "A5: «Москва\u00a0»"]);
  assert.equal(city.innerSpaces.count, 1);
  assert.equal(city.nonBreakingSpaces.count, 1);
  assert.equal(sum.numbers, 1);
  assert.equal(sum.numbersAsText.count, 2);
  assert.equal(sum.ambiguousNumbersAsText.count, 1);
  assert.equal(sum.foreignNumbersAsText.count, 1);
  assert.equal(date.dates, 1, "число с форматом даты — дата");
  assert.equal(date.datesAsText.count, 2);
  assert.equal(date.ambiguousDatesAsText.count, 1);
  assert.equal(code.codesWithLeadingZeros.count, 1);
  assert.equal(code.numbersAsText.count, 3);
  assert.equal(profile.duplicateRows.count, 1);
  assert.deepEqual(profile.duplicateRows.examples, ["строка 7 = строка 3"]);
});

test("duplicates after trimming and case are counted separately, as candidates", () => {
  const values = [["Город"], ["Москва"], [" москва "], ["МОСКВА"]];
  const types = values.map((row) => row.map(() => "String"));
  const profile = profileData({
    values, formulas: values, valueTypes: types, numberFormat: types.map((row) => row.map(() => "General")), hasHeaders: true,
    origin: { rowIndex: 0, columnIndex: 0 }, address: "A1:A4", culture: { ...RU, dateOrder: "DMY" }, columnName: letters
  });
  assert.equal(profile.duplicateRows.count, 0);
  assert.equal(profile.duplicateRowsNormalized.count, 2);
});
