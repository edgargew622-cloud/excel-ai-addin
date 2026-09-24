import test from "node:test";
import assert from "node:assert/strict";
import { cleanText, dateFormatCode, dateFromText, excelSerial, formatDate, numberFromText } from "./cleanModel";

const RU = { decimal: ",", group: " " };
const NBSP = String.fromCharCode(160);

test("extra spaces go, as Excel TRIM removes them; non-breaking ones too", () => {
  assert.equal(cleanText("  Москва ", true), "Москва");
  assert.equal(cleanText(`Омск${NBSP}`, true), "Омск");
  assert.equal(cleanText("Казань   Север", true), "Казань Север");
  assert.equal(cleanText("Казань   Север", false), "Казань   Север", "внутренние — только по просьбе");
  assert.equal(cleanText(" 007 ", true), "007", "код остаётся текстом — это решает запись, не модель");
});

test("a number is taken from text only when the separators are certain", () => {
  assert.deepEqual(numberFromText("1 200,50", RU), { value: 1200.5 });
  assert.deepEqual(numberFromText("1,500", RU), { skip: "ambiguousNumber" });
  assert.deepEqual(numberFromText("1.5", RU), { skip: "foreignNumber" });
  assert.deepEqual(numberFromText("007", RU), { skip: "leadingZeros" });
  assert.deepEqual(numberFromText("Москва", RU), { skip: "notNumber" });
  // Пользователь ответил, какой разделитель десятичный.
  assert.deepEqual(numberFromText("1,500", RU, "."), { value: 1500 });
  assert.deepEqual(numberFromText("1,500", RU, ","), { value: 1.5 });
  assert.deepEqual(numberFromText("1.5", RU, "."), { value: 1.5 });
  assert.deepEqual(numberFromText("007", RU, ","), { skip: "leadingZeros" }, "код не становится числом и по прямой просьбе");
});

test("a date is taken from text only when day and month are certain", () => {
  assert.deepEqual(dateFromText("25.02.2026"), { date: { year: 2026, month: 2, day: 25 } });
  assert.deepEqual(dateFromText("2026-03-05"), { date: { year: 2026, month: 3, day: 5 } });
  assert.deepEqual(dateFromText("01.02.2026"), { skip: "ambiguousDate" });
  assert.deepEqual(dateFromText("01.02.2026", "DMY"), { date: { year: 2026, month: 2, day: 1 } });
  assert.deepEqual(dateFromText("01.02.2026", "MDY"), { date: { year: 2026, month: 1, day: 2 } });
  assert.deepEqual(dateFromText("25.02.2026", "MDY"), { skip: "contradictsOrder" });
  assert.deepEqual(dateFromText("2026-03-05", "MDY"), { date: { year: 2026, month: 3, day: 5 } }, "год впереди — порядок не нужен");
  assert.deepEqual(dateFromText("01.02.26"), { skip: "notDate" });
});

test("date serials and formats match what Excel showed", () => {
  // Замер 24 сентября 2026 года: 46054 с форматом dd.mm.yyyy — «01.02.2026».
  assert.equal(excelSerial({ year: 2026, month: 2, day: 1 }), 46054);
  assert.equal(excelSerial({ year: 1900, month: 3, day: 1 }), 61);
  assert.equal(dateFormatCode("ДД.ММ.ГГГГ"), "dd.mm.yyyy");
  assert.equal(dateFormatCode("M/d/yyyy"), "m/d/yyyy");
  assert.equal(dateFormatCode("странное"), "yyyy-mm-dd");
  assert.equal(formatDate({ year: 2026, month: 2, day: 1 }, "dd.mm.yyyy"), "01.02.2026");
  assert.equal(formatDate({ year: 2026, month: 2, day: 1 }, "m/d/yyyy"), "2/1/2026");
});
