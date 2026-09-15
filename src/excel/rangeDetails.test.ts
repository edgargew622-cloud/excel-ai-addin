import test from "node:test";
import assert from "node:assert/strict";
import { MIXED_FORMAT, markMixed, mergedAreasTouching } from "./excelTools";

test("mixed properties are named explicitly instead of looking unset", () => {
  const mixed: string[] = [];
  // Так Office.js отвечает по строке 1, где A1:H1 полужирные, а K1:L1 нет.
  const font = markMixed({ name: "Calibri", size: 11, bold: null, italic: false, color: null }, "font.", mixed);

  assert.equal(font.bold, MIXED_FORMAT);
  assert.equal(font.color, MIXED_FORMAT);
  assert.deepEqual(mixed, ["font.bold", "font.color"]);
  // Однородные свойства не искажаются, включая ложь и ноль.
  assert.equal(font.italic, false);
  assert.equal(font.size, 11);
});

test("uniform range reports nothing as mixed", () => {
  const mixed: string[] = [];
  const font = markMixed({ name: "Calibri", size: 11, bold: true, italic: false }, "font.", mixed);
  assert.deepEqual(mixed, []);
  assert.equal(font.bold, true);
});

test("prefix keeps property paths distinguishable across format groups", () => {
  const mixed: string[] = [];
  markMixed({ horizontalAlignment: null }, "", mixed);
  markMixed({ locked: null }, "protection.", mixed);
  assert.deepEqual(mixed, ["horizontalAlignment", "protection.locked"]);
});

test("merge covering a single cell is reported instead of looking absent", () => {
  // Опрос окрестности возвращает объединение целиком; запрошена одна ячейка L1.
  const found = mergedAreasTouching(["Продажи!K1:L1"], "Продажи!L1");
  assert.deepEqual(found, ["Продажи!K1:L1"]);
  // И для якоря объединения тоже — прежде оба случая давали пустой список.
  assert.deepEqual(mergedAreasTouching(["Продажи!K1:L1"], "Продажи!K1"), ["Продажи!K1:L1"]);
});

test("merges outside the requested range are filtered out", () => {
  const near = ["Продажи!K1:L1", "Продажи!A10:B10", "Продажи!Z1:Z2"];
  assert.deepEqual(mergedAreasTouching(near, "Продажи!K1:M1"), ["Продажи!K1:L1"]);
  assert.deepEqual(mergedAreasTouching(near, "Продажи!M1:N1"), []);
  assert.deepEqual(mergedAreasTouching(near, "Продажи!K1:K3"), ["Продажи!K1:L1"]);
});

test("full merge bounds survive instead of collapsing to the anchor", () => {
  const found = mergedAreasTouching(["Лист1!B2:D5"], "Лист1!C3");
  assert.deepEqual(found, ["Лист1!B2:D5"]);
});

test("addresses without a sheet prefix and duplicates are handled", () => {
  assert.deepEqual(mergedAreasTouching(["K1:L1", " K1:L1 "], "L1"), ["K1:L1"]);
  assert.deepEqual(mergedAreasTouching(["не адрес", ""], "L1"), []);
});
