import test from "node:test";
import assert from "node:assert/strict";
import { MIXED_FORMAT, markMixed, mergedAddressCandidates, mergedAreasTouching } from "./excelTools";

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

test("a truncated anchor is reported as unresolved, not as absence of merges", () => {
  // Office.js на замеренной сборке отдаёт объединение N1:P1 как один угол N1.
  const report = mergedAreasTouching(["Продажи!N1"], "Продажи!O1");
  assert.deepEqual(report.areas, []);
  // Прежде здесь был пустой список, и ячейка внутри объединения выглядела обычной.
  assert.deepEqual(report.unresolvedAnchors, ["Продажи!N1"]);
});

test("an anchor to the right or below cannot cover the target", () => {
  // Объединение растёт вправо и вниз, поэтому такие углы цель не накрывают.
  assert.deepEqual(mergedAreasTouching(["Продажи!P5"], "Продажи!O1").unresolvedAnchors, []);
  assert.deepEqual(mergedAreasTouching(["Продажи!N3"], "Продажи!O1").unresolvedAnchors, []);
  assert.deepEqual(mergedAreasTouching(["Продажи!N1"], "Продажи!N1").unresolvedAnchors, ["Продажи!N1"]);
});

test("real bounds, when Excel reports them, are trusted and filtered by overlap", () => {
  const near = ["Лист1!B2:D5", "Лист1!H1:J1"];
  assert.deepEqual(mergedAreasTouching(near, "Лист1!C3").areas, ["Лист1!B2:D5"]);
  assert.deepEqual(mergedAreasTouching(near, "Лист1!A1").areas, []);
});

test("addresses without a sheet prefix and duplicates are handled", () => {
  const report = mergedAreasTouching(["K1:L1", " K1:L1 "], "L1");
  assert.deepEqual(report.areas, ["K1:L1"]);
  assert.deepEqual(mergedAreasTouching(["не адрес", ""], "L1").areas, []);
});

test("merge addresses are taken from whichever Office.js source is filled", () => {
  assert.deepEqual(
    mergedAddressCandidates({ address: "Лист1!N1:P2", areas: { items: [] } }),
    ["Лист1!N1:P2"]
  );
  assert.deepEqual(
    mergedAddressCandidates({ address: "", areas: { items: [{ address: "Лист1!N1:P2" }] } }),
    ["Лист1!N1:P2"]
  );
  assert.deepEqual(
    mergedAddressCandidates({ address: "Продажи!K1,Продажи!N1", areas: null }),
    ["Продажи!K1", "Продажи!N1"]
  );
  assert.deepEqual(mergedAddressCandidates({ address: null, areas: null }), []);
});
