import test from "node:test";
import assert from "node:assert/strict";
import {
  countFilled,
  countRefErrors,
  deleteImpact,
  formulaReferences,
  insertBlindSpots,
  rowBand,
  usesTableReference
} from "./rowOps";

test("references are read out of a formula with their sheet and rows", () => {
  assert.deepEqual(formulaReferences("=D2*E2").map((r) => r.text), ["D2", "E2"]);
  const cross = formulaReferences("=SUM(Продажи!A2:A10)");
  assert.equal(cross.length, 1);
  assert.equal(cross[0].sheet, "Продажи");
  assert.equal(cross[0].rowStart, 2);
  assert.equal(cross[0].rowEnd, 10);
  // Имя листа с пробелом Excel берёт в апострофы.
  assert.equal(formulaReferences("='Отчёт за год'!B3")[0].sheet, "Отчёт за год");
  // Не формула — ссылок нет.
  assert.deepEqual(formulaReferences("просто текст"), []);
  assert.deepEqual(formulaReferences(42), []);
});

test("text in quotes and function names are not mistaken for references", () => {
  assert.deepEqual(formulaReferences('=IF(A1>0,"B2 штук","")').map((r) => r.text), ["A1"]);
  assert.deepEqual(formulaReferences("=LOG10(A1)").map((r) => r.text), ["A1"]);
});

test("deleting rows tells apart a broken reference from a silently shrunk one", () => {
  const band = rowBand(5, 2); // удаляем строки 5 и 6
  const broken = deleteImpact("=Продажи!B5", "Отчёт", "Продажи", band);
  assert.equal(broken.broken.length, 1, "ссылка целиком внутри удаляемых строк");
  assert.equal(broken.shrunk.length, 0);

  const shrunk = deleteImpact("=SUM(Продажи!B2:B10)", "Отчёт", "Продажи", band);
  assert.equal(shrunk.broken.length, 0);
  assert.equal(shrunk.shrunk.length, 1, "диапазон уменьшится, итог изменится молча");

  // Совсем в стороне — не задето.
  assert.equal(deleteImpact("=SUM(Продажи!B20:B30)", "Отчёт", "Продажи", band).shrunk.length, 0);
  // Другой лист — не задето.
  assert.equal(deleteImpact("=Склад!B5", "Отчёт", "Продажи", band).broken.length, 0);
  // Ссылка без имени листа принадлежит листу самой формулы.
  assert.equal(deleteImpact("=B5", "Продажи", "Продажи", band).broken.length, 1);
  assert.equal(deleteImpact("=B5", "Отчёт", "Продажи", band).broken.length, 0);
  // Доллары от удаления не спасают: ячейки больше нет.
  assert.equal(deleteImpact("=Продажи!$B$5", "Отчёт", "Продажи", band).broken.length, 1);
});

test("inserting rows finds the ranges that will not cover them", () => {
  const band = rowBand(11, 1); // вставляем строку 11
  // Классика: сумма по A2:A10, строка вставлена сразу под ней.
  assert.equal(insertBlindSpots("=SUM(A2:A10)", "Продажи", "Продажи", band).length, 1);
  // Вставка вплотную сверху — Excel тоже не расширяет диапазон вверх.
  assert.equal(insertBlindSpots("=SUM(A11:A20)", "Продажи", "Продажи", band).length, 1);
  // Вставка внутрь диапазона — Excel расширит его сам, предупреждать не о чем.
  assert.equal(insertBlindSpots("=SUM(A2:A20)", "Продажи", "Продажи", band).length, 0);
  // Одиночная ячейка просто съезжает вниз.
  assert.equal(insertBlindSpots("=A10", "Продажи", "Продажи", band).length, 0);
});

test("formulas with table references are flagged as not analysed", () => {
  assert.ok(usesTableReference("=SUM(SalesTable[Выручка])"));
  assert.ok(!usesTableReference("=SUM(A2:A10)"));
});

test("reference errors are counted in both languages of Excel", () => {
  assert.equal(countRefErrors([["#REF!", 1], ["ок", "#ССЫЛКА!"]]), 2);
  assert.equal(countRefErrors([[1, "", null]]), 0);
});

test("filled cells of the band are counted, because deletion loses them", () => {
  assert.equal(countFilled([["", null, 0], ["текст", undefined, ""]]), 2);
});
