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

/* --- столбцы и ссылки на целые столбцы и строки (этап 7, 7.3.2) ------------------ */

import { deleteImpact as bandDelete, insertBlindSpots as bandInsert, formulaReferences as references } from "./rowOps";
import { columnNumber } from "./columnPlans";

test("whole columns and whole rows are references too", () => {
  const [column] = references("=SUM($B:$D)");
  assert.deepEqual([column.columnStart, column.columnEnd, column.wholeColumns], [2, 4, true]);
  const [row] = references("=SUM(2:2)");
  assert.deepEqual([row.rowStart, row.rowEnd, row.wholeRows], [2, 2, true]);
  assert.equal(references('=LOG10(5)&"C:C"').length, 0, "имя функции и текст — не ссылки");
});

test("deleting column C does to the formulas what Excel did in the measurement", () => {
  // Замер 24 сентября 2026 года, удаление столбца C:
  // =C2*10 и =SUM(C:C) → #ССЫЛКА!; SUM(B:D) → B:C, SUM(B2:C2) → B2:B2, SUM(A2:E2) → A2:D2 — сужение;
  // SUM(2:2) — целая строка, столбцы её не задевают.
  const band = { startRow: 3, endRow: 3 };
  const kinds = (formula: string) => {
    const impact = bandDelete(formula, "Л", "Л", band, "columns");
    return impact.broken.length ? "broken" : impact.shrunk.length ? "shrunk" : "none";
  };
  assert.equal(kinds("=C2*10"), "broken");
  assert.equal(kinds("=SUM(C:C)"), "broken");
  assert.equal(kinds("=SUM(B:D)"), "shrunk");
  assert.equal(kinds("=SUM(B2:C2)"), "shrunk");
  assert.equal(kinds("=SUM(A2:E2)"), "shrunk");
  assert.equal(kinds("=SUM(2:2)"), "none");
  // А для строк целый столбец не страдает, целая строка — ломается.
  assert.equal(bandDelete("=SUM(C:C)", "Л", "Л", { startRow: 2, endRow: 2 }).broken.length + bandDelete("=SUM(C:C)", "Л", "Л", { startRow: 2, endRow: 2 }).shrunk.length, 0);
  assert.equal(bandDelete("=SUM(2:2)", "Л", "Л", { startRow: 2, endRow: 2 }).broken.length, 1);
});

test("a column inserted right after a sum's range is not covered, one inside is", () => {
  // Замер: вставка F за данными A:E не расширила SUM(A2:E2), вставка C внутрь B:C — расширила.
  assert.equal(bandInsert("=SUM(A2:E2)", "Л", "Л", { startRow: 6, endRow: 6 }, "columns").length, 1);
  assert.equal(bandInsert("=SUM(B:D)", "Л", "Л", { startRow: 3, endRow: 3 }, "columns").length, 0);
});

test("column letters are read strictly", () => {
  assert.equal(columnNumber("C"), 3);
  assert.equal(columnNumber("aa"), 27);
  assert.equal(columnNumber("XFD"), 16_384);
  assert.equal(columnNumber("XFE"), null);
  assert.equal(columnNumber("C1"), null);
});
