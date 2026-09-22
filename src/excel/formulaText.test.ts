import test from "node:test";
import assert from "node:assert/strict";
import { canonicalFormula, sameCellContent, sameCellMatrix } from "./formulaText";

test("a sheet name Excel quoted on its own is the same formula", () => {
  // Проверка в Excel 23 сентября 2026 года: =Q1!A1 сохранилось как ='Q1'!A1.
  assert.equal(sameCellContent("='Q1'!A1", "=Q1!A1"), true);
  assert.equal(sameCellContent("='Q1 2026'!A1", "='Q1 2026'!A1"), true);
  assert.equal(sameCellContent("='Итоги.2026'!B2", "=Итоги.2026!B2"), true);
});

test("case outside text does not matter, Excel writes sum as SUM", () => {
  assert.equal(sameCellContent("=SUM(A1:A3)", "=sum(a1:a3)"), true);
  assert.equal(sameCellContent("=SUM(Sales[Q1])", "=sum(sales[q1])"), true);
});

test("text in quotes is compared as is", () => {
  assert.equal(sameCellContent('=IF(A1="да",1,0)', '=IF(A1="ДА",1,0)'), false);
  // Удвоенная кавычка внутри строки не обрывает строку.
  assert.equal(sameCellContent('=A1&"он сказал ""q1"""', '=a1&"он сказал ""q1"""'), true);
  assert.equal(sameCellContent('=A1&"он сказал ""q1"""', '=A1&"он сказал ""Q1"""'), false);
});

test("a different reference is still a different formula", () => {
  assert.equal(sameCellContent("='Q1'!A1", "=Q2!A1"), false);
  assert.equal(sameCellContent("=Q1!A1", "=Q1!A2"), false);
  // Кавычки, без которых не обойтись, не снимаются: иначе пробел слипся бы с адресом.
  assert.equal(canonicalFormula("='Q1 2026'!A1"), "='Q1 2026'!A1");
});

test("plain values are compared strictly, as before", () => {
  assert.equal(sameCellContent("Товар", "товар"), false, "значение Excel не переписывает");
  assert.equal(sameCellContent(10, 10), true);
  assert.equal(sameCellContent("", undefined), true);
  assert.equal(sameCellMatrix([["='Q1'!A1", 5]], [["=Q1!A1", 5]]), true);
  assert.equal(sameCellMatrix([["='Q1'!A1"]], [["=Q1!A1", 5]]), false, "разный размер — не одно и то же");
});

/* --- полный путь записи ---------------------------------------------------- */

/** Макет ячейки, которая хранит формулу так, как её переписывает Excel. */
function quotingExcel() {
  let stored: unknown = "";
  const range: any = {
    address: "Данные!E2",
    rowCount: 1, columnCount: 1, rowIndex: 1, columnIndex: 4,
    load: () => undefined,
    get formulas() { return [[stored]]; },
    set formulas(matrix: unknown[][]) {
      // Так делает Excel: имя листа, похожее на адрес, берёт в кавычки,
      // а имена функций пишет заглавными.
      stored = String(matrix[0][0]).replace(/=Q1!/g, "='Q1'!").replace(/sum\(/g, "SUM(");
    },
    get values() { return [[10]]; },
    get valueTypes() { return [["Double"]]; },
    getCell: () => range
  };
  const sheet: any = { id: "sheet-1", name: "Данные", load: () => undefined, getRange: () => range };
  (globalThis as any).Office = { context: { document: { url: "C:/q.xlsx" }, requirements: { isSetSupported: () => true } } };
  (globalThis as any).Excel = {
    run: async (fn: any) => fn({
      workbook: {
        application: { calculationMode: "automatic", load: () => undefined },
        worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet }
      },
      sync: async () => undefined
    })
  };
  return { writes: () => stored };
}

test("a formula Excel rewrote in its own spelling is verified, not reported as a mismatch", async () => {
  quotingExcel();
  const { executeSetRangePlan, prepareSetRangePlan } = await import("./excelTools");
  const plan = await prepareSetRangePlan({ sheet: "Данные", address: "E2", values: [["=Q1!A1"]], isFormula: true });
  const result = await executeSetRangePlan(plan) as any;
  assert.equal(result.executionState, "verified");
});
