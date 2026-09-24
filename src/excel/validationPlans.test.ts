import test from "node:test";
import assert from "node:assert/strict";
import {
  executeValidationPlan,
  officeRule,
  parseValidationRequest,
  predictInvalid,
  prepareValidationPlan,
  sameValidation
} from "./validationPlans";
import { PLANNED_TOOLS } from "./plans";
import { clear as clearUndo, setUndoMonitorReady, undoLast } from "./undo";

const LIST = parseValidationRequest({ rule: "list", items: ["Новая", "В работе", "Закрыта"] });

test("a request is checked before Excel sees it", () => {
  assert.throws(() => parseValidationRequest({ rule: "list", items: [] }), /нужны items/);
  assert.throws(() => parseValidationRequest({ rule: "list", items: ["А,Б"] }), /запятую/);
  assert.throws(() => parseValidationRequest({ rule: "wholeNumber", operator: "between", value: 1 }), /value2/);
  assert.throws(() => parseValidationRequest({ rule: "wholeNumber", operator: "greaterThan", value: 1.5 }), /целое/);
  assert.throws(() => parseValidationRequest({ rule: "date", operator: "greaterThan", value: "01.02.2026" }), /ГГГГ-ММ-ДД/);
  assert.deepEqual(officeRule(parseValidationRequest({ rule: "decimal", operator: "greaterThan", value: 0 })), { decimal: { formula1: 0, operator: "GreaterThan" } });
});

test("invalid values are predicted the way Excel judged them in the measurement", () => {
  // Замер: «Отменена» и «в работе» (другой регистр) не прошли список, пустое — прошло.
  assert.deepEqual(["Новая", "Закрыта", "Отменена", "в работе", ""].map((value) => predictInvalid(LIST, value)), [false, false, true, true, false]);
  // Число > 0: −5, 0 и текст «abc» не прошли.
  const positive = parseValidationRequest({ rule: "decimal", operator: "greaterThan", value: 0 });
  assert.deepEqual([100, -5, 300, 0, "abc"].map((value) => predictInvalid(positive, value)), [false, true, false, true, true]);
  // Дата в 2026 году: 40000 — это 2009 год.
  const year = parseValidationRequest({ rule: "date", operator: "between", value: "2026-01-01", value2: "2026-12-31" });
  assert.deepEqual([46054, 46100, 40000, ""].map((value) => predictInvalid(year, value)), [false, false, true, false]);
});

test("a rule read back from Excel is compared by meaning, not by text", () => {
  // Замер: дату Excel вернул в американском виде.
  const year = parseValidationRequest({ rule: "date", operator: "between", value: "2026-01-01", value2: "2026-12-31" });
  assert.equal(sameValidation(year, "Date", { date: { formula1: "1/1/2026", formula2: "12/31/2026", operator: "Between" } }), true);
  assert.equal(sameValidation(year, "Date", { date: { formula1: "1/2/2026", formula2: "12/31/2026", operator: "Between" } }), false);
  assert.equal(sameValidation(LIST, "List", { list: { source: "Новая,В работе,Закрыта" } }), true);
  assert.equal(sameValidation(LIST, "Decimal", { decimal: {} }), false);
});

/** Область, которая хранит правило и называет нарушителей, как Excel. */
function validationSheet(options: { previous?: { type: string; rule: any }; ruleIgnored?: boolean } = {}) {
  const state = { type: options.previous?.type ?? "None", rule: options.previous?.rule ?? null as any };
  const values = [["Новая"], ["Отменена"], ["в работе"], [""]];
  const dataValidation: any = {
    load: () => undefined,
    get type() { return state.type; },
    get rule() { return state.rule; },
    set rule(value: any) {
      if (options.ruleIgnored) return;
      state.rule = value;
      state.type = value.list ? "List" : value.wholeNumber ? "WholeNumber" : value.decimal ? "Decimal" : "Date";
    },
    clear: () => { state.type = "None"; state.rule = null; },
    getInvalidCellsOrNullObject: () => {
      const bad = state.rule?.list ? values.map((row, r) => (row[0] && !state.rule.list.source.split(",").includes(row[0]) ? `Лист!A${r + 2}` : null)).filter(Boolean) : [];
      return { isNullObject: bad.length === 0, address: bad.join(","), cellCount: bad.length, load: () => undefined };
    }
  };
  const range: any = {
    address: "Лист!A2:A5", rowIndex: 1, columnIndex: 0, rowCount: 4, columnCount: 1,
    load: () => undefined, values, dataValidation,
    format: { protection: { locked: false, load: () => undefined } }
  };
  const sheet: any = { id: "s1", name: "Лист", load: () => undefined, protection: { protected: false, load: () => undefined }, getRange: () => range };
  (globalThis as any).Office = { context: { document: { url: "C:/dv.xlsx" }, requirements: { isSetSupported: () => true } } };
  (globalThis as any).Excel = { run: async (fn: any) => fn({ workbook: { worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet } }, sync: async () => undefined }) };
  return state;
}

const ASK = { sheet: "Лист", address: "A2:A5", rule: "list", items: ["Новая", "В работе", "Закрыта"] };

test("data validation goes through the plan registry", () => {
  assert.ok(PLANNED_TOOLS.includes("set_data_validation"));
});

test("a rule is set, read back, and the values it does not accept are named by Excel", async () => {
  validationSheet();
  const plan = await prepareValidationPlan(ASK);
  assert.equal(plan.predictedInvalidCount, 2);
  const result = await executeValidationPlan(plan) as any;
  assert.equal(result.executionState, "verified");
  assert.deepEqual(result.invalidExamples, ["A3", "A4"]);
  assert.match(result.invalidNote, /не исправляет/);
});

test("a rule Excel did not take is not reported as set", async () => {
  validationSheet({ ruleIgnored: true });
  await assert.rejects(async () => executeValidationPlan(await prepareValidationPlan(ASK)), (error: any) => error.executionState === "applied");
});

test("undo brings the previous rule back, or clears when there was none", async () => {
  setUndoMonitorReady(true);
  try {
    const previous = { type: "Decimal", rule: { decimal: { formula1: "0", operator: "GreaterThan" } } };
    const state = validationSheet({ previous });
    await executeValidationPlan(await prepareValidationPlan(ASK));
    assert.equal(state.type, "List");
    await undoLast();
    assert.equal(state.type, "Decimal");
    assert.deepEqual(state.rule, previous.rule);

    const empty = validationSheet();
    await executeValidationPlan(await prepareValidationPlan(ASK));
    await undoLast();
    assert.equal(empty.type, "None");
  } finally {
    clearUndo();
    setUndoMonitorReady(false);
  }
});

test("mixed rules cannot be brought back, so there is no undo, and it says so", async () => {
  setUndoMonitorReady(true);
  try {
    validationSheet({ previous: { type: "Inconsistent", rule: {} } });
    const plan = await prepareValidationPlan(ASK);
    assert.equal(plan.undoAvailable, false);
    assert.match(plan.undoNote, /разные правила/);
  } finally {
    setUndoMonitorReady(false);
  }
});

test("a rule changed after the preview stops the operation", async () => {
  const state = validationSheet();
  const plan = await prepareValidationPlan(ASK);
  state.type = "WholeNumber";
  state.rule = { wholeNumber: { formula1: "1", operator: "GreaterThan" } };
  await assert.rejects(() => executeValidationPlan(plan), (error: any) => error.executionState === "failed_before_write");
});
