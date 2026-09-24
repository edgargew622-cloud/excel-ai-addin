import test from "node:test";
import assert from "node:assert/strict";
import { describeRule, executeMoveRulePlan, listConditionalFormats, movedOrder, prepareMoveRulePlan } from "./ruleOrderPlans";
import { PLANNED_TOOLS } from "./plans";
import { clear as clearUndo, setUndoMonitorReady, undoLast } from "./undo";

test("a rule read from Excel is described in plain words", () => {
  assert.equal(describeRule({ id: "0", type: "CellValue", priority: 0, range: "Л!C2:C5", rule: { formula1: "=1000", operator: "GreaterThan" }, fill: "#FF0000" }), "значение больше 1000 → заливка #FF0000");
  assert.equal(describeRule({ id: "1", type: "Custom", priority: 1, range: "Л!A2:C5", rule: { formula: '=$B2="Новая"' }, fill: "#FFC7CE", bold: true }), 'формула =$B2="Новая" → заливка #FFC7CE, жирный');
  assert.equal(describeRule({ id: "2", type: "ColorScale", priority: 2, range: null }), "цветовая шкала");
});

test("moving to first or last keeps the order of the others", () => {
  assert.deepEqual(movedOrder(["a", "b", "c", "d"], 2, "first"), ["c", "a", "b", "d"]);
  assert.deepEqual(movedOrder(["a", "b", "c", "d"], 1, "last"), ["a", "c", "d", "b"]);
});

test("rule order goes through the plan registry", () => {
  assert.ok(PLANNED_TOOLS.includes("move_conditional_format"));
});

/** Правила области так, как их отдаёт Excel по замеру: ID и приоритет — место в списке. */
function sheetWithRules(fills: string[], options: { priorityIgnored?: boolean } = {}) {
  const rules: any[] = [];
  const makeRule = (fill: string) => {
    const rule: any = {
      type: "CellValue",
      get id() { return String(rules.indexOf(rule)); },
      get priority() { return rules.indexOf(rule); },
      set priority(value: number) {
        if (options.priorityIgnored || value >= rules.length) return;
        rules.splice(rules.indexOf(rule), 1);
        rules.splice(value, 0, rule);
      },
      load: () => undefined,
      getRangeOrNullObject: () => ({ isNullObject: false, address: "Лист!C2:C5", load: () => undefined }),
      cellValue: {
        rule: { formula1: "=0", operator: "GreaterThan" },
        format: { fill: { color: fill, load: () => undefined }, font: { color: null, bold: null, load: () => undefined } },
        load: () => undefined
      }
    };
    return rule;
  };
  for (const fill of fills) rules.push(makeRule(fill));
  const range: any = {
    address: "Лист!C2:C5", rowCount: 4, columnCount: 1, load: () => undefined,
    format: { protection: { locked: false, load: () => undefined } },
    conditionalFormats: {
      get items() { return rules; },
      load: () => undefined,
      getItem: (id: string) => rules[Number(id)]
    }
  };
  const sheet: any = { id: "s1", name: "Лист", load: () => undefined, protection: { protected: false, load: () => undefined }, getRange: () => range };
  (globalThis as any).Office = { context: { document: { url: "C:/r.xlsx" }, requirements: { isSetSupported: () => true } } };
  (globalThis as any).Excel = { run: async (fn: any) => fn({ workbook: { worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet } }, sync: async () => undefined }) };
  return { rules, makeRule, fillsNow: () => rules.map((rule) => rule.cellValue.format.fill.color) };
}

test("the list numbers rules by priority for the move tool", async () => {
  sheetWithRules(["#FFFF00", "#FF0000"]);
  const list = await listConditionalFormats({ sheet: "Лист", address: "C2:C5" }) as any;
  assert.deepEqual(list.rules.map((rule: any) => rule.position), [1, 2]);
  assert.match(list.rules[1].rule, /#FF0000/);
});

test("a hidden rule is moved to the top, verified, and undo puts it back", async () => {
  setUndoMonitorReady(true);
  try {
    const book = sheetWithRules(["#FFFF00", "#00B050", "#FF0000"]);
    const plan = await prepareMoveRulePlan({ sheet: "Лист", address: "C2:C5", position: 3, to: "first" });
    assert.match(plan.ruleText, /#FF0000/);
    const result = await executeMoveRulePlan(plan) as any;
    assert.equal(result.executionState, "verified");
    assert.deepEqual(book.fillsNow(), ["#FF0000", "#FFFF00", "#00B050"]);
    await undoLast();
    assert.deepEqual(book.fillsNow(), ["#FFFF00", "#00B050", "#FF0000"]);
  } finally {
    clearUndo();
    setUndoMonitorReady(false);
  }
});

test("a priority Excel ignored is not reported as moved", async () => {
  sheetWithRules(["#FFFF00", "#FF0000"], { priorityIgnored: true });
  await assert.rejects(
    async () => executeMoveRulePlan(await prepareMoveRulePlan({ sheet: "Лист", address: "C2:C5", position: 2, to: "first" })),
    (error: any) => error.executionState === "applied"
  );
});

test("a position outside the list or already in place is refused before Excel", async () => {
  sheetWithRules(["#FFFF00", "#FF0000"]);
  await assert.rejects(() => prepareMoveRulePlan({ sheet: "Лист", address: "C2:C5", position: 3, to: "first" }), /от 1 до 2/);
  await assert.rejects(() => prepareMoveRulePlan({ sheet: "Лист", address: "C2:C5", position: 1, to: "first" }), /уже стоит первым/);
});

test("undo stops if the rules were changed after the move", async () => {
  setUndoMonitorReady(true);
  try {
    const book = sheetWithRules(["#FFFF00", "#FF0000"]);
    await executeMoveRulePlan(await prepareMoveRulePlan({ sheet: "Лист", address: "C2:C5", position: 2, to: "first" }));
    book.rules.push(book.makeRule("#0000FF"));
    await assert.rejects(() => undoLast(), /изменили после операции агента/);
    assert.deepEqual(book.fillsNow(), ["#FF0000", "#FFFF00", "#0000FF"]);
  } finally {
    clearUndo();
    setUndoMonitorReady(false);
  }
});
