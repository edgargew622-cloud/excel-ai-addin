import test from "node:test";
import assert from "node:assert/strict";
import { executeGroupPlan, parseOutlineBand, prepareGroupPlan } from "./outlinePlans";
import { PLANNED_TOOLS } from "./plans";
import { clear as clearUndo, setUndoMonitorReady, undoLast } from "./undo";

/**
 * Лист со строками 1–20, который ведёт себя как Excel в замере 24 сентября
 * 2026 года: группа добавляет уровень; свёртка скрывает только строки
 * с уровнем; разворот той же полосы видимость не возвращает; разгруппировка
 * снимает уровень, но строки оставляет скрытыми.
 */
function outlineSheet(options: { hidden?: number[]; groupIgnored?: boolean } = {}) {
  const level = new Map<number, number>();
  const hidden = new Set<number>(options.hidden ?? []);
  const lines = (address: string) => {
    const [a, b] = address.split(":").map(Number);
    return Array.from({ length: b - a + 1 }, (_, index) => a + index);
  };
  const sheet: any = {
    id: "s1", name: "Лист", load: () => undefined,
    protection: { protected: false, load: () => undefined },
    getRange: (address: string) => ({
      load: () => undefined,
      get rowHidden() { return hidden.has(lines(address)[0]); },
      set rowHidden(value: boolean) { for (const n of lines(address)) value ? hidden.add(n) : hidden.delete(n); },
      group: () => { if (!options.groupIgnored) for (const n of lines(address)) level.set(n, (level.get(n) ?? 0) + 1); },
      ungroup: () => { for (const n of lines(address)) level.set(n, Math.max(0, (level.get(n) ?? 0) - 1)); },
      hideGroupDetails: () => { for (const n of lines(address)) if ((level.get(n) ?? 0) > 0) hidden.add(n); },
      showGroupDetails: () => undefined
    })
  };
  (globalThis as any).Office = { context: { document: { url: "C:/outline.xlsx" }, requirements: { isSetSupported: () => true } } };
  (globalThis as any).Excel = {
    run: async (fn: any) => fn({ workbook: { worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet } }, sync: async () => undefined })
  };
  return { level, hidden };
}

test("grouping goes through the plan registry", () => {
  assert.ok(PLANNED_TOOLS.includes("group_rows_columns"));
});

test("a band is whole rows or whole columns, nothing else", () => {
  assert.deepEqual(parseOutlineBand("5:3"), { axis: "rows", start: 3, end: 5, address: "3:5" });
  assert.deepEqual(parseOutlineBand("$c:$d"), { axis: "columns", start: 3, end: 4, address: "C:D" });
  assert.throws(() => parseOutlineBand("A1:B3"), /целые строки/);
});

test("a group is proven by collapsing it, and visibility comes back exactly as it was", async () => {
  // Строка 4 была скрыта пользователем до операции — такой и останется.
  const state = outlineSheet({ hidden: [4] });
  const result = await executeGroupPlan(await prepareGroupPlan({ sheet: "Лист", address: "3:5" })) as any;
  assert.equal(result.executionState, "verified");
  assert.deepEqual([3, 4, 5].map((n) => state.level.get(n)), [1, 1, 1]);
  assert.deepEqual([3, 4, 5].map((n) => state.hidden.has(n)), [false, true, false], "разворот Excel видимость не вернул бы — вернула панель");
});

test("collapse leaves the group folded", async () => {
  const state = outlineSheet();
  await executeGroupPlan(await prepareGroupPlan({ sheet: "Лист", address: "3:5", collapse: true }));
  assert.deepEqual([2, 3, 4, 5, 6].map((n) => state.hidden.has(n)), [false, true, true, true, false]);
});

test("a group Excel did not create is caught by the collapse, not reported as done", async () => {
  outlineSheet({ groupIgnored: true });
  await assert.rejects(async () => executeGroupPlan(await prepareGroupPlan({ sheet: "Лист", address: "3:5" })), (error: any) => {
    assert.equal(error.executionState, "applied");
    assert.match(error.message, /группа, похоже, не создалась/);
    return true;
  });
});

test("rows that were all hidden already cannot prove the group, and it says so", async () => {
  outlineSheet({ hidden: [3, 4, 5] });
  const result = await executeGroupPlan(await prepareGroupPlan({ sheet: "Лист", address: "3:5" })) as any;
  assert.equal(result.executionState, "applied");
  assert.match(result.note, /не проверена/);
});

test("undo removes our level and brings visibility back, even though ungroup alone would not", async () => {
  const state = outlineSheet();
  setUndoMonitorReady(true);
  try {
    await executeGroupPlan(await prepareGroupPlan({ sheet: "Лист", address: "3:5", collapse: true }));
    await undoLast();
    assert.deepEqual([3, 4, 5].map((n) => state.level.get(n)), [0, 0, 0]);
    assert.deepEqual([3, 4, 5].map((n) => state.hidden.has(n)), [false, false, false]);
  } finally {
    clearUndo();
    setUndoMonitorReady(false);
  }
});

test("visibility changed after the preview stops the grouping", async () => {
  const state = outlineSheet();
  const plan = await prepareGroupPlan({ sheet: "Лист", address: "3:5" });
  state.hidden.add(4);
  await assert.rejects(() => executeGroupPlan(plan), (error: any) => error.executionState === "failed_before_write");
  assert.equal(state.level.get(3), undefined, "группа не создана");
});
