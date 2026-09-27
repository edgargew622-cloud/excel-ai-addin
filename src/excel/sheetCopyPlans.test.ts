import test from "node:test";
import assert from "node:assert/strict";
import { executeCopySheetPlan, prepareCopySheetPlan } from "./sheetCopyPlans";
import { PLANNED_TOOLS } from "./plans";
import { clear as clearUndo, setUndoMonitorReady, undoLast } from "./undo";

/**
 * Книга, где copy() ведёт себя как Excel в замере 28 сентября 2026 года:
 * копия встаёт за исходным, называется «Лист (2)» и несёт значения, формулы
 * и условное форматирование. «Порча» — копия без одного правила.
 */
function workbook(options: { dropRule?: boolean } = {}) {
  const sheets: any[] = [];
  const makeSheet = (name: string, data: { values: unknown[][]; formulas: unknown[][]; rules: number; charts: number }) => {
    const sheet: any = {
      id: `id-${name}`,
      name,
      get position() { return sheets.indexOf(sheet); },
      data,
      load: () => undefined,
      tables: { items: [], load: () => undefined },
      charts: { get items() { return Array.from({ length: sheet.data.charts }, (_, i) => ({ name: `Диаграмма ${i + 1}` })); }, load: () => undefined },
      pivotTables: { items: [], load: () => undefined },
      getUsedRangeOrNullObject: () => ({
        isNullObject: false,
        address: `${sheet.name}!A1:B2`,
        rowCount: 2,
        columnCount: 2,
        get values() { return sheet.data.values; },
        get formulas() { return sheet.data.formulas; },
        conditionalFormats: { get items() { return Array.from({ length: sheet.data.rules }, () => ({ type: "IconSet" })); }, load: () => undefined },
        load: () => undefined
      }),
      copy: (position: string, after?: any) => {
        const copyData = { ...sheet.data, values: sheet.data.values.map((row: unknown[]) => [...row]), rules: options.dropRule ? sheet.data.rules - 1 : sheet.data.rules };
        const created = makeSheet(`${sheet.name} (2)`, copyData);
        const index = position === "End" ? sheets.length : sheets.indexOf(after) + 1;
        sheets.splice(index, 0, created);
        return created;
      },
      delete: () => { sheets.splice(sheets.indexOf(sheet), 1); }
    };
    return sheet;
  };
  const source = makeSheet("Отчёт", { values: [["Выручка", 100], ["Расходы", 60]], formulas: [["Выручка", "=50+50"], ["Расходы", 60]], rules: 1, charts: 1 });
  sheets.push(source, makeSheet("Итоги", { values: [["", ""], ["", ""]], formulas: [["", ""], ["", ""]], rules: 0, charts: 0 }));
  const byKey = (key: string) => sheets.find((item) => item.id === key || item.name === key);
  (globalThis as any).Office = { context: { document: { url: "C:/copy.xlsx" }, requirements: { isSetSupported: () => true } } };
  (globalThis as any).Excel = {
    run: async (fn: any) => fn({
      workbook: {
        worksheets: {
          getActiveWorksheet: () => source,
          getItem: (key: string) => byKey(key),
          getItemOrNullObject: (key: string) => byKey(key) ?? { isNullObject: true, load: () => undefined },
          get items() { return sheets; },
          load: () => undefined
        }
      },
      sync: async () => undefined
    })
  };
  return { sheets, source };
}

test("copy_sheet goes through the plan registry", () => {
  assert.ok(PLANNED_TOOLS.includes("copy_sheet"));
});

test("a sheet is copied next to itself under the asked name, checked against the source, and undone", async () => {
  const book = workbook();
  setUndoMonitorReady(true);
  try {
    const plan = await prepareCopySheetPlan({ sheet: "Отчёт", newName: "Отчёт — копия" });
    assert.ok(plan.warnings.some((text) => /Формул: 1/.test(text)));
    const result = await executeCopySheetPlan(plan) as any;
    assert.equal(result.executionState, "verified");
    assert.equal(result.copy, "Отчёт — копия");
    assert.deepEqual(book.sheets.map((item) => item.name), ["Отчёт", "Отчёт — копия", "Итоги"]);
    await undoLast();
    assert.deepEqual(book.sheets.map((item) => item.name), ["Отчёт", "Итоги"]);
  } finally {
    clearUndo();
    setUndoMonitorReady(false);
  }
});

test("a copy that lost a conditional format rule is not reported as done", async () => {
  workbook({ dropRule: true });
  await assert.rejects(async () => executeCopySheetPlan(await prepareCopySheetPlan({ sheet: "Отчёт" })), (error: any) => {
    assert.equal(error.executionState, "applied");
    assert.match(error.message, /правил условного форматирования 0 вместо 1/);
    return true;
  });
});

test("a taken name is refused with a free one; a copy edited after the operation is not deleted by undo", async () => {
  workbook();
  await assert.rejects(() => prepareCopySheetPlan({ sheet: "Отчёт", newName: "итоги" }), /Свободно, например/);

  const book = workbook();
  setUndoMonitorReady(true);
  try {
    await executeCopySheetPlan(await prepareCopySheetPlan({ sheet: "Отчёт", position: "end" }));
    const copy = book.sheets[book.sheets.length - 1];
    assert.equal(copy.name, "Отчёт (2)");
    copy.data.values[0][1] = 999;
    await assert.rejects(() => undoLast(), /внесли изменения/);
    assert.equal(book.sheets.length, 3);
  } finally {
    clearUndo();
    setUndoMonitorReady(false);
  }
});
