import test from "node:test";
import assert from "node:assert/strict";
import {
  executeAddSlicerPlan,
  executeRefreshPivotPlan,
  executeSheetViewPlan,
  executeUpdatePivotPlan,
  prepareAddSlicerPlan,
  prepareRefreshPivotPlan,
  prepareSheetViewPlan,
  prepareUpdatePivotPlan
} from "./pivotTools";
import { PLANNED_TOOLS } from "./plans";
import { clear as clearUndo, setUndoMonitorReady, undoLast } from "./undo";

/** Книга со сводной «СвВыручка» по полям Год, Регион, Категория, Выручка —
 * ведёт себя так, как Excel в замере 07.10.2026. */
function pivotBook() {
  const source = ["Год", "Регион", "Категория", "Выручка"];
  const state = {
    rows: ["Категория"] as string[],
    columns: [] as string[],
    filters: [] as string[],
    values: [] as any[],
    layout: { showRowGrandTotals: true, showColumnGrandTotals: true, autoFormat: true, load: () => undefined } as any,
    refreshed: 0,
    sorted: null as any,
    removedValues: 0,
    refreshedAll: 0,
    slicers: [] as any[],
    sheets: [
      { name: "Сводка", id: "s1", showGridlines: true, visibility: "Visible" },
      { name: "Данные", id: "s2", showGridlines: true, visibility: "Visible" }
    ] as any[]
  };
  const valueItem = (source: string) => ({ name: `Сумма по полю ${source}`, summarizeBy: "Sum", numberFormat: "Общий", field: { name: source, load: () => undefined } });
  state.values.push(valueItem("Выручка"));
  const sheetObject = (sheet: any) => Object.assign(sheet, { load: () => undefined });
  const hierarchyList = (list: string[]) => ({
    load: () => undefined,
    get items() { return list.map((name) => ({ name, fields: { getItem: () => ({ set subtotals(v: unknown) { state.layout.subtotals = v; }, sortByValues: (order: string, by: any) => { state.sorted = { field: name, order, by: by.name }; } }) } })); },
    add: (name: string) => { list.push(name); },
    remove: (item: any) => { list.splice(list.indexOf(item.name), 1); },
    getItem: (name: string) => ({ name, fields: { getItem: () => ({ set subtotals(v: unknown) { state.layout.subtotals = v; } }) } })
  });
  const pivot: any = {
    name: "СвВыручка",
    worksheet: sheetObject(state.sheets[0]),
    load: () => undefined,
    hierarchies: { load: () => undefined, get items() { return source.map((name) => ({ name })); }, getItem: (name: string) => name },
    rowHierarchies: hierarchyList(state.rows),
    columnHierarchies: hierarchyList(state.columns),
    filterHierarchies: hierarchyList(state.filters),
    dataHierarchies: {
      load: () => undefined,
      get items() { return state.values; },
      add: (field: string) => { const item = valueItem(field); state.values.push(item); return item; },
      remove: (item: any) => { state.removedValues += 1; state.values.splice(state.values.indexOf(item), 1); }
    },
    layout: Object.assign(state.layout, { getRange: () => ({ address: "Сводка!A3:B6", values: [["Категория", "Итог"]], load: () => undefined }) }),
    refresh: () => { state.refreshed += 1; }
  };
  const worksheets = {
    load: () => undefined,
    get items() { return state.sheets.map(sheetObject); },
    getItem: (key: string) => sheetObject(state.sheets.find((s) => s.id === key || s.name === key)),
    getItemOrNullObject: (key: string) => { const s = state.sheets.find((x) => x.name === key); return s ? sheetObject(s) : { isNullObject: true, load: () => undefined }; },
    getActiveWorksheet: () => sheetObject(state.sheets[0])
  };
  (globalThis as any).Office = { context: { document: { url: "C:/pivot.xlsx" }, requirements: { isSetSupported: () => true } } };
  (globalThis as any).Excel = {
    run: async (fn: any) => fn({
      workbook: {
        worksheets,
        pivotTables: { load: () => undefined, items: [pivot], refreshAll: () => { state.refreshedAll += 1; } },
        slicers: {
          load: () => undefined,
          get items() { return state.slicers; },
          add: (_pivot: unknown, field: string) => { const s: any = { name: field, load: () => undefined }; state.slicers.push(s); return s; },
          getItemOrNullObject: (name: string) => ({ delete: () => { state.slicers = state.slicers.filter((s) => s.name !== name); } })
        }
      },
      application: { cultureInfo: { numberFormat: { numberDecimalSeparator: ",", numberGroupSeparator: " ", load: () => undefined } } },
      sync: async () => undefined
    })
  };
  return state;
}

test("10.7 tools go through the plan registry", () => {
  for (const name of ["update_pivot", "refresh_pivot", "add_slicer", "set_sheet_view"]) assert.ok(PLANNED_TOOLS.includes(name), name);
});

test("update_pivot swaps the row field, names and formats the value field, drops totals — and undo brings it back", async () => {
  // Инструкция автора: «меняет поля (с категорий на заказчиков)», не строя сводную заново.
  const state = pivotBook();
  setUndoMonitorReady(true);
  try {
    const plan = await prepareUpdatePivotPlan({
      pivot: "свВыручка", rows: ["Регион"], grandTotals: "none",
      values: [{ field: "Выручка", label: "Выручка", numberFormat: '#,##0,"к"' }]
    });
    assert.ok(plan.preview.some((line) => /Категория → Регион/.test(line)));
    const result = await executeUpdatePivotPlan(plan) as any;
    assert.equal(result.executionState, "verified");
    assert.deepEqual(state.rows, ["Регион"]);
    assert.equal(state.values[0].name, "Выручка ", "совпадает с полем источника — с пробелом");
    assert.equal(state.values[0].numberFormat, '# ##0 "к"');
    assert.equal(state.layout.showRowGrandTotals, false);
    await undoLast();
    assert.deepEqual(state.rows, ["Категория"]);
    assert.equal(state.values[0].name, "Сумма по полю Выручка");
    assert.equal(state.layout.showRowGrandTotals, true);
  } finally {
    clearUndo();
    setUndoMonitorReady(false);
  }
});

test("update_pivot refuses fields the pivot does not have, and a change made after the preview", async () => {
  const state = pivotBook();
  await assert.rejects(() => prepareUpdatePivotPlan({ pivot: "СвВыручка", rows: ["Заказчик"] }), /нет полей «Заказчик».*Поля источника/s);
  await assert.rejects(() => prepareUpdatePivotPlan({ pivot: "Нет такой", rows: ["Регион"] }), /Сводной «Нет такой» нет\. Есть: «СвВыручка»/);
  const plan = await prepareUpdatePivotPlan({ pivot: "СвВыручка", rows: ["Регион"] });
  state.rows.push("Год");
  await assert.rejects(() => executeUpdatePivotPlan(plan), (error: any) => error.executionState === "failed_before_write");
});

test("refresh_pivot: one pivot by name or all of them", async () => {
  const state = pivotBook();
  await executeRefreshPivotPlan(await prepareRefreshPivotPlan({ pivot: "СвВыручка" }));
  assert.equal(state.refreshed, 1);
  const all = await executeRefreshPivotPlan(await prepareRefreshPivotPlan({})) as any;
  assert.equal(state.refreshedAll, 1);
  assert.equal(all.undoable, false);
});

test("add_slicer: slicers on another sheet, said honestly that they filter only their pivot; undo removes them", async () => {
  const state = pivotBook();
  setUndoMonitorReady(true);
  try {
    const plan = await prepareAddSlicerPlan({ pivot: "СвВыручка", fields: ["год", "Регион"], destSheet: "Данные" });
    assert.deepEqual(plan.fields, ["Год", "Регион"]);
    const result = await executeAddSlicerPlan(plan) as any;
    assert.deepEqual(result.slicers, ["Год", "Регион"]);
    assert.match(result.note, /Подключение к отчётам/);
    await assert.rejects(() => prepareAddSlicerPlan({ pivot: "СвВыручка", fields: ["Год"] }), /уже есть/);
    await undoLast();
    assert.equal(state.slicers.length, 0);
  } finally {
    clearUndo();
    setUndoMonitorReady(false);
  }
});

test("set_sheet_view: gridlines off and a sheet hidden; the last visible sheet is never hidden", async () => {
  const state = pivotBook();
  await executeSheetViewPlan(await prepareSheetViewPlan({ sheet: "Сводка", gridlines: false }));
  assert.equal(state.sheets[0].showGridlines, false);
  await executeSheetViewPlan(await prepareSheetViewPlan({ sheet: "Данные", visible: false }));
  assert.equal(state.sheets[1].visibility, "Hidden");
  await assert.rejects(() => prepareSheetViewPlan({ sheet: "Сводка", visible: false }), /единственный видимый лист/);
});

test("update_pivot: a new label and format keep the value field (and its sort); the order can be set again", async () => {
  // Живая проверка 10.7: агент «поправил» подпись «Выручка » с пробелом,
  // поле значений пересоздалось — и сортировка по убыванию сбросилась.
  const state = pivotBook();
  const plan = await prepareUpdatePivotPlan({ pivot: "СвВыручка", values: [{ field: "Выручка", label: "Выручка, тыс." }], sort: { field: "Категория", order: "desc" } });
  assert.equal(plan.values, undefined, "те же поля и агрегации — значения не пересоздаются");
  const result = await executeUpdatePivotPlan(plan) as any;
  assert.equal(state.removedValues, 0);
  assert.equal(state.values[0].name, "Выручка, тыс.");
  assert.deepEqual(state.sorted, { field: "Категория", order: "Descending", by: "Выручка, тыс." });
  assert.equal(result.executionState, "verified");
  await assert.rejects(() => prepareUpdatePivotPlan({ pivot: "СвВыручка", sort: { field: "Год", order: "desc" } }), /Сортировать можно поле строк/);
});

test("10.9: a value field shown as a share of the total, a field put into the pivot's filters", async () => {
  // Инструкция пользователя: «Дополнительные вычисления» — доли, нарастающий итог, отличие, ранг.
  const state = pivotBook();
  const plan = await prepareUpdatePivotPlan({
    pivot: "СвВыручка", filters: ["Регион"],
    values: [{ field: "Выручка", label: "Доля", showAs: { calculation: "percentOfGrandTotal" } }]
  });
  assert.ok(plan.preview.some((line) => /% от общего итога/.test(line)));
  assert.ok(plan.preview.some((line) => /Фильтрах.*Регион/.test(line)));
  const result = await executeUpdatePivotPlan(plan) as any;
  assert.equal(result.executionState, "verified");
  assert.equal(state.values[0].showAs.calculation, "PercentOfGrandTotal");
  assert.equal(state.values[0].numberFormat, "0,0%", "доля по умолчанию — в процентах");
  assert.deepEqual(state.filters, ["Регион"]);
});

test("10.9: what Excel needs for a calculation is asked before Excel", async () => {
  const { checkShowAs } = await import("./pivotFinish");
  assert.throws(() => checkShowAs({ calculation: "runningTotal" }), /baseField/);
  assert.throws(() => checkShowAs({ calculation: "differenceFrom", baseField: "Месяц" }), /baseItem.*Предыдущий/s);
  assert.throws(() => checkShowAs({ calculation: "previous" }), /calculation/);
  assert.deepEqual(checkShowAs({ calculation: "rankDescending", baseField: " Месяц " }), { calculation: "rankDescending", baseField: "Месяц" });
});

test("10.9: four value fields are not tuned — Excel 2021 crashes on it; plain fields are fine", async () => {
  pivotBook();
  const four = ["Выручка", "Выручка", "Выручка", "Выручка"].map((field, index) => ({ field, ...(index === 1 ? { label: "Доля", showAs: { calculation: "percentOfGrandTotal" } } : {}) }));
  await assert.rejects(() => prepareUpdatePivotPlan({ pivot: "СвВыручка", values: four }), /4 полей значений.*Excel падает.*Параметры полей значений.*не выполнялась/s);
  const plain = await prepareUpdatePivotPlan({ pivot: "СвВыручка", values: four.map(({ field }) => ({ field })) });
  assert.equal(plain.finish.values.length, 0);
  await assert.doesNotReject(() => prepareUpdatePivotPlan({ pivot: "СвВыручка", values: four.slice(0, 3) }));
});
