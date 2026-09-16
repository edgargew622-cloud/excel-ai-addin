import test from "node:test";
import assert from "node:assert/strict";
import {
  describeCriteria,
  excelSortCompare,
  isSortedLikeExcel,
  parseFilterCriteria,
  partialRowSortProblem,
  sameAutoFilterState,
  sameRowMultiset,
  sortRowsLikeExcel
} from "./sortFilter";
import { executeSortRangePlan, prepareApplyFilterPlan, prepareSortRangePlan } from "./excelTools";

// --- чистая логика -----------------------------------------------------------

test("Excel order: numbers, then text, then logicals, blanks always last", () => {
  const values = ["Омск", 10, "", true, 2, "казань"];
  const asc = [...values].sort((a, b) => excelSortCompare(a, b, true));
  assert.deepEqual(asc, [2, 10, "казань", "Омск", true, ""]);
  // По убыванию пустые всё равно в конце — так делает Excel.
  const desc = [...values].sort((a, b) => excelSortCompare(a, b, false));
  assert.equal(desc[desc.length - 1], "");
});

test("text is compared without regard to case", () => {
  assert.equal(excelSortCompare("москва", "Москва"), 0);
  assert.ok(excelSortCompare("Казань", "москва") < 0);
});

test("the expected order is stable, like Excel's", () => {
  const rows = [["Москва", 1], ["Казань", 2], ["Москва", 3]];
  assert.deepEqual(sortRowsLikeExcel(rows, 0), [["Казань", 2], ["Москва", 1], ["Москва", 3]]);
  assert.equal(isSortedLikeExcel(sortRowsLikeExcel(rows, 0), 0), true);
  assert.equal(isSortedLikeExcel(rows, 0), false);
});

test("scrambled rows are caught even when every cell survived", () => {
  const before = [["Москва", 900], ["Казань", 840]];
  // Те же ячейки, но столбцы переставлены по отдельности: строки чужие.
  const scrambled = [["Казань", 900], ["Москва", 840]];
  assert.equal(sameRowMultiset(before, scrambled), false);
  assert.equal(sameRowMultiset(before, [["Казань", 840], ["Москва", 900]]), true);
  assert.equal(sameRowMultiset([["a"], ["a"]], [["a"], ["b"]]), false, "повторы считаются");
});

test("sorting part of a data block is recognised as dangerous", () => {
  const block = { rowStart: 1, rowEnd: 6, columnStart: 1, columnEnd: 7 };
  // Только столбец B внутри блока A:G — строки перемешаются.
  assert.ok(partialRowSortProblem({ rowStart: 2, rowEnd: 6, columnStart: 2, columnEnd: 2 }, block));
  // Весь блок — безопасно.
  assert.equal(partialRowSortProblem(block, block), null);
  // Нет сведений о блоке — не выдумываем проблему.
  assert.equal(partialRowSortProblem(block, null), null);
});

test("filter conditions are parsed into values or comparisons", () => {
  assert.deepEqual(parseFilterCriteria("Москва"), { filterOn: "values", values: ["Москва"] });
  assert.deepEqual(parseFilterCriteria("Москва | Казань"), { filterOn: "values", values: ["Москва", "Казань"] });
  assert.deepEqual(parseFilterCriteria(">500"), { filterOn: "custom", criterion1: ">500" });
  assert.throws(() => parseFilterCriteria("  "), /пустым/);
  assert.throws(() => parseFilterCriteria("|"), /ни одного значения/);
});

test("filter state compares by meaning, ignoring Excel's service fields", () => {
  const a = describeCriteria([{ filterOn: "Values", values: ["Москва"], color: "", icon: null }, {}]);
  const b = describeCriteria([{ filterOn: "Values", values: ["Москва"] }, {}]);
  assert.equal(a.text, b.text);
  assert.equal(a.activeColumns, 1);
  const state = { enabled: true, address: "Лист!A1:G6", activeColumns: 1, criteria: a.text };
  assert.equal(sameAutoFilterState(state, { ...state, address: "лист!a1:g6" }), true);
  assert.equal(sameAutoFilterState(state, { ...state, criteria: "[]" }), false);
});

// --- планы на макетах ------------------------------------------------------------

/** Макет листа «Продажи» из эталонной книги: A1:C4 и сплошной блок вокруг. */
function salesExcel(options: { regionAddress?: string; sortWrites?: "ok" | "scramble" | "throw" } = {}) {
  const state = {
    values: [
      ["Город", "Сумма", "Товар"],
      ["Омск", 1800, "Кофе"],
      ["Казань", 840, "Чай"],
      ["Москва", 900, "Кофе"]
    ] as unknown[][]
  };
  const range: any = {
    address: "Продажи!A1:C4",
    rowCount: 4,
    columnCount: 3,
    rowIndex: 0,
    columnIndex: 0,
    load: () => undefined,
    get values() { return state.values; },
    get formulas() { return state.values; },
    format: { protection: { locked: false, load: () => undefined } },
    getSurroundingRegion: () => ({ address: options.regionAddress ?? "Продажи!A1:C4", load: () => undefined }),
    sort: {
      apply: (fields: any[], _match: boolean, headers: boolean) => {
        if (options.sortWrites === "throw") throw new Error("Excel отказал");
        const head = headers ? state.values.slice(0, 1) : [];
        const body = headers ? state.values.slice(1) : state.values;
        let sorted = sortRowsLikeExcel(body, fields[0].key, fields[0].ascending).map((row) => [...row]);
        if (options.sortWrites === "scramble") {
          // Переставлен только первый столбец: как при сортировке части блока.
          const firstColumn = sorted.map((row) => row[0]);
          sorted = body.map((row, index) => [firstColumn[index], ...row.slice(1)]);
        }
        state.values = [...head, ...sorted];
      }
    }
  };
  const sheet: any = {
    id: "sheet-1",
    name: "Продажи",
    load: () => undefined,
    protection: { protected: false, load: () => undefined },
    getRange: () => range,
    getRangeByIndexes: () => range,
    tables: { items: [], load: () => undefined },
    autoFilter: { enabled: false, criteria: [], load: () => undefined, getRangeOrNullObject: () => ({ isNullObject: true, load: () => undefined }) }
  };
  (globalThis as any).Excel = {
    SortOn: { value: "Value" },
    run: async (fn: any) => fn({
      workbook: {
        application: { calculationMode: "automatic", load: () => undefined },
        worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet }
      },
      sync: async () => undefined
    })
  };
  return { state, sheet };
}

test("a sort that would cut rows out of a data block is refused before it runs", async () => {
  const excel = salesExcel({ regionAddress: "Продажи!A1:G4" });
  const before = JSON.stringify(excel.state.values);
  await assert.rejects(
    () => prepareSortRangePlan({ sheet: "Продажи", address: "A1:C4", column: 1, hasHeaders: true }),
    /строки перемешаются/
  );
  assert.equal(JSON.stringify(excel.state.values), before);
});

test("the preview shows the first rows now and after, headers kept apart", async () => {
  salesExcel();
  const plan = await prepareSortRangePlan({ sheet: "Продажи", address: "A1:C4", column: 1, hasHeaders: true });
  assert.equal(plan.keyHeader, "Сумма");
  assert.deepEqual(plan.previewAfter.map((row) => row[1]), [840, 900, 1800]);
  assert.equal(plan.headerWarning, undefined);
});

test("header-looking first row without hasHeaders is warned about", async () => {
  salesExcel();
  const plan = await prepareSortRangePlan({ sheet: "Продажи", address: "A1:C4", column: 1 });
  assert.match(plan.headerWarning ?? "", /похожа на заголовки/);
});

test("a manual change after the preview stops the sort", async () => {
  const excel = salesExcel();
  const plan = await prepareSortRangePlan({ sheet: "Продажи", address: "A1:C4", column: 1, hasHeaders: true });
  excel.state.values = excel.state.values.map((row) => [...row]);
  excel.state.values[2][1] = 999;
  await assert.rejects(() => executeSortRangePlan(plan), (error: any) => {
    assert.equal(error.executionState, "failed_before_write");
    return true;
  });
});

test("a clean sort is verified and reports the rows as Excel left them", async () => {
  salesExcel();
  const plan = await prepareSortRangePlan({ sheet: "Продажи", address: "A1:C4", column: 1, hasHeaders: true });
  const result = await executeSortRangePlan(plan) as any;
  assert.equal(result.executionState, "verified");
  assert.equal(result.rowsPreserved, true);
  assert.deepEqual(result.firstRowsAfter.map((row: unknown[]) => row[0]), ["Казань", "Москва", "Омск"]);
});

test("scrambled rows after a sort are never reported as success", async () => {
  salesExcel({ sortWrites: "scramble" });
  const plan = await prepareSortRangePlan({ sheet: "Продажи", address: "A1:C4", column: 1, hasHeaders: true });
  await assert.rejects(() => executeSortRangePlan(plan), (error: any) => {
    assert.match(error.message, /набор строк после неё не совпадает/);
    assert.equal(error.executionState, "applied");
    return true;
  });
});

test("a sort that failed inside Excel is undecided", async () => {
  salesExcel({ sortWrites: "throw" });
  const plan = await prepareSortRangePlan({ sheet: "Продажи", address: "A1:C4", column: 1, hasHeaders: true });
  await assert.rejects(() => executeSortRangePlan(plan), (error: any) => {
    assert.equal(error.executionState, "unknown");
    return true;
  });
});

test("a filter over an Excel table is refused: the table has its own filter", async () => {
  const excel = salesExcel();
  excel.sheet.tables = {
    load: () => undefined,
    items: [{ name: "SalesTable", getRange: () => ({ address: "Продажи!A1:G6", load: () => undefined }) }]
  };
  await assert.rejects(
    () => prepareApplyFilterPlan({ sheet: "Продажи", address: "A1:C4", column: 0, criteria: "Москва" }),
    /таблицей Excel «SalesTable»/
  );
});
