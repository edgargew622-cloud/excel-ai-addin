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
import { executeApplyFilterPlan, executeSortRangePlan, prepareApplyFilterPlan, prepareSortRangePlan } from "./excelTools";

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

test("headers are recognised when the key column is text", async () => {
  const { firstRowLooksLikeHeader } = await import("./sortFilter");
  // Сортировка по «Городу»: ключ текстовый, но числа есть в соседних столбцах.
  assert.equal(firstRowLooksLikeHeader([
    ["Дата", "Город", "Количество"],
    ["2026-09-04", "Омск", 4],
    ["2026-09-01", "Москва", 2]
  ]), true);
  // Сплошной текст без числовых столбцов — заголовки не доказаны.
  assert.equal(firstRowLooksLikeHeader([["Омск", "Кофе"], ["Москва", "Чай"]]), false);
  // Первая строка с числом — это данные.
  assert.equal(firstRowLooksLikeHeader([["Омск", 4], ["Москва", 2]]), false);
});

test("the preview warns about headers when sorting by a text column", async () => {
  salesExcel();
  const plan = await prepareSortRangePlan({ sheet: "Продажи", address: "A1:C4", column: 0 });
  assert.match(plan.headerWarning ?? "", /похожа на заголовки/);
});

test("Excel's empty placeholder criteria are not counted as filter conditions", async () => {
  const { describeCriteria } = await import("./sortFilter");
  // Так Excel отвечал по «Справочнику» без фильтра: заготовка в каждом столбце.
  const placeholders = describeCriteria([
    { filterOn: "BottomItems" },
    { filterOn: "BottomItems", values: [] },
    { filterOn: "BottomItems", criterion1: "" }
  ]);
  assert.equal(placeholders.activeColumns, 0);
  assert.deepEqual(placeholders.activeIndexes, []);

  // Настоящий «последние N элементов» несёт порог и считается.
  const real = describeCriteria([{ filterOn: "BottomItems", criterion1: "3" }, { filterOn: "Values", values: ["Кофе"] }]);
  assert.deepEqual(real.activeIndexes, [0, 1]);
});

test("a filter on the same area adds, on another area replaces", async () => {
  const { filterChangeKind } = await import("./sortFilter");
  const coffee = {
    enabled: true,
    address: "Справочник!$A$1:$C$4",
    activeColumns: 1,
    activeIndexes: [0],
    criteria: "[]"
  };
  // Проверка 6: та же область, другой столбец — условия сложились.
  assert.equal(filterChangeKind(coffee, "A1:C4", 1), "adds");
  // Тот же столбец — заменяется только его условие.
  assert.equal(filterChangeKind(coffee, "A1:C4", 0), "replacesColumn");
  // Другая область — прежний фильтр пропадает целиком.
  assert.equal(filterChangeKind(coffee, "E1:G10", 0), "replacesFilter");
  // Фильтра нет, или стоит без условий — ничего не теряется.
  assert.equal(filterChangeKind({ ...coffee, enabled: false }, "E1:G10", 0), "new");
  assert.equal(filterChangeKind({ ...coffee, activeColumns: 0, activeIndexes: [] }, "E1:G10", 0), "new");
});

/** Заготовка условия для столбца без фильтра: непустые значения по умолчанию,
 * которые ничего не отбирают. Так, по всей видимости, отвечает Excel. */
const placeholder = () => ({
  filterOn: "BottomItems",
  criterion1: "",
  criterion2: "",
  values: [],
  dynamicCriteria: "Unknown",
  color: "",
  icon: { set: "Invalid", index: 0 }
});

test("default placeholder values are not conditions", async () => {
  const { hasCondition } = await import("./sortFilter");
  assert.equal(hasCondition(placeholder()), false);
  assert.equal(hasCondition({ ...placeholder(), filterOn: "Values", values: ["Кофе"] }), true);
  assert.equal(hasCondition({ ...placeholder(), filterOn: "Custom", criterion1: ">0,1" }), true);
  assert.equal(hasCondition({ ...placeholder(), filterOn: "Dynamic", dynamicCriteria: "Today" }), true);
  assert.equal(hasCondition({ ...placeholder(), filterOn: "Icon", icon: { set: "ThreeArrows", index: 1 } }), true);
});

/** «Справочник» с автофильтром, который складывает условия на одной области
 * и держит заготовки в остальных столбцах — повторяет поведение из проверки 6. */
function referenceSheetWithFilter() {
  const values = [
    ["Товар", "Категория", "Наценка"],
    ["Кофе", "Напитки", 0.15],
    ["Чай", "Напитки", 0.1],
    ["Какао", "Напитки", 0.2]
  ];
  const criteria: any[] = [placeholder(), placeholder(), placeholder()];
  let enabled = false;
  const visible = () => {
    if (!enabled) return values.length;
    return 1 + values.slice(1).filter((row) => criteria.every((item, index) => {
      if (Array.isArray(item.values) && item.values.length) return item.values.includes(String(row[index]));
      if (item.criterion1) {
        const threshold = Number(String(item.criterion1).replace(/^[<>=]+/, "").replace(",", "."));
        return String(item.criterion1).startsWith(">") ? Number(row[index]) > threshold : true;
      }
      return true;
    })).length;
  };
  const range: any = {
    address: "Справочник!A1:C4", rowCount: 4, columnCount: 3, load: () => undefined, values,
    getVisibleView: () => ({ load: () => undefined, get rowCount() { return visible(); } })
  };
  const filterRange: any = {
    load: () => undefined,
    get isNullObject() { return !enabled; },
    address: "Справочник!$A$1:$C$4",
    getRow: () => ({ load: () => undefined, values: [values[0]] })
  };
  const sheet: any = {
    id: "sheet-r", name: "Справочник", load: () => undefined,
    getRange: () => range,
    tables: { items: [], load: () => undefined },
    autoFilter: {
      load: () => undefined,
      get enabled() { return enabled; },
      get criteria() { return criteria; },
      getRangeOrNullObject: () => filterRange,
      apply: (_range: unknown, column: number, condition: any) => {
        enabled = true;
        criteria[column] = { ...placeholder(), ...condition };
      }
    }
  };
  (globalThis as any).Excel = {
    FilterOn: { values: "Values", custom: "Custom" },
    run: async (fn: any) => fn({
      workbook: { worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet } },
      sync: async () => undefined
    })
  };
}

test("the three filters from check 6 are described as they really stack", async () => {
  referenceSheetWithFilter();
  const { executeApplyFilterPlan } = await import("./excelTools");

  const first = await prepareApplyFilterPlan({ sheet: "Справочник", address: "A1:C4", column: 0, criteria: "Кофе" });
  assert.equal(first.change, "new", "заготовки в столбцах не выдают себя за прежний фильтр");
  await executeApplyFilterPlan(first);

  const second = await prepareApplyFilterPlan({ sheet: "Справочник", address: "A1:C4", column: 1, criteria: "Напитки" });
  assert.equal(second.change, "adds");
  await executeApplyFilterPlan(second);

  const third = await prepareApplyFilterPlan({ sheet: "Справочник", address: "A1:C4", column: 2, criteria: ">0,1" });
  assert.equal(third.change, "adds", "у «Наценки» условия не было — оно добавляется");
  const result = await executeApplyFilterPlan(third) as any;

  // Отчёт опирается на действующие условия и понятные числа строк.
  assert.deepEqual(result.conditionsAfter.map((item: any) => item.header), ["Товар", "Категория", "Наценка"]);
  assert.equal(result.areaRows, 4);
  assert.equal(result.visibleRowsAfter, 2);
  assert.equal(result.hiddenRowsAfter, 2);
  assert.ok(Array.isArray(result.criteriaRaw));

  const again = await prepareApplyFilterPlan({ sheet: "Справочник", address: "A1:C4", column: 2, criteria: ">0,05" });
  assert.equal(again.change, "replacesColumn", "повтор в том же столбце заменяет только его условие");
});

test("a filter on a protected sheet is refused up front, unless protection allows filtering", async () => {
  // План стабилизации, S4: фильтр не проверял защиту листа вовсе. Excel
  // отказал бы уже во время операции, и итог вышел бы «неизвестен»
  // вместо честного «не выполнялось».
  const excel = salesExcel();
  excel.sheet.protection = { protected: true, options: { allowAutoFilter: false }, load: () => undefined };
  await assert.rejects(
    () => prepareApplyFilterPlan({ sheet: "Продажи", address: "A1:C4", column: 0, criteria: "Москва" }),
    /защищён/
  );

  // Защита, в которой фильтр разрешён, — не повод отказывать.
  excel.sheet.protection = { protected: true, options: { allowAutoFilter: true }, load: () => undefined };
  await prepareApplyFilterPlan({ sheet: "Продажи", address: "A1:C4", column: 0, criteria: "Москва" });
});

test("protection switched on after the preview stops the filter before it runs", async () => {
  const excel = salesExcel();
  const plan = await prepareApplyFilterPlan({ sheet: "Продажи", address: "A1:C4", column: 0, criteria: "Москва" });
  excel.sheet.protection = { protected: true, options: { allowAutoFilter: false }, load: () => undefined };
  await assert.rejects(() => executeApplyFilterPlan(plan), (error: any) => {
    assert.equal(error.executionState, "failed_before_write");
    assert.match(error.message, /защищён/);
    return true;
  });
});
