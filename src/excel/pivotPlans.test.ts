import test from "node:test";
import assert from "node:assert/strict";
import { expectPivot, pivotHeaderProblems, pivotMismatches } from "./pivotModel";
import { executeCreatePivotPlan, prepareCreatePivotPlan } from "./pivotPlans";
import { PLANNED_TOOLS } from "./plans";
import { clear as clearUndo, setUndoMonitorReady, undoLast } from "./undo";

const ORDERS = [
  ["Город", "Статус", "Сумма", "Менеджер"],
  ["Москва", "Новая", 900, "Иванов"],
  ["Москва", "В работе", 1200, "Иванов"],
  ["Москва", "Закрыта", 450, "Сидоров"],
  ["Казань", "Новая", 700, "Петрова"],
  ["Казань", "В работе", 250, "Иванов"],
  ["Омск", "Новая", 1500, "Сидоров"]
];

/* --- расчёт ---------------------------------------------------------------- */

test("the panel sums the groups and the grand total the way Excel will", () => {
  const e = expectPivot(ORDERS, ["Город"], [{ field: "Сумма", aggregation: "sum" }]);
  assert.deepEqual(e.groups, [
    { label: "Москва", totals: [2550] },
    { label: "Казань", totals: [950] },
    { label: "Омск", totals: [1500] }
  ]);
  assert.deepEqual(e.grandTotals, [5000]);
  // Шапка, три города, общий итог; столбец подписей и столбец суммы.
  assert.equal(e.height, 5);
  assert.equal(e.width, 2);
  assert.deepEqual(e.warnings, []);
});

test("two row fields add a line for every pair under every group", () => {
  const e = expectPivot(ORDERS, ["Город", "Статус"], [{ field: "Сумма", aggregation: "sum" }, { field: "Сумма", aggregation: "count" }]);
  // 3 города + 6 пар город-статус + шапка + итог.
  assert.equal(e.height, 11);
  // Макет табличный: у каждого поля строк свой столбец, плюс два поля значений.
  assert.equal(e.width, 4);
  assert.deepEqual(e.grandTotals, [5000, 6]);
});

test("summing text is named as a mistake, counting it is fine", () => {
  const sum = expectPivot(ORDERS, ["Город"], [{ field: "Менеджер", aggregation: "sum" }]);
  assert.ok(sum.warnings.some((text) => /нет чисел/.test(text)));
  const count = expectPivot(ORDERS, ["Город"], [{ field: "Менеджер", aggregation: "count" }]);
  assert.deepEqual(count.warnings, []);
  assert.deepEqual(count.grandTotals, [6]);
});

test("headers a pivot cannot use are refused up front", () => {
  assert.deepEqual(pivotHeaderProblems(["Город", "Сумма"]), []);
  const problems = pivotHeaderProblems(["Город", "", "город"]);
  assert.ok(problems.some((text) => /без заголовка/.test(text)));
  assert.ok(problems.some((text) => /повторяется/.test(text)));
});

test("the built pivot is compared number by number, grand total found by position", () => {
  const e = expectPivot(ORDERS, ["Город"], [{ field: "Сумма", aggregation: "sum" }]);
  const good = [["Названия строк", "Сумма по полю Сумма"], ["Казань", 950], ["Москва", 2550], ["Омск", 1500], ["Общий итог", 5000]];
  assert.deepEqual(pivotMismatches(e, good), []);
  const counted = [["Названия строк", "Количество по полю Сумма"], ["Казань", 2], ["Москва", 3], ["Омск", 1], ["Общий итог", 6]];
  const problems = pivotMismatches(e, counted);
  assert.ok(problems.some((text) => /общий итог/.test(text)));
});

/* --- вложенные уровни (S3.1) --------------------------------------------------- */

// Источник и макеты сняты с настоящего Excel 24 сентября 2026 года:
// табличный макет, промежуточные итоги внизу группы. Заодно видно, как
// Excel сводит подписи: текст «1» и число 1 — один элемент, «москва»
// и «Москва» — один, «Москва » с пробелом — отдельный; пустое — «(пусто)».
const PROBE = [
  ["Город", "Статус", "Сумма", "Год", "Код"],
  ["Москва", "Новая", 900, 2024, "1"],
  ["Москва", "Закрыта", 450, 2025, 1],
  ["Москва", "Новая", 100, 2025, "1"],
  ["Казань", "Новая", 700, 2024, 1],
  ["Казань", "Закрыта", 800, 2024, "1"],
  ["Омск", "Закрыта", 1300, 2025, 2],
  ["Омск", "", 50, 2025, 2],
  ["", "Новая", 5, 2024, 2]
];
const PROBE_LAYOUT = [
  ["Город", "Статус", "Сумма по полю Сумма"],
  ["Казань", "Закрыта", 800],
  ["", "Новая", 700],
  ["Казань Итог", "", 1500],
  ["Москва", "Закрыта", 450],
  ["", "Новая", 1000],
  ["Москва Итог", "", 1450],
  ["Омск", "Закрыта", 1300],
  ["", "(пусто)", 50],
  ["Омск Итог", "", 1350],
  ["(пусто)", "Новая", 5],
  ["(пусто) Итог", "", 5],
  ["Общий итог", "", 4305]
];
const SUM = [{ field: "Сумма", aggregation: "sum" as const }];

test("the example from the plan: wrong nested totals are no longer accepted", () => {
  // План стабилизации, S3.1: такой макет прежняя сверка принимала.
  const source = [["City", "Status", "Amount"], ["A", "Open", 10], ["A", "Closed", 20]];
  const e = expectPivot(source, ["City", "Status"], [{ field: "Amount", aggregation: "sum" }]);
  const wrong = [["City", "Status", "Sum"], ["A", "Closed", 999], ["", "Open", 999], ["A Total", "", 30], ["Grand Total", "", 30]];
  assert.notDeepEqual(pivotMismatches(e, wrong), []);
  const right = [["City", "Status", "Sum"], ["A", "Closed", 20], ["", "Open", 10], ["A Total", "", 30], ["Grand Total", "", 30]];
  assert.deepEqual(pivotMismatches(e, right), []);
  // И в прежнем, компактном виде он тоже не проходит.
  assert.notDeepEqual(pivotMismatches(e, [["Rows", "Sum"], ["A", 30], ["Open", 999], ["Closed", 999], ["Grand Total", 30]]), []);
});

test("a layout taken from real Excel is accepted: blanks, merged text and numbers", () => {
  const e = expectPivot(PROBE, ["Город", "Статус"], SUM);
  assert.equal(e.height, PROBE_LAYOUT.length);
  assert.equal(e.width, 3);
  assert.deepEqual(pivotMismatches(e, PROBE_LAYOUT), []);

  const years = expectPivot(PROBE, ["Год", "Код"], SUM);
  const yearsLayout = [
    ["Год", "Код", "Сумма по полю Сумма"],
    [2024, 1, 2400], ["", 2, 5], ["2024 Итог", "", 2405],
    [2025, 1, 550], ["", 2, 1350], ["2025 Итог", "", 1900],
    ["Общий итог", "", 4305]
  ];
  assert.equal(years.height, yearsLayout.length);
  assert.deepEqual(pivotMismatches(years, yearsLayout), []);
});

test("case merges items, a trailing space does not", () => {
  const lower = PROBE.map((row, index) => (index === 3 ? ["москва", ...row.slice(1)] : row));
  assert.deepEqual(pivotMismatches(expectPivot(lower, ["Город", "Статус"], SUM), PROBE_LAYOUT), []);

  const spaced = PROBE.map((row, index) => (index === 3 ? ["Москва ", ...row.slice(1)] : row));
  const e = expectPivot(spaced, ["Город", "Статус"], SUM);
  const layout = [
    ...PROBE_LAYOUT.slice(0, 4),
    ["Москва", "Закрыта", 450], ["", "Новая", 900], ["Москва Итог", "", 1350],
    ["Москва ", "Новая", 100], ["Москва  Итог", "", 100],
    ...PROBE_LAYOUT.slice(7)
  ];
  assert.equal(e.height, layout.length);
  assert.deepEqual(pivotMismatches(e, layout), []);
});

test("the same label under different parents is checked under its own parent", () => {
  const e = expectPivot(PROBE, ["Город", "Статус"], SUM);
  // «Новая» у Казани и Москвы поменялись местами: каждая сумма где-то
  // в сводной есть, но не под своим городом.
  const swapped = PROBE_LAYOUT.map((row) => [...row]);
  swapped[2][2] = 1000;
  swapped[5][2] = 700;
  const problems = pivotMismatches(e, swapped);
  assert.ok(problems.some((text) => /Казань › Новая/.test(text)), problems.join("; "));
  assert.ok(problems.some((text) => /Москва › Новая/.test(text)), problems.join("; "));
});

test("a wrong subtotal, a wrong leaf and a subtotal of another group are all caught", () => {
  const e = expectPivot(PROBE, ["Город", "Статус"], SUM);
  const badSubtotal = PROBE_LAYOUT.map((row) => [...row]);
  badSubtotal[3][2] = 1499;
  assert.ok(pivotMismatches(e, badSubtotal).some((text) => /Казань: 1499 вместо 1500/.test(text)));

  const badLeaf = PROBE_LAYOUT.map((row) => [...row]);
  badLeaf[1][2] = 801;
  assert.ok(pivotMismatches(e, badLeaf).some((text) => /Казань › Закрыта: 801 вместо 800/.test(text)));

  // Строка итога, подпись которой не называет свою группу, — макет не тот,
  // что ожидался, и молча пропускать её нельзя.
  const foreignSubtotal = PROBE_LAYOUT.map((row) => [...row]);
  foreignSubtotal[3][0] = "Омск Итог";
  assert.ok(pivotMismatches(e, foreignSubtotal).some((text) => /итог.*Казань/i.test(text)));
});

test("nested averages are taken over the rows, and every value field is checked", () => {
  const e = expectPivot(PROBE, ["Город", "Статус"], [
    { field: "Сумма", aggregation: "average" },
    { field: "Сумма", aggregation: "count" }
  ]);
  const layout = [
    ["Город", "Статус", "Среднее по полю Сумма", "Количество по полю Сумма"],
    ["Казань", "Закрыта", 800, 1], ["", "Новая", 700, 1], ["Казань Итог", "", 750, 2],
    ["Москва", "Закрыта", 450, 1], ["", "Новая", 500, 2], ["Москва Итог", "", 1450 / 3, 3],
    ["Омск", "Закрыта", 1300, 1], ["", "(пусто)", 50, 1], ["Омск Итог", "", 675, 2],
    ["(пусто)", "Новая", 5, 1], ["(пусто) Итог", "", 5, 1],
    ["Общий итог", "", 4305 / 8, 8]
  ];
  assert.deepEqual(pivotMismatches(e, layout), []);
  const wrongCount = layout.map((row) => [...row]);
  wrongCount[6][3] = 2;
  assert.ok(pivotMismatches(e, wrongCount).some((text) => /Москва.*столбец 4: 2 вместо 3/.test(text)));
});

test("three levels: subtotals of the middle level are placed by their full path", () => {
  const source = [
    ["Город", "Статус", "Менеджер", "Сумма"],
    ["Москва", "Новая", "Иванов", 10],
    ["Москва", "Новая", "Петрова", 20],
    ["Москва", "Закрыта", "Иванов", 30],
    ["Казань", "Новая", "Иванов", 40]
  ];
  const e = expectPivot(source, ["Город", "Статус", "Менеджер"], SUM);
  const layout = [
    ["Город", "Статус", "Менеджер", "Сумма по полю Сумма"],
    ["Казань", "Новая", "Иванов", 40], ["", "Новая Итог", "", 40], ["Казань Итог", "", "", 40],
    ["Москва", "Закрыта", "Иванов", 30], ["", "Закрыта Итог", "", 30],
    ["", "Новая", "Иванов", 10], ["", "", "Петрова", 20], ["", "Новая Итог", "", 30],
    ["Москва Итог", "", "", 60],
    ["Общий итог", "", "", 100]
  ];
  assert.equal(e.height, layout.length);
  assert.deepEqual(pivotMismatches(e, layout), []);
  const wrong = layout.map((row) => [...row]);
  wrong[8][3] = 31;
  assert.ok(pivotMismatches(e, wrong).some((text) => /Москва › Новая: 31 вместо 30/.test(text)));
});

/* --- полный путь ------------------------------------------------------------ */

function ordersSheet(options: {
  builds?: "right" | "count";
  occupied?: string;
  /** Формулы с пустым результатом, например { G3: '=""' }: значение у них пустое. */
  emptyFormulas?: Record<string, string>;
  /** Формулы источника: текст формулы виден, значение берётся из данных. */
  sourceFormulas?: Record<string, string>;
  protectedSheet?: boolean;
  /** Excel не смог прочитать таблицы листа. */
  tablesThrow?: boolean;
  /** Объединение на листе, например "F3:G3". */
  merged?: string;
  /** Excel отказывает на sync, в котором добавляются поля сводной. */
  failFields?: boolean;
  /** Excel отказывает в удалении сводной. */
  failDelete?: boolean;
  /** Что Excel построит — раскладка целиком (8.2); без неё — сумма по городам. */
  buildLayout?: unknown[][];
  /** Поддерживает ли Excel наборы API (фильтры сводной — 1.12). */
  apiSupported?: boolean;
  /** Другие данные листа вместо ORDERS (8.2, даты). */
  grid?: unknown[][];
  /** Числовой формат всех ячеек листа. */
  numberFormat?: string;
  /** Что Excel вычислит по записанной формуле. */
  evaluate?: (formula: string) => unknown;
} = {}) {
  const grid: unknown[][] = (options.grid ?? ORDERS).map((row) => [...row]);
  /** Записанное исполнителем: ячейка → формула или текст (8.2, вспомогательные столбцы). */
  const written = new Map<string, string>();
  const pivots: any[] = [];
  const extra = new Map<string, unknown>();
  if (options.occupied) extra.set(options.occupied, "занято");
  const emptyFormulas = new Map<string, string>(Object.entries(options.emptyFormulas ?? {}));
  const sourceFormulas = new Map<string, string>(Object.entries(options.sourceFormulas ?? {}));
  const tables: { name: string; address: string }[] = [];
  const protection = { protected: options.protectedSheet === true, load: () => undefined };
  const layouts: string[] = [];
  /** Что исполнитель попросил у Excel сверх полей: фильтры и порядок (8.2). */
  const calls: any[] = [];

  const cellValue = (row: number, column: number) => {
    const key = `${String.fromCharCode(65 + column)}${row + 1}`;
    if (written.has(key)) {
      const text = written.get(key)!;
      return text.startsWith("=") ? (options.evaluate ?? (() => ""))(text) : text;
    }
    if (emptyFormulas.has(key)) return "";
    if (extra.has(key)) return extra.get(key);
    for (const pivot of pivots) {
      const r = row - pivot.row;
      const c = column - pivot.column;
      if (r >= 0 && c >= 0 && r < pivot.values.length && c < pivot.values[0].length) return pivot.values[r][c];
    }
    return grid[row]?.[column] ?? "";
  };
  const cellFormula = (row: number, column: number) => {
    const key = `${String.fromCharCode(65 + column)}${row + 1}`;
    return written.get(key) ?? emptyFormulas.get(key) ?? sourceFormulas.get(key) ?? cellValue(row, column);
  };
  const parse = (address: string) => {
    const [a, b = a] = address.replace(/.*!/, "").split(":");
    const cell = (text: string) => ({ column: text.charCodeAt(0) - 65, row: Number(text.slice(1)) - 1 });
    return { start: cell(a), end: cell(b) };
  };
  const makeRange = (address: string): any => {
    const { start, end } = parse(address);
    const read = () => Array.from({ length: end.row - start.row + 1 }, (_, r) =>
      Array.from({ length: end.column - start.column + 1 }, (_, c) => cellValue(start.row + r, start.column + c)));
    return {
      address: `Заказы!${address}`,
      rowIndex: start.row,
      columnIndex: start.column,
      rowCount: end.row - start.row + 1,
      columnCount: end.column - start.column + 1,
      load: () => undefined,
      get values() { return read(); },
      get formulas() {
        return Array.from({ length: end.row - start.row + 1 }, (_, r) =>
          Array.from({ length: end.column - start.column + 1 }, (_, c) => cellFormula(start.row + r, start.column + c)));
      },
      set formulas(rows: string[][]) {
        rows.forEach((row, r) => row.forEach((text, c) => written.set(`${String.fromCharCode(65 + start.column + c)}${start.row + r + 1}`, text)));
      },
      get numberFormat() {
        return Array.from({ length: end.row - start.row + 1 }, () => Array.from({ length: end.column - start.column + 1 }, () => options.numberFormat ?? "General"));
      },
      clear: () => {
        for (let r = start.row; r <= end.row; r++) for (let c = start.column; c <= end.column; c++) written.delete(`${String.fromCharCode(65 + c)}${r + 1}`);
      },
      getMergedAreasOrNullObject: () => ({
        isNullObject: !options.merged,
        address: options.merged ? `Заказы!${options.merged}` : "",
        areaCount: options.merged ? 1 : 0,
        areas: { items: options.merged ? [{ address: `Заказы!${options.merged}` }] : [], load: () => undefined },
        load: () => undefined
      })
    };
  };
  const letter = (index: number) => String.fromCharCode(65 + index);

  const sheet: any = {
    id: "sheet-1",
    name: "Заказы",
    isNullObject: false,
    load: () => undefined,
    getRange: (address: string) => makeRange(address),
    getRangeByIndexes: (row: number, column: number, rows: number, columns: number) =>
      makeRange(`${letter(column)}${row + 1}:${letter(column + columns - 1)}${row + rows}`),
    getUsedRangeOrNullObject: () => ({ isNullObject: false, rowIndex: 0, columnIndex: 0, rowCount: grid.length, columnCount: grid[0].length, load: () => undefined }),
    protection,
    tables: {
      get items() {
        return tables.map((table) => ({ name: table.name, getRange: () => ({ address: table.address, load: () => undefined }) }));
      },
      getItem: (name: string) => ({ kind: "table", name }),
      load: () => { if (options.tablesThrow) throw new Error("Во время обработки запроса произошла внутренняя ошибка."); }
    },
    pivotTables: {
      get items() { return pivots.map((pivot) => ({ name: pivot.name, layout: { getRange: () => makeRange(pivot.address) } })); },
      load: () => undefined,
      add: (name: string, source: unknown, destination: any) => {
        const fields: string[] = [];
        const data: any[] = [];
        const columnFields: string[] = [];
        const pageFields: string[] = [];
        const hierarchyOn = (axis: string) => (hierarchy: string) => ({
          name: hierarchy,
          fields: {
            getItem: (field: string) => ({
              applyFilter: (filter: unknown) => calls.push({ axis, hierarchy, field, filter }),
              sortByValues: (order: string, by: any) => calls.push({ axis, field, sort: order, by: by.name })
            })
          }
        });
        const pivot: any = {
          name,
          source,
          row: destination.rowIndex,
          column: destination.columnIndex,
          values: [] as unknown[][],
          address: "",
          hierarchies: { getItem: (field: string) => field },
          rowHierarchies: {
            add: (field: string) => { pivot.pendingFields = true; return fields.push(field); },
            getItem: hierarchyOn("rows"),
            load: () => undefined,
            get items() { return fields.map((field) => hierarchyOn("rows")(field)); }
          },
          filterHierarchies: {
            add: (field: string) => { pageFields.push(field); calls.push({ page: field }); },
            load: () => undefined,
            get items() { return pageFields.map((field) => ({ name: field })); }
          },
          columnHierarchies: {
            add: (field: string) => { pivot.pendingFields = true; columnFields.push(field); calls.push({ column: field }); },
            getItem: hierarchyOn("columns"),
            load: () => undefined,
            get items() { return columnFields.map((field) => ({ name: field })); }
          },
          dataHierarchies: {
            add: (field: string) => {
              pivot.pendingFields = true;
              // Имя поля значений Excel даёт на языке интерфейса.
              const item: any = { field, summarizeBy: "Sum", name: `Сумма по полю ${field}`, numberFormat: "Общий" };
              data.push(item);
              return item;
            },
            get items() { return data; },
            load: () => undefined
          },
          layout: {
            load: () => undefined,
            autoFormat: true,
            showRowGrandTotals: true,
            showColumnGrandTotals: true,
            getRange: () => makeRange(pivot.address),
            set layoutType(value: string) { layouts.push(value); },
            set subtotalLocation(value: string) { layouts.push(`итоги: ${value}`); }
          }
        };
        // Так Excel сводит: по городу, сумма или — в режиме порчи — количество.
        pivot.build = () => {
          if (options.buildLayout) {
            pivot.values = options.buildLayout.map((row) => [...row]);
            const letter = (index: number) => String.fromCharCode(65 + index);
            pivot.address = `${letter(pivot.column)}${pivot.row + 1}:${letter(pivot.column + pivot.values[0].length - 1)}${pivot.row + pivot.values.length}`;
            return;
          }
          const counting = options.builds === "count";
          const cities = ["Казань", "Москва", "Омск"];
          const total = (city?: string) => {
            const rows = grid.slice(1).filter((row) => !city || row[0] === city);
            return counting ? rows.length : rows.reduce((sum, row) => sum + (row[2] as number), 0);
          };
          pivot.values = [["Названия строк", "Сумма по полю Сумма"], ...cities.map((city) => [city, total(city)]), ["Общий итог", total()]];
          const letter = (index: number) => String.fromCharCode(65 + index);
          pivot.address = `${letter(pivot.column)}${pivot.row + 1}:${letter(pivot.column + 1)}${pivot.row + pivot.values.length}`;
        };
        pivots.push(pivot);
        return pivot;
      }
    }
  };
  (globalThis as any).Office = { context: { document: { url: "C:/pivot.xlsx" }, requirements: { isSetSupported: () => options.apiSupported !== false } } };
  (globalThis as any).Excel = {
    run: async (fn: any) => fn({
      application: { cultureInfo: { numberFormat: { numberDecimalSeparator: ",", numberGroupSeparator: " ", load: () => undefined } } },
      workbook: {
        worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet, getItemOrNullObject: () => sheet },
        pivotTables: {
          getItemOrNullObject: (name: string) => {
            const pivot = pivots.find((item) => item.name === name);
            return {
              isNullObject: !pivot,
              load: () => undefined,
              delete: () => {
                if (options.failDelete) throw new Error("Удаление сводной не удалось.");
                pivots.splice(pivots.indexOf(pivot), 1);
              }
            };
          }
        }
      },
      sync: async () => {
        if (options.failFields && pivots.some((pivot) => pivot.pendingFields)) {
          for (const pivot of pivots) pivot.pendingFields = false;
          throw new Error("Поле сводной не найдено.");
        }
        for (const pivot of pivots) if (!pivot.values.length) pivot.build();
      }
    })
  };
  return { grid, pivots, extra, emptyFormulas, tables, protection, layouts, calls, written };
}

test("create_pivot_table goes through the plan registry", () => {
  assert.ok(PLANNED_TOOLS.includes("create_pivot_table"));
});

test("a pivot lands past the data and its totals are checked against the panel's", async () => {
  ordersSheet();
  const plan = await prepareCreatePivotPlan({ sheet: "Заказы", sourceAddress: "A1:D7", rows: ["Город"], values: [{ field: "Сумма" }] });
  assert.equal(plan.destCell, "F1");
  assert.equal(plan.destArea, "F1:G5");
  assert.ok(plan.preview.includes("Общий итог: 5000"));

  const result = await executeCreatePivotPlan(plan) as any;
  assert.equal(result.executionState, "verified");
  assert.deepEqual(result.grandTotals, [5000]);
});

test("a pivot that counts instead of summing is caught and can be undone", async () => {
  const state = ordersSheet({ builds: "count" });
  setUndoMonitorReady(true);
  try {
    const plan = await prepareCreatePivotPlan({ sheet: "Заказы", sourceAddress: "A1:D7", rows: ["Город"], values: [{ field: "Сумма" }] });
    await assert.rejects(() => executeCreatePivotPlan(plan), (error: any) => {
      assert.equal(error.executionState, "applied");
      assert.match(error.message, /общий итог.*6 вместо 5000/);
      return true;
    });
    await undoLast();
    assert.equal(state.pivots.length, 0);
  } finally {
    clearUndo();
    setUndoMonitorReady(false);
  }
});

test("an occupied destination is refused instead of overwritten, and a free one is named", async () => {
  ordersSheet({ occupied: "G3" });
  // Проверка 20 сентября 2026 года: отказ не называл свободного места,
  // и агент на этом бросал задачу, хотя рядом было пусто.
  await assert.rejects(
    () => prepareCreatePivotPlan({ sheet: "Заказы", sourceAddress: "A1:D7", rows: ["Город"], values: [{ field: "Сумма" }] }),
    (error: any) => {
      assert.match(error.message, /1 непустых ячеек — они были бы затёрты/);
      assert.match(error.message, /Свободно, например, Заказы!A10 — повторите с destAddress: "A10"/);
      return true;
    }
  );
});

test("a field name that is not a header is named with the real headers", async () => {
  ordersSheet();
  await assert.rejects(
    () => prepareCreatePivotPlan({ sheet: "Заказы", sourceAddress: "A1:D7", rows: ["Регион"], values: [{ field: "Сумма" }] }),
    /Нет полей «Регион».*«Город»/
  );
});

test("the same field may be both a row and a counted value, as in Excel", async () => {
  // Проверка 20 сентября 2026 года: панель это запрещала, хотя «Количество
  // по полю Статус» при строках по статусу — обычная сводная Excel.
  ordersSheet();
  const plan = await prepareCreatePivotPlan({
    sheet: "Заказы",
    sourceAddress: "A1:D7",
    rows: ["Город", "Статус"],
    values: [{ field: "Статус", aggregation: "count" }]
  });
  assert.deepEqual(plan.expectation.grandTotals, [6]);
});

test("a pivot on top of its own source is refused", async () => {
  ordersSheet();
  await assert.rejects(
    () => prepareCreatePivotPlan({ sheet: "Заказы", sourceAddress: "A1:D7", destAddress: "C3", rows: ["Город"], values: [{ field: "Сумма" }] }),
    /наложится на источник/
  );
});

/* --- S2: место и исходные данные сводной ------------------------------------ */

const SUMS = { sheet: "Заказы", sourceAddress: "A1:D7", rows: ["Город"], values: [{ field: "Сумма" }] };

test("a formula with an empty result is an occupied cell, not a free one", async () => {
  // План стабилизации, S2: формула ="" даёт пустое значение и проходила
  // проверку места, хотя сводная её затёрла бы.
  ordersSheet({ emptyFormulas: { G3: '=""' } });
  await assert.rejects(() => prepareCreatePivotPlan(SUMS), /1 непустых ячеек — они были бы затёрты/);
});

test("the free place offered in a refusal skips empty-looking formulas too", async () => {
  // Правее данных занято значением, под данными — формулой с пустым результатом.
  ordersSheet({ occupied: "G3", emptyFormulas: { A10: '=""' } });
  await assert.rejects(() => prepareCreatePivotPlan(SUMS), (error: any) => {
    assert.doesNotMatch(error.message, /A10/, "формула =\"\" не свободное место");
    assert.match(error.message, /не нашлось/);
    return true;
  });
});

test("an empty-looking formula written after the preview stops the pivot", async () => {
  const state = ordersSheet();
  const plan = await prepareCreatePivotPlan(SUMS);
  state.emptyFormulas.set("F2", '=""');
  await assert.rejects(() => executeCreatePivotPlan(plan), (error: any) => {
    assert.equal(error.executionState, "failed_before_write");
    assert.match(error.message, /перестало быть пустым/);
    return true;
  });
  assert.equal(state.pivots.length, 0, "сводная не создана");
  assert.equal(state.emptyFormulas.get("F2"), '=""', "формула пользователя цела");
});

test("a source whose values changed under the same formulas stops the pivot", async () => {
  // Формула источника ссылается на другой лист: её текст прежний, а значение
  // изменилось — расчёт групп и итогов панели устарел.
  const state = ordersSheet({ sourceFormulas: { C2: "=Внешний!A1" } });
  const plan = await prepareCreatePivotPlan(SUMS);
  state.grid[1][2] = 5000;
  await assert.rejects(() => executeCreatePivotPlan(plan), (error: any) => {
    assert.equal(error.executionState, "failed_before_write");
    assert.match(error.message, /изменились после предпросмотра/);
    return true;
  });
  assert.equal(state.pivots.length, 0);
});

test("a protected destination sheet is refused before anything is attempted", async () => {
  ordersSheet({ protectedSheet: true });
  await assert.rejects(() => prepareCreatePivotPlan(SUMS), /защищён/);
});

test("protection switched on after the preview stops the pivot", async () => {
  const state = ordersSheet();
  const plan = await prepareCreatePivotPlan(SUMS);
  state.protection.protected = true;
  await assert.rejects(() => executeCreatePivotPlan(plan), (error: any) => {
    assert.equal(error.executionState, "failed_before_write");
    assert.match(error.message, /защищён/);
    return true;
  });
});

test("a table that appeared in the place after the preview stops the pivot", async () => {
  const state = ordersSheet();
  const plan = await prepareCreatePivotPlan(SUMS);
  state.tables.push({ name: "Новая", address: "Заказы!F1:G3" });
  await assert.rejects(() => executeCreatePivotPlan(plan), (error: any) => {
    assert.equal(error.executionState, "failed_before_write");
    assert.match(error.message, /таблицу Новая/);
    return true;
  });
});

test("tables that could not be read are not taken for no tables", async () => {
  // readTableRanges глотает ошибку и отдаёт пустой список; для места сводной
  // «не смогли прочитать» — это не «таблиц нет».
  ordersSheet({ tablesThrow: true });
  await assert.rejects(() => prepareCreatePivotPlan(SUMS), /не удалось прочитать таблицы/i);
});

test("a merged area in the place is refused", async () => {
  ordersSheet({ merged: "F3:G3" });
  await assert.rejects(() => prepareCreatePivotPlan(SUMS), /объединен/i);
});

test("the pivot layout is set explicitly, so its size is the one computed", async () => {
  // Макет по умолчанию задаётся в настройках Excel. Размер и сверку панель
  // считает для табличного макета с итогами внизу групп (S3.1) — значит,
  // его и нужно выставить.
  const state = ordersSheet();
  await executeCreatePivotPlan(await prepareCreatePivotPlan(SUMS));
  assert.deepEqual(state.layouts, ["Tabular", "итоги: AtBottom"]);
});

/* --- сводная на новом листе (этап 7, 7.3.4) ------------------------------------------ */

test("a pivot on a new sheet names the sheet up front and refuses a taken name", async () => {
  ordersSheet();
  const run = (globalThis as any).Excel.run;
  // Список листов книги: «Заказы» — источник.
  (globalThis as any).Excel.run = async (fn: any) => run(async (ctx: any) => {
    ctx.workbook.worksheets.load = () => undefined;
    Object.defineProperty(ctx.workbook.worksheets, "items", { get: () => [{ name: "Заказы" }], configurable: true });
    return fn(ctx);
  });
  const plan = await prepareCreatePivotPlan({ sheet: "Заказы", sourceAddress: "A1:D7", rows: ["Город"], values: [{ field: "Сумма" }], newSheet: "Итоги по городам" });
  assert.equal(plan.newSheet, true);
  assert.equal(plan.destSheet, "Итоги по городам");
  assert.equal(plan.destArea, "A1:B5");
  await assert.rejects(
    () => prepareCreatePivotPlan({ sheet: "Заказы", sourceAddress: "A1:D7", rows: ["Город"], values: [{ field: "Сумма" }], newSheet: "заказы" }),
    /уже есть.*Свободно, например/
  );
  await assert.rejects(
    () => prepareCreatePivotPlan({ sheet: "Заказы", sourceAddress: "A1:D7", rows: ["Город"], values: [{ field: "Сумма" }], newSheet: "Х", destAddress: "F1" }),
    /не сочетается/
  );
});

test("the task sheet is not put into destSheet when the pivot goes to a new sheet", async () => {
  // Проверка в Excel 24 сентября 2026 года: подстановка листа назначения
  // по умолчанию превращала newSheet в отказ «не сочетается с destSheet».
  const { resolveToolArgs } = await import("./excelTools");
  const withNew: any = await resolveToolArgs("create_pivot_table", { sheet: "Заказы", sourceAddress: "A1:D7", rows: ["Город"], values: ["Сумма"], newSheet: "Итоги" });
  assert.equal(withNew.destSheet, undefined);
  const plain: any = await resolveToolArgs("create_pivot_table", { sheet: "Заказы", sourceAddress: "A1:D7", rows: ["Город"], values: ["Сумма"] });
  assert.equal(plain.destSheet, "Заказы");
});

/* --- сбой при добавлении полей ------------------------------------------------------- */

test("a pivot whose fields fail is removed before the error is returned", async () => {
  const state = ordersSheet({ failFields: true });
  const plan = await prepareCreatePivotPlan(SUMS);
  await assert.rejects(() => executeCreatePivotPlan(plan), (error: any) => {
    assert.equal(error.executionState, "failed_before_write");
    assert.match(error.message, /Поле сводной не найдено.*удалена.*в прежнем виде/);
    return true;
  });
  assert.equal(state.pivots.length, 0);
});

test("a pivot that could not be removed after a field failure stays unknown", async () => {
  const state = ordersSheet({ failFields: true, failDelete: true });
  const plan = await prepareCreatePivotPlan(SUMS);
  await assert.rejects(() => executeCreatePivotPlan(plan), (error: any) => {
    assert.equal(error.executionState, "unknown");
    assert.match(error.message, /убрать не удалось/);
    return true;
  });
  assert.equal(state.pivots.length, 1);
});

test("a sheet created for a pivot whose fields fail is removed with it", async () => {
  const state = ordersSheet({ failFields: true });
  const run = (globalThis as any).Excel.run;
  let created: any = null;
  (globalThis as any).Excel.run = async (fn: any) => run(async (ctx: any) => {
    const source = ctx.workbook.worksheets.getItem();
    ctx.workbook.worksheets.load = () => undefined;
    Object.defineProperty(ctx.workbook.worksheets, "items", {
      get: () => [{ name: "Заказы" }, ...(created && !created.deleted ? [{ name: created.name }] : [])],
      configurable: true
    });
    ctx.workbook.worksheets.add = (name: string) => {
      created = {
        id: "sheet-new", name, deleted: false,
        get isNullObject() { return this.deleted; },
        load: () => undefined,
        getRange: source.getRange,
        pivotTables: source.pivotTables,
        getUsedRangeOrNullObject: () => ({ isNullObject: state.pivots.length === 0, load: () => undefined }),
        delete() { this.deleted = true; }
      };
      return created;
    };
    ctx.workbook.worksheets.getItemOrNullObject = (id: string) => (id === "sheet-new" ? created : source);
    return fn(ctx);
  });

  const plan = await prepareCreatePivotPlan({ ...SUMS, newSheet: "Итоги" });
  await assert.rejects(() => executeCreatePivotPlan(plan), (error: any) => {
    assert.equal(error.executionState, "failed_before_write");
    assert.match(error.message, /вместе с созданным под неё листом «Итоги»/);
    return true;
  });
  assert.equal(state.pivots.length, 0);
  assert.equal(created?.deleted, true);
});

/* --- поле в столбцах, фильтры, порядок (этап 8, 8.2) ------------------------------ */

// Как Excel раскладывает сводную по ORDERS: статусы по алфавиту в столбцах,
// пустая ячейка там, где сочетания нет (замер 27.09.2026 на похожих данных).
const BY_STATUS = [
  ["Сумма по полю Сумма", "Статус", "", "", ""],
  ["Город", "В работе", "Закрыта", "Новая", "Общий итог"],
  ["Казань", 250, "", 700, 950],
  ["Москва", 1200, 450, 900, 2550],
  ["Омск", "", "", 1500, 1500],
  ["Общий итог", 1450, 450, 3100, 5000]
];

test("a column field is added to the pivot and every cell is checked, empty ones too", async () => {
  const state = ordersSheet({ buildLayout: BY_STATUS });
  const plan = await prepareCreatePivotPlan({ sheet: "Заказы", sourceAddress: "A1:D7", rows: ["Город"], columns: ["статус"], values: [{ field: "Сумма" }] });
  assert.equal(plan.columnField, "Статус", "имя поля — как в шапке источника");
  assert.equal(plan.destArea, "F1:J6");
  assert.ok(plan.preview.some((line) => /Столбцы по «Статус»: .*Новая/.test(line)));
  const result = await executeCreatePivotPlan(plan) as any;
  assert.equal(result.executionState, "verified");
  assert.deepEqual(state.calls, [{ column: "Статус" }]);

  const zero = BY_STATUS.map((row) => [...row]);
  zero[4][1] = 0;
  ordersSheet({ buildLayout: zero });
  const again = await prepareCreatePivotPlan({ sheet: "Заказы", sourceAddress: "A1:D7", rows: ["Город"], columns: ["Статус"], values: [{ field: "Сумма" }] });
  await assert.rejects(() => executeCreatePivotPlan(again), /Омск \(В работе\): 0 вместо пустой ячейки/);
});

test("top 2 cities: the filter goes to Excel by the value field's own name, the total shrinks", async () => {
  const state = ordersSheet({ buildLayout: [["Город", "Сумма по полю Сумма"], ["Москва", 2550], ["Омск", 1500], ["Общий итог", 4050]] });
  const plan = await prepareCreatePivotPlan({
    sheet: "Заказы", sourceAddress: "A1:D7", rows: ["Город"], values: [{ field: "Сумма" }], filters: [{ field: "Город", top: 2 }]
  });
  assert.equal(plan.destArea, "F1:G4");
  assert.ok(plan.preview.some((line) => /Фильтр «Город»: первые 2 по «Сумма»; уберутся «Казань»/.test(line)));
  assert.ok(plan.preview.includes("Общий итог: 4050"));
  const result = await executeCreatePivotPlan(plan) as any;
  assert.equal(result.executionState, "verified");
  assert.deepEqual(result.filteredOut, [{ field: "Город", items: ["Казань"] }]);
  assert.deepEqual(state.calls, [{
    axis: "rows", hierarchy: "Город", field: "Город",
    filter: { valueFilter: { condition: "TopN", threshold: 2, value: "Сумма по полю Сумма", selectionType: "Items" } }
  }]);
});

test("the place is checked for the pivot before its filter shrinks it", async () => {
  // До фильтра сводная по трём городам — F1:G5; G5 занята.
  ordersSheet({ occupied: "G5" });
  await assert.rejects(
    () => prepareCreatePivotPlan({ sheet: "Заказы", sourceAddress: "A1:D7", rows: ["Город"], values: [{ field: "Сумма" }], filters: [{ field: "Город", top: 2 }] }),
    /затёрты/
  );
});

test("chosen items go to Excel with their labels from the data; unknown ones are refused", async () => {
  const state = ordersSheet({ buildLayout: [["Город", "Сумма по полю Сумма"], ["Москва", 2550], ["Омск", 1500], ["Общий итог", 4050]] });
  const plan = await prepareCreatePivotPlan({
    sheet: "Заказы", sourceAddress: "A1:D7", rows: ["Город"], values: [{ field: "Сумма" }], filters: [{ field: "Город", include: ["москва", "Омск"] }]
  });
  assert.equal((await executeCreatePivotPlan(plan) as any).executionState, "verified");
  assert.deepEqual(state.calls[0].filter, { manualFilter: { selectedItems: ["Москва", "Омск"] } });

  ordersSheet();
  await assert.rejects(
    () => prepareCreatePivotPlan({ sheet: "Заказы", sourceAddress: "A1:D7", rows: ["Город"], values: [{ field: "Сумма" }], filters: [{ field: "Город", include: ["Пермь"] }] }),
    /«Пермь».*в данных нет.*«Москва»/
  );
});

test("filters the panel cannot check are refused before the card", async () => {
  ordersSheet();
  const base = { sheet: "Заказы", sourceAddress: "A1:D7", values: [{ field: "Сумма" }] };
  await assert.rejects(() => prepareCreatePivotPlan({ ...base, rows: ["Город", "Статус"], filters: [{ field: "Статус", top: 1 }] }), /только для первого поля rows/);
  await assert.rejects(() => prepareCreatePivotPlan({ ...base, rows: ["Город"], filters: [{ field: "Менеджер", top: 1 }] }), /ни в rows, ни в columns/);
  await assert.rejects(() => prepareCreatePivotPlan({ ...base, rows: ["Город"], filters: [{ field: "Город", top: 1, include: ["Омск"] }] }), /ровно одно/);
  await assert.rejects(() => prepareCreatePivotPlan({ ...base, rows: ["Город"], filters: [{ field: "Город", top: 1, by: "Менеджер" }] }), /поля «Менеджер» нет в values/);
  await assert.rejects(() => prepareCreatePivotPlan({ ...base, rows: ["Город"], columns: ["Город"] }), /и строками, и столбцами/);
  ordersSheet({ apiSupported: false });
  await assert.rejects(() => prepareCreatePivotPlan({ ...base, rows: ["Город"], filters: [{ field: "Город", top: 1 }] }), /ExcelApi 1\.12/);
});

test("sorting by value is asked of Excel and the order is checked", async () => {
  const sorted = [["Город", "Сумма по полю Сумма"], ["Москва", 2550], ["Омск", 1500], ["Казань", 950], ["Общий итог", 5000]];
  const state = ordersSheet({ buildLayout: sorted });
  const plan = await prepareCreatePivotPlan({
    sheet: "Заказы", sourceAddress: "A1:D7", rows: ["Город"], values: [{ field: "Сумма" }], sort: { field: "Город", order: "desc" }
  });
  assert.equal((await executeCreatePivotPlan(plan) as any).executionState, "verified");
  assert.deepEqual(state.calls, [{ axis: "rows", field: "Город", sort: "Descending", by: "Сумма по полю Сумма" }]);

  ordersSheet({ buildLayout: [["Город", "Сумма по полю Сумма"], ["Казань", 950], ["Москва", 2550], ["Омск", 1500], ["Общий итог", 5000]] });
  const again = await prepareCreatePivotPlan({
    sheet: "Заказы", sourceAddress: "A1:D7", rows: ["Город"], values: [{ field: "Сумма" }], sort: { field: "Город", order: "desc" }
  });
  await assert.rejects(() => executeCreatePivotPlan(again), (error: any) => {
    assert.equal(error.executionState, "applied");
    assert.match(error.message, /порядок по убыванию нарушен/);
    return true;
  });
});

/* --- группировка дат (этап 8, 8.2) ---------------------------------------------- */

// Даты — серийные номера Excel, как в замере 27.09.2026.
const DATES = [
  ["Дата", "Город", "Сумма"],
  [46037, "Москва", 100],
  [46056, "Москва", 200],
  [46073, "Казань", 50],
  [46122, "Казань", 70],
  [46204, "Омск", 30]
];
// Excel вычисляет формулы вспомогательных столбцов; здесь — по дате из ячейки.
const excelDates = (formula: string) => {
  const cell = /YEAR\(A(\d+)\)/.exec(formula);
  if (!cell) return "";
  const serial = DATES[Number(cell[1]) - 1][0] as number;
  const date = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
  const year = date.getUTCFullYear();
  const month = date.getUTCMonth() + 1;
  if (formula.includes("ROUNDUP")) return `${year} К${Math.ceil(month / 3)}`;
  if (formula.includes("TEXT(MONTH")) return `${year}-${String(month).padStart(2, "0")}`;
  return year;
};
const QUARTERS_MONTHS = [
  ["Дата (квартал)", "Дата (месяц)", "Сумма по полю Сумма"],
  ["2026 К1", "2026-01", 100], ["", "2026-02", 250], ["2026 К1 Итог", "", 350],
  ["2026 К2", "2026-04", 70], ["2026 К2 Итог", "", 70],
  ["2026 К3", "2026-07", 30], ["2026 К3 Итог", "", 30],
  ["Общий итог", "", 450]
];

test("dates grouped by quarter and month: helper columns are written, checked, and the pivot goes over them", async () => {
  const state = ordersSheet({ grid: DATES, numberFormat: "dd.mm.yyyy", evaluate: excelDates, buildLayout: QUARTERS_MONTHS });
  setUndoMonitorReady(true);
  try {
    const plan = await prepareCreatePivotPlan({
      sheet: "Заказы", sourceAddress: "A1:C6", rows: ["дата"], values: [{ field: "Сумма" }], groupDates: { field: "Дата", by: ["month", "quarter"] }
    });
    assert.deepEqual(plan.rowFields, ["Дата (квартал)", "Дата (месяц)"], "крупный уровень первым, на месте поля дат");
    assert.equal(plan.dateGroups?.address, "D1:E6");
    assert.equal(plan.pivotSourceAddress, "A1:E6");
    assert.equal(plan.destCell, "G1", "сводная — правее вспомогательных столбцов, а не поверх них");
    assert.ok(plan.preview.some((line) => /Вспомогательные столбцы D1:E6: «Дата \(квартал\)», «Дата \(месяц\)»/.test(line)));

    const result = await executeCreatePivotPlan(plan) as any;
    assert.equal(result.executionState, "verified");
    assert.equal(state.written.get("D1"), "Дата (квартал)");
    assert.equal(state.written.get("D2"), '=IF(A2="","",YEAR(A2)&" К"&ROUNDUP(MONTH(A2)/3,0))');
    assert.equal(state.written.get("E6"), '=IF(A6="","",YEAR(A6)&"-"&TEXT(MONTH(A6),"00"))');

    await undoLast();
    assert.equal(state.pivots.length, 0);
    assert.equal(state.written.size, 0, "отмена убрала и вспомогательные столбцы");
  } finally {
    clearUndo();
    setUndoMonitorReady(false);
  }
});

test("if the helper formulas give something else, they are removed and no pivot is built", async () => {
  const state = ordersSheet({ grid: DATES, numberFormat: "dd.mm.yyyy", evaluate: (formula) => (formula.includes("TEXT") ? "2026-1" : excelDates(formula)), buildLayout: QUARTERS_MONTHS });
  const plan = await prepareCreatePivotPlan({ sheet: "Заказы", sourceAddress: "A1:C6", rows: ["Дата"], values: [{ field: "Сумма" }], groupDates: { field: "Дата", by: ["month"] } });
  await assert.rejects(() => executeCreatePivotPlan(plan), (error: any) => {
    assert.equal(error.executionState, "failed_before_write");
    assert.match(error.message, /«2026-1» вместо «2026-01».*Столбцы убраны, сводная не строилась/);
    return true;
  });
  assert.equal(state.written.size, 0);
  assert.equal(state.pivots.length, 0);
});

test("dates the panel cannot group are refused before the card", async () => {
  const base = { sheet: "Заказы", sourceAddress: "A1:C6", values: [{ field: "Сумма" }] };
  ordersSheet({ grid: DATES.map((row, index) => (index === 2 ? ["03.02.2026", ...row.slice(1)] : row)), numberFormat: "dd.mm.yyyy" });
  await assert.rejects(() => prepareCreatePivotPlan({ ...base, rows: ["Дата"], groupDates: { field: "Дата", by: ["month"] } }), /записаны текстом.*«03\.02\.2026»/);
  ordersSheet({ grid: DATES, numberFormat: "General" });
  await assert.rejects(() => prepareCreatePivotPlan({ ...base, rows: ["Дата"], groupDates: { field: "Дата", by: ["month"] } }), /числа без формата даты/);
  ordersSheet({ grid: DATES, numberFormat: "dd.mm.yyyy" });
  await assert.rejects(() => prepareCreatePivotPlan({ ...base, rows: ["Город"], groupDates: { field: "Дата", by: ["month"] } }), /поставьте это поле в rows или columns/);
  await assert.rejects(() => prepareCreatePivotPlan({ ...base, rows: ["Город"], columns: ["Дата"], groupDates: { field: "Дата", by: ["quarter", "month"] } }), /одно поле/);
  ordersSheet({ grid: DATES, numberFormat: "dd.mm.yyyy", occupied: "D3" });
  await assert.rejects(() => prepareCreatePivotPlan({ ...base, rows: ["Дата"], groupDates: { field: "Дата", by: ["month"] } }), /правее источника.*D1:D6.*1 непустых/);
});

test("10.7: a pivot over a whole table is built by the table; name, label, format and totals are applied after the check", async () => {
  // Срез 10.7: сводная по таблице подхватывает новые строки после обновления;
  // подпись «Сумма» совпадает с заголовком — Excel примет её только с пробелом;
  // формат поля Excel читает в записи своего языка (замер 07.10.2026).
  const state = ordersSheet();
  state.tables.push({ name: "ТЗаказы", address: "Заказы!A1:D7" });
  const plan = await prepareCreatePivotPlan({
    ...SUMS, name: "ВыручкаГорода", grandTotals: "none",
    values: [{ field: "Сумма", label: "Сумма", numberFormat: '#,##0,"к"' }]
  });
  assert.equal(plan.sourceTable, "ТЗаказы");
  assert.equal(plan.name, "ВыручкаГорода");
  assert.ok(plan.preview.some((line) => /таблица «ТЗаказы»/.test(line)));
  assert.ok(plan.preview.some((line) => /Без общих итогов/.test(line)));
  const result = await executeCreatePivotPlan(plan) as any;
  assert.equal(result.executionState, "verified");
  assert.deepEqual(state.pivots[0].source, { kind: "table", name: "ТЗаказы" });
  const field = state.pivots[0].dataHierarchies.items[0];
  assert.equal(field.name, "Сумма ");
  assert.equal(field.numberFormat, '# ##0 "к"');
  assert.equal(state.pivots[0].layout.showRowGrandTotals, false);
  assert.equal(state.pivots[0].layout.autoFormat, false, "ширина столбцов не прыгает при обновлении");
  assert.equal(result.sourceTable, "ТЗаказы");
});

test("10.7: an English number format becomes the workbook's own spelling", async () => {
  const { toLocalNumberFormat, acceptedValueLabel } = await import("./pivotFinish");
  assert.equal(toLocalNumberFormat('#,##0,"к"', ",", " "), '# ##0 "к"');
  assert.equal(toLocalNumberFormat("#,##0.00", ",", " "), "# ##0,00");
  assert.equal(toLocalNumberFormat('#,##0 "руб., коп."', ",", " "), '# ##0 "руб., коп."', "текст в кавычках не трогается");
  assert.equal(toLocalNumberFormat("[Red]#,##0.0", ",", " "), "[Red]# ##0,0");
  assert.equal(toLocalNumberFormat("#,##0.00", ".", ","), "#,##0.00", "английский Excel — без изменений");
  assert.equal(acceptedValueLabel("Выручка", ["Регион", "Выручка"]), "Выручка ");
  assert.equal(acceptedValueLabel("Итого, ₽", ["Регион", "Выручка"]), "Итого, ₽");
});

test("10.9: a new pivot with four tuned value fields is refused before Excel (Excel 2021 crashes on it)", async () => {
  ordersSheet();
  const values = [{ field: "Сумма", label: "Сумма " }, { field: "Сумма" }, { field: "Сумма" }, { field: "Сумма" }];
  await assert.rejects(() => prepareCreatePivotPlan({ sheet: "Заказы", sourceAddress: "A1:D7", rows: ["Город"], values }), /4 полей значений.*Excel падает/s);
});

test("08.10: a field goes into the pivot's Filters area right at creation; filters on a field outside the pivot point to it", async () => {
  // «Книга602»: агент хотел «Экспортер» в фильтры, дважды получил отказ и сдался.
  ordersSheet();
  setUndoMonitorReady(true);
  try {
    await assert.rejects(() => prepareCreatePivotPlan({ sheet: "Заказы", sourceAddress: "A1:D7", rows: ["Город"], values: [{ field: "Сумма" }], filters: [{ field: "Менеджер", include: ["Иванов"] }] }),
      /передай filterFields: \["Менеджер"\]/);
    await assert.rejects(() => prepareCreatePivotPlan({ sheet: "Заказы", sourceAddress: "A1:D7", rows: ["Город"], values: [{ field: "Сумма" }], filterFields: ["Город"] }), /уже в строках или столбцах/);
    const plan = await prepareCreatePivotPlan({ sheet: "Заказы", sourceAddress: "A1:D7", rows: ["Город"], values: [{ field: "Сумма" }], filterFields: ["менеджер"] });
    assert.ok(plan.preview.some((line: string) => /Поля в «Фильтрах» сводной: «Менеджер»/.test(line)));
    const result = await executeCreatePivotPlan(plan) as any;
    assert.equal(result.executionState, "verified");
    assert.deepEqual(result.finish.filterFields, ["Менеджер"]);
  } finally {
    setUndoMonitorReady(false);
  }
});

test("08.10: with a field in columns the order is set first — Excel would sort by the first column, not the grand total", async () => {
  // «Книга602»: область сортировки — «Год постав. = 2000», строки вразнобой.
  const state = ordersSheet();
  const plan = await prepareCreatePivotPlan({
    sheet: "Заказы", sourceAddress: "A1:D7", rows: ["Город"], columns: ["Статус"], values: [{ field: "Сумма" }], sort: { field: "Город", order: "desc" }
  });
  await executeCreatePivotPlan(plan).catch(() => undefined);
  const sortAt = state.calls.findIndex((call: any) => call.sort);
  const columnAt = state.calls.findIndex((call: any) => call.column);
  assert.ok(sortAt >= 0 && columnAt > sortAt, JSON.stringify(state.calls));
});
