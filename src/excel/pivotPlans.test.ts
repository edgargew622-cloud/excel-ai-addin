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
  assert.equal(e.width, 3);
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
} = {}) {
  const grid: unknown[][] = ORDERS.map((row) => [...row]);
  const pivots: any[] = [];
  const extra = new Map<string, unknown>();
  if (options.occupied) extra.set(options.occupied, "занято");
  const emptyFormulas = new Map<string, string>(Object.entries(options.emptyFormulas ?? {}));
  const sourceFormulas = new Map<string, string>(Object.entries(options.sourceFormulas ?? {}));
  const tables: { name: string; address: string }[] = [];
  const protection = { protected: options.protectedSheet === true, load: () => undefined };
  const layouts: string[] = [];

  const cellValue = (row: number, column: number) => {
    const key = `${String.fromCharCode(65 + column)}${row + 1}`;
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
    return emptyFormulas.get(key) ?? sourceFormulas.get(key) ?? cellValue(row, column);
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
    getUsedRangeOrNullObject: () => ({ isNullObject: false, rowIndex: 0, columnIndex: 0, rowCount: 7, columnCount: 4, load: () => undefined }),
    protection,
    tables: {
      get items() {
        return tables.map((table) => ({ name: table.name, getRange: () => ({ address: table.address, load: () => undefined }) }));
      },
      load: () => { if (options.tablesThrow) throw new Error("Во время обработки запроса произошла внутренняя ошибка."); }
    },
    pivotTables: {
      get items() { return pivots.map((pivot) => ({ name: pivot.name, layout: { getRange: () => makeRange(pivot.address) } })); },
      load: () => undefined,
      add: (name: string, _source: unknown, destination: any) => {
        const fields: string[] = [];
        const data: any[] = [];
        const pivot: any = {
          name,
          row: destination.rowIndex,
          column: destination.columnIndex,
          values: [] as unknown[][],
          address: "",
          hierarchies: { getItem: (field: string) => field },
          rowHierarchies: { add: (field: string) => fields.push(field) },
          dataHierarchies: {
            add: (field: string) => {
              const item: any = { field, summarizeBy: "Sum" };
              data.push(item);
              return item;
            }
          },
          layout: {
            getRange: () => makeRange(pivot.address),
            set layoutType(value: string) { layouts.push(value); }
          }
        };
        // Так Excel сводит: по городу, сумма или — в режиме порчи — количество.
        pivot.build = () => {
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
  (globalThis as any).Office = { context: { document: { url: "C:/pivot.xlsx" }, requirements: { isSetSupported: () => true } } };
  (globalThis as any).Excel = {
    run: async (fn: any) => fn({
      workbook: {
        worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet, getItemOrNullObject: () => sheet },
        pivotTables: {
          getItemOrNullObject: (name: string) => {
            const pivot = pivots.find((item) => item.name === name);
            return { isNullObject: !pivot, load: () => undefined, delete: () => { pivots.splice(pivots.indexOf(pivot), 1); } };
          }
        }
      },
      sync: async () => { for (const pivot of pivots) if (!pivot.values.length) pivot.build(); }
    })
  };
  return { grid, pivots, extra, emptyFormulas, tables, protection, layouts };
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
  // Макет по умолчанию задаётся в настройках Excel. Размер сводной панель
  // считает для компактного макета — значит, его и нужно выставить.
  const state = ordersSheet();
  await executeCreatePivotPlan(await prepareCreatePivotPlan(SUMS));
  assert.deepEqual(state.layouts, ["Compact"]);
});
