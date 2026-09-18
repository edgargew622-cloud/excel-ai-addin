import test from "node:test";
import assert from "node:assert/strict";
import {
  executeConditionalFormatPlan,
  executeCreateTablePlan,
  executeFreezePanesPlan,
  prepareConditionalFormatPlan,
  prepareCreateTablePlan,
  prepareFreezePanesPlan
} from "./sheetFormatPlans";
import { PLANNED_TOOLS } from "./plans";
import { clear as clearUndo, depth as undoDepth, setUndoMonitorReady, undoLast } from "./undo";

const letter = (index: number) => String.fromCharCode(64 + index);

/**
 * Лист «Сотрудники»: шапка и шесть строк, как в format-smoke.xlsx. Макет
 * хранит закрепление, правила условного форматирования и таблицы честно,
 * а режимы порчи показывают случаи, когда Excel сделал не то, что просили.
 */
function staffSheet(options: {
  headers?: unknown[];
  fillIgnored?: boolean;
  existingTable?: string;
  freezeIgnored?: boolean;
  /** Ручное оформление, оставшееся с прошлых проверок. */
  headerFill?: string;
  insideBorder?: string;
} = {}) {
  const headers = options.headers ?? ["ФИО", "Должность", "Отдел", "Оклад", "Комментарий"];
  const grid: unknown[][] = [
    headers,
    ["Иванов Иван", "Ведущий специалист", "Продажи", 185000, "Перевод"],
    ["Петрова Анна", "Бухгалтер", "Финансы", 95000, "—"],
    ["Сидоров Пётр", "Руководитель", "Развитие", 240000, "Испытательный срок"],
    ["Кузнецова Мария", "Аналитик", "Финансы", 120000, "—"],
    ["Смирнов Олег", "Инженер", "Производство", 110000, "Отпуск"],
    ["Попова Елена", "Юрист", "Правовой", 150000, "—"]
  ];
  const freeze = { rows: 0, columns: 0 };
  const rules: any[] = [];
  const tables: any[] = [];
  let nextRule = 1;

  if (options.existingTable) {
    tables.push({
      name: "Старая",
      getRange: () => ({ address: `Сотрудники!${options.existingTable}`, load: () => undefined })
    });
  }

  function makeRange(address: string): any {
    const match = /^([A-Z])(\d+):([A-Z])(\d+)$/.exec(address)!;
    const columnIndex = match[1].charCodeAt(0) - 65;
    const rowIndex = Number(match[2]) - 1;
    const columnCount = match[3].charCodeAt(0) - 65 - columnIndex + 1;
    const rowCount = Number(match[4]) - rowIndex;
    const slice = () => Array.from({ length: rowCount }, (_, r) =>
      Array.from({ length: columnCount }, (_, c) => grid[rowIndex + r]?.[columnIndex + c] ?? ""));
    return {
      address: `Сотрудники!${address}`,
      rowIndex,
      columnIndex,
      rowCount,
      columnCount,
      load: () => undefined,
      get values() { return slice(); },
      get formulas() { return slice(); },
      format: {
        protection: { locked: false, load: () => undefined },
        fill: { color: "#FFFFFF", load: () => undefined },
        borders: { getItem: () => ({ style: options.insideBorder ?? "None", load: () => undefined }) }
      },
      getRow: () => ({ format: { fill: { color: options.headerFill ?? "#FFFFFF", load: () => undefined } } }),
      getOffsetRange: () => ({
        getResizedRange: () => ({ format: { fill: { color: "#FFFFFF", load: () => undefined } } })
      }),
      conditionalFormats: {
        get items() { return rules.map((rule) => ({ id: rule.id, type: rule.type })); },
        load: () => undefined,
        add: (type: string) => {
          const fill = {
            _color: null as string | null,
            get color() { return this._color; },
            set color(value: string | null) { this._color = options.fillIgnored ? "#FFFFFF" : value; },
            load: () => undefined
          };
          const rule: any = {
            id: `rule-${nextRule++}`,
            type,
            // Так легло в файле при проверке: новое правило встаёт последним.
            priority: rules.length,
            load: () => undefined,
            cellValue: { format: { fill, font: {} }, rule: null, load: () => undefined },
            textComparison: { format: { fill, font: {} }, rule: null, load: () => undefined },
            colorScale: { criteria: null, load: () => undefined },
            dataBar: { positiveFormat: { fillColor: null, load: () => undefined } }
          };
          rules.push(rule);
          return rule;
        },
        getItem: (id: string) => ({
          delete: () => { rules.splice(rules.findIndex((rule) => rule.id === id), 1); }
        })
      }
    };
  }

  const sheet: any = {
    id: "sheet-1",
    name: "Сотрудники",
    load: () => undefined,
    protection: { protected: false, load: () => undefined },
    autoFilter: { enabled: false, load: () => undefined },
    freezePanes: {
      getLocationOrNullObject: () => {
        const { rows, columns } = freeze;
        const address = rows && columns ? `Сотрудники!A1:${letter(columns)}${rows}`
          : rows ? `Сотрудники!$1:$${rows}`
          : columns ? `Сотрудники!$A:$${letter(columns)}`
          : "";
        return { isNullObject: !rows && !columns, address, load: () => undefined };
      },
      unfreeze: () => { if (!options.freezeIgnored) { freeze.rows = 0; freeze.columns = 0; } },
      freezeRows: (count: number) => { if (!options.freezeIgnored) freeze.rows = count; },
      freezeColumns: (count: number) => { if (!options.freezeIgnored) freeze.columns = count; },
      freezeAt: (range: any) => {
        if (options.freezeIgnored) return;
        freeze.rows = range.rowCount;
        freeze.columns = range.columnCount;
      }
    },
    getRangeByIndexes: (row: number, column: number, rows: number, columns: number) =>
      ({ rowIndex: row, columnIndex: column, rowCount: rows, columnCount: columns, load: () => undefined }),
    getRange: (address: string) => makeRange(address.replace(/\$/g, "")),
    tables: {
      get items() { return tables; },
      load: () => undefined,
      add: (address: string) => {
        const table: any = {
          name: `Таблица${tables.length + 1}`,
          style: "TableStyleMedium2",
          showHeaders: true,
          load: () => undefined,
          getRange: () => ({ address: `Сотрудники!${address}`, load: () => undefined }),
          // Так Excel поступает с шапкой: пустые заголовки получают имена.
          getHeaderRowRange: () => ({
            values: [grid[0].map((value, index) => (value === "" || value === null ? `Столбец${index + 1}` : String(value)))],
            load: () => undefined
          })
        };
        tables.push(table);
        return table;
      }
    }
  };

  (globalThis as any).Office = {
    context: { document: { url: "C:/format-smoke.xlsx" }, requirements: { isSetSupported: () => true } }
  };
  (globalThis as any).Excel = {
    run: async (fn: any) => fn({
      workbook: {
        application: { calculationMode: "automatic", load: () => undefined },
        worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet },
        tables: {
          getItemOrNullObject: (name: string) => ({
            isNullObject: !tables.some((table) => table.name === name),
            load: () => undefined
          })
        }
      },
      sync: async () => undefined
    })
  };
  return { grid, freeze, rules, tables };
}

test("the three new tools go through the plan registry", () => {
  for (const name of ["freeze_panes", "add_conditional_format", "create_table"]) {
    assert.ok(PLANNED_TOOLS.includes(name), name);
  }
});

/* --- закрепление ---------------------------------------------------------- */

test("freezing the header row is verified by where Excel says the panes are", async () => {
  const state = staffSheet();
  const plan = await prepareFreezePanesPlan({ sheet: "Сотрудники", rows: 1 });
  assert.equal(plan.beforeText, "ничего не закреплено");
  assert.equal(plan.expectedText, "строки 1–1");

  const result = await executeFreezePanesPlan(plan) as any;
  assert.equal(result.executionState, "verified");
  assert.deepEqual(state.freeze, { rows: 1, columns: 0 });
  assert.equal(result.after, "строки 1–1");
});

test("rows and columns together are frozen at one corner", async () => {
  const state = staffSheet();
  const plan = await prepareFreezePanesPlan({ sheet: "Сотрудники", rows: 1, columns: 1 });
  const result = await executeFreezePanesPlan(plan) as any;
  assert.equal(result.executionState, "verified");
  assert.deepEqual(state.freeze, { rows: 1, columns: 1 });
});

test("asking for the freeze already in place changes nothing and says so", async () => {
  const state = staffSheet();
  state.freeze.rows = 1;
  await assert.rejects(() => prepareFreezePanesPlan({ sheet: "Сотрудники", rows: 1 }), /уже закреплено/);
});

test("a manual freeze between preview and confirmation cancels the operation", async () => {
  const state = staffSheet();
  const plan = await prepareFreezePanesPlan({ sheet: "Сотрудники", rows: 1 });
  state.freeze.columns = 2;
  await assert.rejects(() => executeFreezePanesPlan(plan), (error: any) => {
    assert.equal(error.executionState, "failed_before_write");
    return true;
  });
});

test("a freeze Excel ignored is reported, not counted as done", async () => {
  staffSheet({ freezeIgnored: true });
  const plan = await prepareFreezePanesPlan({ sheet: "Сотрудники", rows: 1 });
  await assert.rejects(() => executeFreezePanesPlan(plan), (error: any) => {
    assert.match(error.message, /не изменилось/);
    assert.equal(error.executionState, "applied");
    return true;
  });
});

test("undo puts the earlier freeze back", async () => {
  const state = staffSheet();
  setUndoMonitorReady(true);
  try {
    const plan = await prepareFreezePanesPlan({ sheet: "Сотрудники", rows: 1 });
    const result = await executeFreezePanesPlan(plan) as any;
    assert.equal(result.undoable, true);
    assert.match(await undoLast(), /закрепление на листе Сотрудники/);
    assert.deepEqual(state.freeze, { rows: 0, columns: 0 });
  } finally {
    clearUndo();
    setUndoMonitorReady(false);
  }
});

/* --- условное форматирование ------------------------------------------------ */

test("a highlight rule predicts its matches and is verified by reading the rule back", async () => {
  const state = staffSheet();
  const plan = await prepareConditionalFormatPlan({
    sheet: "Сотрудники", address: "D2:D7", rule: "greaterThan", value: 150000, fillColor: "#FFC7CE"
  });
  assert.equal(plan.prediction?.matches, 2, "185 000 и 240 000");
  assert.deepEqual(plan.prediction?.sample, ["D2", "D4"]);
  assert.match(plan.prediction?.note ?? "", /оценка панели/);

  const result = await executeConditionalFormatPlan(plan) as any;
  assert.equal(result.executionState, "verified");
  assert.equal(result.rulesAfter, 1);
  assert.equal(state.rules[0].cellValue.rule.operator, "GreaterThan");
  assert.equal(state.rules[0].cellValue.rule.formula1, "150000");
});

test("a rule Excel stored with the wrong colour is named as a mismatch", async () => {
  staffSheet({ fillIgnored: true });
  const plan = await prepareConditionalFormatPlan({
    sheet: "Сотрудники", address: "D2:D7", rule: "greaterThan", value: 150000, fillColor: "#FFC7CE"
  });
  await assert.rejects(() => executeConditionalFormatPlan(plan), (error: any) => {
    assert.match(error.message, /цвет заливки/);
    assert.equal(error.executionState, "applied");
    return true;
  });
});

test("an existing rule is announced, and one added meanwhile stops the operation", async () => {
  const state = staffSheet();
  const first = await prepareConditionalFormatPlan({ sheet: "Сотрудники", address: "D2:D7", rule: "dataBar" });
  await executeConditionalFormatPlan(first);

  const second = await prepareConditionalFormatPlan({
    sheet: "Сотрудники", address: "D2:D7", rule: "lessThan", value: 100000, bold: true
  });
  assert.match(second.existingNote ?? "", /не заменит/);

  // Пока висит предпросмотр, кто-то добавил ещё одно правило руками.
  state.rules.push({ id: "manual", type: "CellValue" });
  await assert.rejects(() => executeConditionalFormatPlan(second), (error: any) => {
    assert.equal(error.executionState, "failed_before_write");
    return true;
  });
});

test("undo removes exactly the rule that was added", async () => {
  const state = staffSheet();
  setUndoMonitorReady(true);
  try {
    const plan = await prepareConditionalFormatPlan({
      sheet: "Сотрудники", address: "C2:C7", rule: "textContains", text: "финанс", fillColor: "#C6EFCE"
    });
    assert.equal(plan.prediction?.matches, 2, "две строки отдела «Финансы»");
    await executeConditionalFormatPlan(plan);
    assert.equal(state.rules.length, 1);
    assert.equal(undoDepth(), 1);
    await undoLast();
    assert.equal(state.rules.length, 0);
  } finally {
    clearUndo();
    setUndoMonitorReady(false);
  }
});

/* --- таблица -------------------------------------------------------------- */

test("a table is created with the chosen style and checked against the plan", async () => {
  staffSheet();
  const plan = await prepareCreateTablePlan({ sheet: "Сотрудники", address: "A1:E7", style: "TableStyleMedium9" });
  assert.deepEqual(plan.headerProblems, []);
  assert.equal(plan.undoAvailable, false);
  assert.match(plan.undoNote, /Преобразовать в диапазон/);

  const result = await executeCreateTablePlan(plan) as any;
  assert.equal(result.executionState, "verified");
  assert.equal(result.style, "TableStyleMedium9");
  assert.equal(result.address, "A1:E7");
  assert.equal(result.renamedHeaders, undefined);
});

test("headers Excel renames are warned about before and named after", async () => {
  staffSheet({ headers: ["ФИО", "", "Отдел", "Оклад", "Комментарий"] });
  const plan = await prepareCreateTablePlan({ sheet: "Сотрудники", address: "A1:E7" });
  assert.ok(plan.headerProblems.some((text) => /пустой заголовок/.test(text)));

  const result = await executeCreateTablePlan(plan) as any;
  assert.deepEqual(result.renamedHeaders, [{ column: 2, before: "", after: "Столбец2" }]);
  assert.match(result.renamedNote, /Назови это пользователю/);
});

test("a table overlapping another is refused before anything is attempted", async () => {
  staffSheet({ existingTable: "D1:F4" });
  await assert.rejects(() => prepareCreateTablePlan({ sheet: "Сотрудники", address: "A1:E7" }), /пересекается с таблицей Старая/);
});

test("data edited between preview and confirmation stops the table", async () => {
  const state = staffSheet();
  const plan = await prepareCreateTablePlan({ sheet: "Сотрудники", address: "A1:E7" });
  state.grid[3][3] = 999;
  await assert.rejects(() => executeCreateTablePlan(plan), (error: any) => {
    assert.equal(error.executionState, "failed_before_write");
    return true;
  });
  assert.equal(state.tables.length, 0, "таблица не создана");
});

test("a style that does not exist is refused up front", async () => {
  staffSheet();
  await assert.rejects(() => prepareCreateTablePlan({ sheet: "Сотрудники", address: "A1:E7", style: "Синий" }), /не существует/);
});

test("the priority a rule actually got is reported, not promised in advance", async () => {
  staffSheet();
  const first = await prepareConditionalFormatPlan({
    sheet: "Сотрудники", address: "D2:D7", rule: "greaterThan", value: 150000, fillColor: "#FFC7CE"
  });
  await executeConditionalFormatPlan(first);

  // Проверка 18 сентября 2026 года: карточка обещала, что победит новое
  // правило, а в файле шкала встала второй и под прежней заливкой не видна.
  const second = await prepareConditionalFormatPlan({
    sheet: "Сотрудники", address: "D2:D7", rule: "colorScale", minColor: "#F8696B", maxColor: "#63BE7B"
  });
  assert.doesNotMatch(second.existingNote ?? "", /добавленное позже/);
  const result = await executeConditionalFormatPlan(second) as any;
  assert.equal(result.priority, 2);
  assert.match(result.priorityNote, /действуют прежние правила выше него/);
});

test("manual formatting that will hide the table style is named before creation", async () => {
  staffSheet({ headerFill: "#1F4E79", insideBorder: "Continuous" });
  const plan = await prepareCreateTablePlan({ sheet: "Сотрудники", address: "A1:E7" });
  assert.match(plan.manualFormattingWarning ?? "", /заливка шапки, границы ячеек/);

  staffSheet();
  const clean = await prepareCreateTablePlan({ sheet: "Сотрудники", address: "A1:E7" });
  assert.equal(clean.manualFormattingWarning, undefined);
});
