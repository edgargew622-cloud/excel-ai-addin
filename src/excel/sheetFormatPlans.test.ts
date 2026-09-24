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
  /**
   * Свойства правила, которые Excel «не принял» (S3.2): fill, font, bold,
   * operator, formula1, text, min, mid, max, bar, range.
   */
  ignore?: string[];
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
        get items() { return rules; },
        load: () => undefined,
        add: (type: string) => {
          const rule = makeRule(type, address);
          rules.push(rule);
          return rule;
        },
        getItem: (id: string) => ({
          delete: () => { rules.splice(rules.findIndex((rule) => rule.id === id), 1); }
        })
      }
    };
  }

  /**
   * Правило условного форматирования так, как его отдаёт Excel (замер
   * 24 сентября 2026 года): цвета заглавными, незаданные цвет текста
   * и жирность — null, формула условия со знаком равенства.
   */
  function makeRule(type: string, address: string): any {
    const ignored = (name: string) => options.ignore?.includes(name) === true;
    const format = () => ({
      fill: {
        _color: null as string | null,
        get color() { return this._color; },
        set color(value: string | null) {
          this._color = options.fillIgnored || ignored("fill") ? "#FFFFFF" : String(value).toUpperCase();
        },
        load: () => undefined
      },
      font: {
        _color: null as string | null,
        _bold: null as boolean | null,
        get color() { return this._color; },
        set color(value: string | null) { this._color = ignored("font") ? null : String(value).toUpperCase(); },
        get bold() { return this._bold; },
        set bold(value: boolean | null) { this._bold = ignored("bold") ? null : value; },
        load: () => undefined
      }
    });
    const cellValue: any = {
      format: format(),
      _rule: null as any,
      get rule() { return this._rule; },
      set rule(value: any) {
        this._rule = {
          formula1: ignored("formula1") ? "=0" : `=${String(value.formula1).replace(/^=/, "")}`,
          formula2: value.formula2 === undefined ? null : `=${String(value.formula2).replace(/^=/, "")}`,
          operator: ignored("operator") ? "EqualTo" : value.operator
        };
      },
      load: () => undefined
    };
    const textComparison: any = {
      format: format(),
      _rule: null as any,
      get rule() { return this._rule; },
      set rule(value: any) { this._rule = { operator: value.operator, text: ignored("text") ? "" : value.text }; },
      load: () => undefined
    };
    const colorScale: any = {
      _criteria: null as any,
      get criteria() { return this._criteria; },
      set criteria(value: any) {
        const point = (name: string, item: any) => (item ? { ...item, color: ignored(name) ? "#000000" : String(item.color).toUpperCase() } : null);
        this._criteria = {
          minimum: point("min", value.minimum),
          midpoint: ignored("mid") ? null : point("mid", value.midpoint),
          maximum: point("max", value.maximum)
        };
      },
      load: () => undefined
    };
    const positiveFormat: any = {
      _fill: null as string | null,
      get fillColor() { return this._fill; },
      set fillColor(value: string) { this._fill = ignored("bar") ? "#638EC6" : String(value).toUpperCase(); },
      load: () => undefined
    };
    const where = ignored("range") ? address.replace(/:.*$/, "") : address;
    return {
      id: `rule-${nextRule++}`,
      type,
      // Так легло в файле при проверке: новое правило встаёт последним.
      priority: rules.length,
      load: () => undefined,
      getRangeOrNullObject: () => ({ isNullObject: false, address: `Сотрудники!${where}`, load: () => undefined }),
      cellValue,
      textComparison,
      colorScale,
      dataBar: { positiveFormat, load: () => undefined }
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
  return { grid, freeze, rules, tables, makeRule };
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
  assert.equal(state.rules[0].cellValue.rule.formula1, "=150000", "Excel хранит условие со знаком равенства");
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
  const manual = state.makeRule("CellValue", "D2:D7");
  manual.cellValue.rule = { formula1: "1", operator: "GreaterThan" };
  state.rules.push(manual);
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

/* --- сверка всех запрошенных свойств (S3.2) ----------------------------------- */

const HIGHLIGHT = {
  sheet: "Сотрудники", address: "D2:D7", rule: "greaterThan", value: 150000,
  fillColor: "#FFC7CE", fontColor: "#9C0006", bold: true
};

test("every requested property of a highlight rule is read back", async () => {
  // План стабилизации, S3.2: цвет текста и жирность записывались, но не
  // сверялись — правило без них выглядело проверенным.
  for (const [property, name] of [
    ["fill", /цвет заливки/], ["font", /цвет текста/], ["bold", /жирн/],
    ["operator", /оператор/], ["formula1", /значение/], ["range", /область/]
  ] as const) {
    staffSheet({ ignore: [property] });
    const plan = await prepareConditionalFormatPlan(HIGHLIGHT);
    await assert.rejects(() => executeConditionalFormatPlan(plan), (error: any) => {
      assert.equal(error.executionState, "applied", property);
      assert.match(error.message, name, property);
      return true;
    });
  }
  staffSheet();
  const result = await executeConditionalFormatPlan(await prepareConditionalFormatPlan(HIGHLIGHT)) as any;
  assert.equal(result.executionState, "verified");
});

test("a text rule checks its text, fill, font colour and boldness", async () => {
  const request = {
    sheet: "Сотрудники", address: "C2:C7", rule: "textContains", text: "финанс",
    fillColor: "#C6EFCE", fontColor: "#006100", bold: false
  };
  // Прежде у такого правила сверялся только текст: заливка читалась,
  // но не сравнивалась.
  for (const [property, name] of [["fill", /цвет заливки/], ["font", /цвет текста/], ["bold", /жирн/], ["text", /текст условия/]] as const) {
    staffSheet({ ignore: [property] });
    const plan = await prepareConditionalFormatPlan(request);
    await assert.rejects(() => executeConditionalFormatPlan(plan), (error: any) => {
      assert.match(error.message, name, property);
      return true;
    });
  }
  staffSheet();
  assert.equal((await executeConditionalFormatPlan(await prepareConditionalFormatPlan(request)) as any).executionState, "verified");
});

test("a colour scale checks all three points, and a midpoint nobody asked for is not accepted", async () => {
  const request = {
    sheet: "Сотрудники", address: "D2:D7", rule: "colorScale",
    minColor: "#F8696B", midColor: "#FFEB84", maxColor: "#63BE7B"
  };
  for (const [property, name] of [["min", /минимум/], ["mid", /середин/], ["max", /максимум/]] as const) {
    staffSheet({ ignore: [property] });
    const plan = await prepareConditionalFormatPlan(request);
    await assert.rejects(() => executeConditionalFormatPlan(plan), (error: any) => {
      assert.match(error.message, name, property);
      return true;
    });
  }
  staffSheet();
  assert.equal((await executeConditionalFormatPlan(await prepareConditionalFormatPlan(request)) as any).executionState, "verified");
});

test("a data bar checks its colour and its area", async () => {
  for (const [property, name] of [["bar", /цвет полосы/], ["range", /область/]] as const) {
    staffSheet({ ignore: [property] });
    const plan = await prepareConditionalFormatPlan({ sheet: "Сотрудники", address: "D2:D7", rule: "dataBar", barColor: "#FF0000" });
    await assert.rejects(() => executeConditionalFormatPlan(plan), (error: any) => {
      assert.match(error.message, name, property);
      return true;
    });
  }
});

test("an existing rule edited after the preview, under the same ID, stops the operation", async () => {
  // План стабилизации, S3.2: прежние правила сверялись по числу и ID —
  // правка их содержимого проходила незамеченной.
  const state = staffSheet();
  await executeConditionalFormatPlan(await prepareConditionalFormatPlan(HIGHLIGHT));
  const plan = await prepareConditionalFormatPlan({ sheet: "Сотрудники", address: "D2:D7", rule: "dataBar" });

  state.rules[0].cellValue.format.fill.color = "#00FF00";
  await assert.rejects(() => executeConditionalFormatPlan(plan), (error: any) => {
    assert.equal(error.executionState, "failed_before_write");
    assert.match(error.message, /изменились после предпросмотра/);
    return true;
  });
  assert.equal(state.rules.length, 1, "новое правило не добавлялось");

  // То же с условием и с приоритетом.
  state.rules[0].cellValue.format.fill.color = "#FFC7CE";
  const again = await prepareConditionalFormatPlan({ sheet: "Сотрудники", address: "D2:D7", rule: "dataBar" });
  state.rules[0].cellValue.rule = { formula1: "100000", operator: "GreaterThan" };
  await assert.rejects(() => executeConditionalFormatPlan(again), (error: any) => error.executionState === "failed_before_write");
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
