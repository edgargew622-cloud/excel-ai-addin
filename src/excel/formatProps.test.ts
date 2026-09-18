import test from "node:test";
import assert from "node:assert/strict";
import {
  charsToPoints,
  digitWidthFrom,
  edgesOf,
  expectedBorders,
  expectedFormatSnapshot,
  formatDifferences,
  formatSnapshotsEqual,
  parseFormatRequest,
  pointsToChars,
  requestedFormatKeys
} from "./formatProps";
import { executeFormatRangePlan, prepareFormatRangePlan } from "./excelTools";
import { clear as clearUndo, setUndoMonitorReady } from "./undo";

test("the arguments become a request, checked before anything reaches Excel", () => {
  const { request, autofit } = parseFormatRequest({
    italic: true,
    fontColor: "1f4e79",
    fontSize: 12,
    horizontalAlignment: "Center",
    borders: "all",
    borderWeight: "Medium",
    autofit: "columns"
  });
  assert.equal(request.italic, true);
  assert.equal(request.fontColor, "#1F4E79", "цвет приведён к одному виду");
  assert.deepEqual(request.borders, { mode: "all", weight: "Medium" });
  assert.equal(autofit, "columns");

  assert.throws(() => parseFormatRequest({ fontSize: 0 }), /от 1 до 409/);
  assert.throws(() => parseFormatRequest({ fontColor: "синий" }), /HEX/);
  assert.throws(() => parseFormatRequest({ fontName: "  " }), /пустой/);
  // Цвет рамки без самой рамки — непонятно, что рисовать.
  assert.throws(() => parseFormatRequest({ borderColor: "#FF0000" }), /вместе с borders/);
  // Явная ширина и автоподбор затёрли бы друг друга.
  assert.throws(() => parseFormatRequest({ columnWidth: 80, autofit: "columns" }), /одновременно/);
  assert.throws(() => parseFormatRequest({ rowHeight: 20, autofit: "both" }), /одновременно/);
  // А вот ширина с автоподбором высоты — разные измерения, конфликта нет.
  assert.doesNotThrow(() => parseFormatRequest({ columnWidth: 80, autofit: "rows" }));
});

test("the first three properties keep their old order in the preview", () => {
  assert.deepEqual(
    requestedFormatKeys({ fillColor: "#FFFFFF", italic: true, numberFormat: "0", bold: true }),
    ["numberFormat", "bold", "fillColor", "italic"]
  );
});

test("a range only has the inside borders its shape allows", () => {
  assert.deepEqual(edgesOf({ rowCount: 1, columnCount: 1 }, "every"), ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"]);
  assert.deepEqual(edgesOf({ rowCount: 3, columnCount: 1 }, "inside"), ["InsideHorizontal"]);
  assert.deepEqual(edgesOf({ rowCount: 3, columnCount: 4 }, "inside"), ["InsideHorizontal", "InsideVertical"]);
  assert.equal(edgesOf({ rowCount: 3, columnCount: 4 }, "outline").length, 4);
});

test("border expectations name every edge the mode touches, and only those", () => {
  const shape = { rowCount: 5, columnCount: 3 };
  const all = expectedBorders({ mode: "all" }, shape);
  assert.equal(Object.keys(all).length, 6);
  assert.equal(all.EdgeTop, "Continuous|Thin|#000000");

  const outline = expectedBorders({ mode: "outline", color: "#ff0000", weight: "Thick" }, shape);
  assert.deepEqual(Object.keys(outline), ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"]);
  assert.equal(outline.EdgeLeft, "Continuous|Thick|#FF0000");

  const none = expectedBorders({ mode: "none" }, shape);
  assert.ok(Object.values(none).every((value) => value === "None"));
});

test("outline leaves inside borders out of the comparison", () => {
  const expected = expectedFormatSnapshot({ borders: { mode: "outline" } }, { rowCount: 3, columnCount: 3 });
  // Внутри оставались старые линии — рамка их не трогает и сверять не должна.
  const actual = {
    borders: {
      EdgeTop: "Continuous|Thin|#000000",
      EdgeBottom: "Continuous|Thin|#000000",
      EdgeLeft: "Continuous|Thin|#000000",
      EdgeRight: "Continuous|Thin|#000000",
      InsideHorizontal: "Continuous|Hairline|#C0C0C0",
      InsideVertical: "None"
    }
  };
  assert.equal(formatSnapshotsEqual(actual, expected), true);
});

test("sizes snap to the screen grid without counting as a mismatch", () => {
  const expected = expectedFormatSnapshot({ columnWidth: 100, rowHeight: 20 });
  assert.equal(formatSnapshotsEqual({ columnWidth: 99.75, rowHeight: 20.25 }, expected), true);
  assert.deepEqual(formatDifferences({ columnWidth: 60, rowHeight: 20 }, expected), ["columnWidth"]);
  // Неоднородная ширина — не «примерно та же».
  assert.deepEqual(formatDifferences({ columnWidth: null, rowHeight: 20 }, expected), ["columnWidth"]);
});

test("underline is asked as yes or no and read back the same way", () => {
  const expected = expectedFormatSnapshot({ underline: true });
  assert.equal(expected.underline, true);
});

/**
 * Макет области, который хранит всё оформление честно: шрифт, выравнивание,
 * размеры и каждую из шести границ. Режимы порчи показывают случаи, когда
 * Excel применил не то, что просили.
 */
function styledExcel(options: {
  rows?: number;
  columns?: number;
  bordersIgnored?: boolean;
  autofitWidth?: number;
  standardWidth?: number;
} = {}) {
  const rowCount = options.rows ?? 3;
  const columnCount = options.columns ?? 2;
  const font: Record<string, unknown> = { bold: false, italic: false, underline: "None", color: "#000000", size: 11, name: "Calibri" };
  const format: Record<string, unknown> = {
    horizontalAlignment: "General",
    verticalAlignment: "Bottom",
    wrapText: false,
    columnWidth: 48,
    rowHeight: 15
  };
  const borders: Record<string, { style: string; weight: string; color: string }> = {};
  for (const edge of ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"]) {
    borders[edge] = { style: "None", weight: "Thin", color: "#000000" };
  }
  const calls: string[] = [];
  // Счётчик обращений к ячейкам: на целых столбцах их миллионы, и обход
  // всех подвешивает Excel. Макет обрывает такой обход, а не ждёт его конца.
  const cellCalls = { count: 0 };

  const rangeFormat: any = {
    load: () => undefined,
    protection: { locked: false, load: () => undefined },
    font: new Proxy(font, {
      get: (target, key) => (key === "load" ? () => undefined : target[key as string]),
      set: (target, key, value) => { target[key as string] = value; return true; }
    }),
    fill: { color: "#FFFFFF", load: () => undefined, clear() { this.color = "#FFFFFF"; } },
    borders: {
      getItem: (edge: string) => ({
        load: () => undefined,
        get style() { return borders[edge].style; },
        set style(value: string) { if (!options.bordersIgnored) borders[edge].style = value; },
        get weight() { return borders[edge].weight; },
        set weight(value: string) { if (!options.bordersIgnored) borders[edge].weight = value; },
        get color() { return borders[edge].color; },
        set color(value: string) { if (!options.bordersIgnored) borders[edge].color = value; }
      })
    },
    autofitColumns: () => { calls.push("autofitColumns"); format.columnWidth = options.autofitWidth ?? 72.75; },
    autofitRows: () => { calls.push("autofitRows"); }
  };
  for (const key of Object.keys(format)) {
    Object.defineProperty(rangeFormat, key, {
      get: () => format[key],
      set: (value) => { format[key] = value; },
      enumerable: true
    });
  }
  const range: any = {
    address: "Данные!A1:B3",
    rowCount: rowCount,
    columnCount: columnCount,
    rowIndex: 0,
    columnIndex: 0,
    load: () => undefined,
    format: rangeFormat,
    getColumn: (index: number) => ({
      address: `Данные!${String.fromCharCode(65 + index)}1:${String.fromCharCode(65 + index)}${rowCount}`,
      load: () => undefined,
      format: { load: () => undefined, get columnWidth() { return format.columnWidth; } }
    }),
    getRow: () => ({ format: { load: () => undefined, get rowHeight() { return format.rowHeight; } } }),
    getCell: () => {
      cellCalls.count += 1;
      if (cellCalls.count > 10_000) throw new Error("Обход ячеек целых столбцов: Excel бы завис.");
      return range;
    }
  };
  const sheet: any = {
    id: "sheet-1",
    name: "Данные",
    load: () => undefined,
    protection: { protected: false, load: () => undefined },
    standardWidth: options.standardWidth,
    getRange: () => range,
    getRangeByIndexes: () => ({ load: () => undefined, values: [["x"]], text: [["x"]] })
  };
  (globalThis as any).Excel = {
    run: async (fn: any) => fn({
      workbook: {
        application: { calculationMode: "automatic", load: () => undefined },
        worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet }
      },
      sync: async () => undefined
    })
  };
  return { font, format, borders, calls, cellCalls };
}

test("a table header gets font, alignment and wrapping in one verified step", async () => {
  const state = styledExcel();
  const plan = await prepareFormatRangePlan({
    sheet: "Данные",
    address: "A1:B3",
    italic: true,
    fontColor: "#1F4E79",
    fontSize: 12,
    fontName: "Arial",
    underline: true,
    horizontalAlignment: "Center",
    verticalAlignment: "Center",
    wrapText: true
  });
  assert.equal(plan.before.italic, false);
  assert.equal(plan.before.underline, false, "«None» из Excel читается как «нет»");

  const result = await executeFormatRangePlan(plan) as any;
  assert.equal(result.executionState, "verified");
  assert.equal(state.font.italic, true);
  assert.equal(state.font.underline, "Single");
  assert.equal(state.font.name, "Arial");
  assert.equal(state.format.horizontalAlignment, "Center");
  assert.equal(state.format.wrapText, true);
  assert.equal(result.actual.fontColor, "#1F4E79");
});

test("a grid of borders is drawn on every edge and checked edge by edge", async () => {
  const state = styledExcel();
  const plan = await prepareFormatRangePlan({ sheet: "Данные", address: "A1:B3", borders: "all", borderWeight: "Medium" });
  assert.equal(plan.before.borders && (plan.before.borders as any).EdgeTop, "None");

  const result = await executeFormatRangePlan(plan) as any;
  assert.equal(result.executionState, "verified");
  for (const edge of Object.keys(state.borders)) {
    assert.equal(state.borders[edge].style, "Continuous", edge);
    assert.equal(state.borders[edge].weight, "Medium", edge);
  }
});

test("borders that Excel silently ignored are reported, not counted as drawn", async () => {
  styledExcel({ bordersIgnored: true });
  const plan = await prepareFormatRangePlan({ sheet: "Данные", address: "A1:B3", borders: "outline" });
  await assert.rejects(() => executeFormatRangePlan(plan), (error: any) => {
    assert.match(error.message, /не дало эффекта/);
    assert.equal(error.executionState, "applied");
    return true;
  });
});

test("autofit reports the widths Excel chose instead of pretending to know them", async () => {
  const state = styledExcel({ autofitWidth: 72.75 });
  const plan = await prepareFormatRangePlan({ sheet: "Данные", address: "A1:B3", bold: true, autofit: "columns" });
  assert.match(plan.autofitNote ?? "", /заранее его не узнать/);

  const result = await executeFormatRangePlan(plan) as any;
  assert.equal(result.executionState, "verified");
  assert.deepEqual(state.calls, ["autofitColumns"]);
  assert.equal(result.sizesBefore.columnWidths[0].width, 48);
  assert.equal(result.sizesAfter.columnWidths[0].width, 72.75);
  assert.equal(result.sizesAfter.columnWidths[0].column, "A");
});

test("autofit alone is a valid request, with nothing else to change", async () => {
  styledExcel();
  const plan = await prepareFormatRangePlan({ sheet: "Данные", address: "A1:B3", autofit: "both" });
  assert.deepEqual(plan.expected, {});
  const result = await executeFormatRangePlan(plan) as any;
  assert.equal(result.executionState, "verified");
});

test("width in characters is measured from the workbook, not guessed", () => {
  // Calibri 11: стандартные 8,43 знака занимают 48 пунктов — цифра в 7 пикселей.
  assert.equal(Math.round(digitWidthFrom(8.43, 48) * 100) / 100, 7);
  // Проверка 18 сентября 2026 года: модель считала 60 пунктов за 30–35 знаков.
  assert.equal(pointsToChars(60, 7), 10.7);
  assert.equal(charsToPoints(30, 7), 161.25);
  // Туда и обратно — те же знаки.
  assert.equal(pointsToChars(charsToPoints(12, 7), 7), 12);
  // Неудачное измерение не превращается в дикое соотношение.
  assert.equal(digitWidthFrom(undefined, 48), 7);
  assert.equal(digitWidthFrom(8.43, 1), 7);
  assert.equal(pointsToChars(0, 7), 0, "скрытый столбец — ноль знаков");
});

test("width is asked once: characters and points together are refused", () => {
  assert.throws(() => parseFormatRequest({ columnWidth: 80, columnWidthChars: 12 }), /дважды/);
  assert.throws(() => parseFormatRequest({ columnWidthChars: 12, autofit: "columns" }), /одновременно/);
  assert.equal(parseFormatRequest({ columnWidthChars: 12 }).columnWidthChars, 12);
});

test("a width in characters is written in points and reported in both", async () => {
  const state = styledExcel({ standardWidth: 8.43 });
  const plan = await prepareFormatRangePlan({ sheet: "Данные", address: "A1:B3", columnWidthChars: 30 });
  // Цифра измерена по книге: 8,43 знака в 48 пунктах — это 6,999 пикселя, не ровно 7.
  assert.equal(plan.request.columnWidth, charsToPoints(30, digitWidthFrom(8.43, 48)));
  assert.equal(plan.columnWidthChars?.requested, 30);
  assert.equal(plan.columnWidthChars?.before, 8.4, "стандартные 48 пунктов — это 8,4 знака");

  const result = await executeFormatRangePlan(plan) as any;
  assert.equal(result.executionState, "verified");
  assert.equal(state.format.columnWidth, plan.request.columnWidth);
  assert.equal(result.columnWidthChars.actual, 30);
});

test("whole columns can be resized, because widths belong to columns, not cells", async () => {
  styledExcel({ rows: 1_048_576, columns: 5 });
  // Проверка 18 сентября 2026 года: «подбери ширину A:E» упиралась в предел ячеек.
  const plan = await prepareFormatRangePlan({ sheet: "Данные", address: "A:E", autofit: "columns" });
  const result = await executeFormatRangePlan(plan) as any;
  assert.equal(result.executionState, "verified");
  assert.equal(result.sizesAfter.columnWidths.length, 5);
});

test("whole columns still refuse cell formatting and a million row heights", async () => {
  styledExcel({ rows: 1_048_576, columns: 5 });
  await assert.rejects(
    () => prepareFormatRangePlan({ sheet: "Данные", address: "A:E", bold: true }),
    /для области данных/
  );
  await assert.rejects(
    () => prepareFormatRangePlan({ sheet: "Данные", address: "A:E", autofit: "rows" }),
    /не более чем у 1000 строк/
  );
});

test("undo of a whole-column autofit never walks the cells of those columns", async () => {
  // Проверка 18 сентября 2026 года: автоподбор ширины A:E повесил Excel.
  // Снимок отмены обходил все ячейки целых столбцов, хотя снимать с них было нечего.
  const state = styledExcel({ rows: 1_048_576, columns: 5 });
  setUndoMonitorReady(true);
  try {
    const plan = await prepareFormatRangePlan({ sheet: "Данные", address: "A:E", autofit: "columns" });
    assert.equal(plan.undoAvailable, true, "пять столбцов — отмена по силам");
    const result = await executeFormatRangePlan(plan) as any;
    assert.equal(result.executionState, "verified");
    assert.equal(result.undoable, true);
    assert.equal(state.cellCalls.count, 0, "ни одной ячейки не тронуто");
  } finally {
    clearUndo();
    setUndoMonitorReady(false);
  }
});
