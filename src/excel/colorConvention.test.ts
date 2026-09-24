import test from "node:test";
import assert from "node:assert/strict";
import { cellRole, executeConventionPlan, headerYears, prepareConventionPlan } from "./colorConvention";
import { PLANNED_TOOLS } from "./plans";
import { parseA1Rect } from "./a1";
import { clear as clearUndo, setUndoMonitorReady, undoLast } from "./undo";

test("a cell's role comes from what is in it", () => {
  assert.equal(cellRole(1200, "Модель"), "input");
  assert.equal(cellRole("=1000", "Модель"), "input");
  assert.equal(cellRole("=12*4", "Модель"), "input");
  assert.equal(cellRole("=B2*1.1", "Модель"), "formula");
  assert.equal(cellRole("=SUM(B2:B5)", "Модель"), "formula");
  assert.equal(cellRole("=Ставка*B2", "Модель"), "formula");
  assert.equal(cellRole("=Допущения!B3", "Модель"), "link");
  assert.equal(cellRole("='Исходные данные'!B3*2", "Модель"), "link");
  assert.equal(cellRole("=Модель!B3", "Модель"), "formula");
  assert.equal(cellRole("=[Бюджет.xlsx]Лист1!A1", "Модель"), "link");
  assert.equal(cellRole('="Итого: "&B2', "Модель"), "formula");
  assert.equal(cellRole("Выручка", "Модель"), null);
  assert.equal(cellRole("", "Модель"), null);
  assert.equal(cellRole(true, "Модель"), null);
});

test("years in the header row are labels, other numbers are not", () => {
  assert.deepEqual(headerYears(["Показатель", 2025, 2026]), [1, 2]);
  assert.deepEqual(headerYears(["Показатель", 2025, 1200]), []);
  assert.deepEqual(headerYears(["Показатель", "План", "Факт"]), []);
  assert.deepEqual(headerYears([2025.5]), []);
});

test("the color convention goes through the plan registry", () => {
  assert.ok(PLANNED_TOOLS.includes("apply_color_convention"));
});

/** Лист модели: формулы и цвет текста каждой ячейки хранятся честно. */
function modelSheet(options: { fontIgnored?: string[]; ruleFontColor?: string } = {}) {
  const formulas: Record<string, unknown> = {
    A1: "Показатель", B1: 2025, C1: 2026,
    A2: "Выручка", B2: 1000, C2: "=B2*(1+Допущения!B1)",
    A3: "Затраты", B3: 600, C3: "=B3*1.05",
    A4: "Прибыль", B4: "=B2-B3", C4: "=C2-C3",
    A5: "Проверка", B5: "=B4-(B2-B3)", C5: "=C4-(C2-C3)"
  };
  const colors: Record<string, string> = { B3: "#FF0000" };
  const col = (n: number) => String.fromCharCode(64 + n);
  function rangeFor(address: string): any {
    const rect = parseA1Rect(address.replace(/^.*!/, ""))!;
    const name = (r: number, c: number) => `${col(rect.columnStart + c)}${rect.rowStart + r}`;
    const rowCount = rect.rowEnd - rect.rowStart + 1;
    const columnCount = rect.columnEnd - rect.columnStart + 1;
    const cellRange = (key: string) => ({
      load: () => undefined,
      format: {
        font: {
          load: () => undefined,
          get color() { return colors[key] ?? "#000000"; },
          set color(value: string) { if (!options.fontIgnored?.includes(key)) colors[key] = value.toUpperCase(); }
        }
      }
    });
    return {
      address: `Модель!${address}`,
      rowIndex: rect.rowStart - 1, columnIndex: rect.columnStart - 1, rowCount, columnCount,
      load: () => undefined,
      get formulas() { return Array.from({ length: rowCount }, (_, r) => Array.from({ length: columnCount }, (_, c) => formulas[name(r, c)] ?? "")); },
      format: { protection: { locked: false, load: () => undefined }, font: cellRange(address).format.font },
      getCell: (r: number, c: number) => cellRange(name(r, c)),
      conditionalFormats: {
        load: () => undefined,
        items: options.ruleFontColor
          ? [{
              id: "0", type: "CellValue", priority: 0, load: () => undefined,
              getRangeOrNullObject: () => ({ isNullObject: false, address: "Модель!B5:C5", load: () => undefined }),
              cellValue: { rule: { formula1: "=0", operator: "NotEqualTo" }, load: () => undefined, format: { fill: { color: null, load: () => undefined }, font: { color: options.ruleFontColor, bold: null, load: () => undefined } } }
            }]
          : []
      }
    };
  }
  const sheet: any = { id: "m1", name: "Модель", load: () => undefined, protection: { protected: false, load: () => undefined }, getRange: rangeFor };
  (globalThis as any).Office = { context: { document: { url: "C:/m.xlsx" }, requirements: { isSetSupported: () => true } } };
  (globalThis as any).Excel = { run: async (fn: any) => fn({ workbook: { worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet } }, sync: async () => undefined }) };
  return { colors, formulas };
}

test("roles, overwritten colors and conflicting rules are named before anything changes", async () => {
  modelSheet({ ruleFontColor: "#FF0000" });
  const plan = await prepareConventionPlan({ sheet: "Модель", address: "A1:C5", checks: "B5:C5" });
  assert.deepEqual(plan.counts, { input: 2, formula: 3, link: 1, check: 2 });
  assert.deepEqual(plan.overwritten, ["B3"]);
  assert.equal(plan.skipped, 7);
  assert.deepEqual(plan.yearLabels, ["B1", "C1"]);
  assert.match(plan.paletteNote, /по умолчанию/);
  assert.match(plan.conditionalNote ?? "", /значение не равно 0/);
});

test("colors are set by role, verified cell by cell, and undo brings the old ones back", async () => {
  setUndoMonitorReady(true);
  try {
    const { colors } = modelSheet();
    const plan = await prepareConventionPlan({ sheet: "Модель", address: "A1:C5", checks: "B5:C5", palette: { input: "#1F4E79" } });
    const result = await executeConventionPlan(plan) as any;
    assert.equal(result.executionState, "verified");
    assert.equal(colors.B2, "#1F4E79");
    assert.equal(colors.C2, "#008000");
    assert.equal(colors.C3, "#000000");
    assert.equal(colors.B5, "#C00000");
    assert.equal(colors.A2, undefined);
    await undoLast();
    assert.equal(colors.B3, "#FF0000");
    assert.equal(colors.B2, "#000000");
  } finally {
    clearUndo();
    setUndoMonitorReady(false);
  }
});

test("a color Excel did not take is not reported as set", async () => {
  modelSheet({ fontIgnored: ["C2"] });
  await assert.rejects(
    async () => executeConventionPlan(await prepareConventionPlan({ sheet: "Модель", address: "A1:C5" })),
    (error: any) => error.executionState === "applied" && /C2: #000000 вместо #008000/.test(error.message)
  );
});

test("content changed after the preview stops the operation", async () => {
  const { formulas } = modelSheet();
  const plan = await prepareConventionPlan({ sheet: "Модель", address: "A1:C5" });
  formulas.C3 = "=Допущения!B2";
  await assert.rejects(() => executeConventionPlan(plan), (error: any) => error.executionState === "failed_before_write");
});

test("bad palettes and checks outside the area are refused before Excel", async () => {
  modelSheet();
  await assert.rejects(() => prepareConventionPlan({ sheet: "Модель", address: "A1:C5", palette: { input: "синий" } }), /HEX/);
  await assert.rejects(() => prepareConventionPlan({ sheet: "Модель", address: "A1:C5", checks: "B9:C9" }), /внутри/);
});
