import test from "node:test";
import assert from "node:assert/strict";
import { executeShareGrowthPlan, prepareShareGrowthPlan, readSourceBlock, shareGrowthLayout, templateMismatches } from "./templates";
import { PLANNED_TOOLS } from "./plans";
import { parseA1Rect } from "./a1";
import { clear as clearUndo, setUndoMonitorReady, undoLast } from "./undo";

const SOURCE = [
  ["Город", 2024, 2025, 2026],
  ["Москва", 100, 110, 121],
  ["Казань", 0, 50, 60],
  ["Омск", 300, "", 330]
];

test("the source table is read as periods, items and numbers; text is refused", () => {
  const block = readSourceBlock(SOURCE, 1, 1);
  assert.deepEqual(block.periods, [2024, 2025, 2026]);
  assert.deepEqual(block.numbers[2], [300, null, 330]);
  assert.throws(() => readSourceBlock([["", 2024, 2025], ["Москва", "100", 1]], 1, 1), /B2 не число/);
  assert.throws(() => readSourceBlock([["", 2024, 2025], ["", 1, 2]], 1, 1), /подпись статьи/);
});

test("the layout holds formulas with expected values, a share control and undefined growth", () => {
  const layout = shareGrowthLayout(readSourceBlock(SOURCE, 1, 1), 6, 1);
  assert.equal(layout.rows.length, 2 * 3 + 7);
  assert.equal(layout.rows[2][1].formula, '=IF(SUM(B$2:B$4)=0,"",B2/SUM(B$2:B$4))');
  assert.equal(layout.rows[2][1].expected, 0.25);
  assert.equal(layout.rows[layout.controlRow][1].formula, "=SUM(B8:B10)");
  assert.deepEqual(layout.rows[layout.controlRow].slice(1).map((cell) => cell.expected), [1, 1, 1]);
  // Рост Москвы 2025: 110/100 − 1; у Омска в 2025 пусто — это −100 %, а 2026 от пустого не определён.
  const growthMoscow = layout.rows[layout.controlRow + 4];
  assert.ok(Math.abs((growthMoscow[2].expected as number) - 0.1) < 1e-12);
  assert.equal(growthMoscow[2].formula, '=IF(N(B2)=0,"",C2/B2-1)');
  assert.deepEqual(layout.undefinedGrowth, ["C16", "D17"]);
  assert.equal(layout.rows[layout.controlRow + 6][2].expected, -1);
});

test("the check compares numbers with a tolerance and labels as text", () => {
  const layout = shareGrowthLayout(readSourceBlock(SOURCE, 1, 1), 6, 1);
  const good = layout.rows.map((row) => row.map((cell) => cell.expected ?? ""));
  assert.deepEqual(templateMismatches(layout, good), []);
  const bad = good.map((row) => [...row]);
  bad[layout.controlRow][1] = 0.99;
  assert.deepEqual(templateMismatches(layout, bad), ["B11: 0.99 вместо 1"]);
});

test("the template goes through the plan registry", () => {
  assert.ok(PLANNED_TOOLS.includes("add_share_growth"));
});

/** Лист, где Excel «считает» формулы блока: значения берутся из ожидаемого, с возможной порчей. */
function templateSheet(options: { corrupt?: string; occupied?: boolean } = {}) {
  const cells: Record<string, { formula: unknown; format: string }> = {};
  const col = (n: number) => String.fromCharCode(64 + n);
  SOURCE.forEach((row, r) => row.forEach((value, c) => { cells[`${col(c + 1)}${r + 1}`] = { formula: value, format: "General" }; }));
  if (options.occupied) cells.B8 = { formula: "занято", format: "General" };
  const state = { expected: null as null | ((address: string) => unknown) };
  function rangeFor(address: string): any {
    const rect = parseA1Rect(address.replace(/^.*!/, ""))!;
    const rows = rect.rowEnd - rect.rowStart + 1;
    const columns = rect.columnEnd - rect.columnStart + 1;
    const key = (r: number, c: number) => `${col(rect.columnStart + c)}${rect.rowStart + r}`;
    const grid = <T>(read: (name: string) => T) => Array.from({ length: rows }, (_, r) => Array.from({ length: columns }, (_, c) => read(key(r, c))));
    return {
      address: `Модель!${address}`, rowIndex: rect.rowStart - 1, columnIndex: rect.columnStart - 1, rowCount: rows, columnCount: columns,
      load: () => undefined,
      format: { protection: { locked: false, load: () => undefined } },
      get formulas() { return grid((name) => cells[name]?.formula ?? ""); },
      set formulas(matrix: unknown[][]) { matrix.forEach((row, r) => row.forEach((value, c) => { cells[key(r, c)] = { formula: value, format: cells[key(r, c)]?.format ?? "General" }; })); },
      get numberFormat() { return grid((name) => cells[name]?.format ?? "General"); },
      set numberFormat(matrix: string[][]) { matrix.forEach((row, r) => row.forEach((value, c) => { cells[key(r, c)] = { formula: cells[key(r, c)]?.formula ?? "", format: value }; })); },
      get values() {
        return grid((name) => {
          const formula = cells[name]?.formula;
          if (typeof formula !== "string" || !formula.startsWith("=")) return formula ?? "";
          if (name === options.corrupt) return 0.5;
          return state.expected?.(name) ?? "";
        });
      }
    };
  }
  const sheet: any = {
    id: "t1", name: "Модель", load: () => undefined, protection: { protected: false, load: () => undefined },
    getRange: rangeFor,
    getUsedRangeOrNullObject: () => ({ isNullObject: false, rowIndex: 0, rowCount: options.occupied ? 8 : 4, load: () => undefined })
  };
  (globalThis as any).Office = { context: { document: { url: "C:/t.xlsx" }, requirements: { isSetSupported: () => true } } };
  (globalThis as any).Excel = { run: async (fn: any) => fn({ workbook: { worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet } }, sync: async () => undefined }) };
  return { cells, state };
}

const excelComputes = (plan: any) => (name: string) => {
  const rect = parseA1Rect(name)!;
  return plan.layout.rows[rect.rowStart - plan.layout.top]?.[rect.columnStart - plan.layout.left]?.expected ?? "";
};

test("the block is written below the table, every value verified, and undo clears it", async () => {
  setUndoMonitorReady(true);
  try {
    const { cells, state } = templateSheet();
    const plan = await prepareShareGrowthPlan({ sheet: "Модель", address: "A1:D4" });
    assert.equal(plan.destAddress, "A6:D18");
    state.expected = excelComputes(plan);
    const result = await executeShareGrowthPlan(plan) as any;
    assert.equal(result.executionState, "verified");
    assert.deepEqual(result.undefinedGrowth, ["C16", "D17"]);
    assert.equal(cells.B8.format, "0.0%");
    await undoLast();
    assert.equal(cells.B8.formula, "");
    assert.equal(cells.B8.format, "General");
  } finally {
    clearUndo();
    setUndoMonitorReady(false);
  }
});

test("a value that disagrees with the panel's own calculation is not reported as done", async () => {
  const { state } = templateSheet({ corrupt: "B11" });
  const plan = await prepareShareGrowthPlan({ sheet: "Модель", address: "A1:D4" });
  state.expected = excelComputes(plan);
  await assert.rejects(() => executeShareGrowthPlan(plan), (error: any) => error.executionState === "applied" && /B11: 0.5 вместо 1/.test(error.message));
});

test("an occupied place is refused with a free one named", async () => {
  templateSheet({ occupied: true });
  await assert.rejects(() => prepareShareGrowthPlan({ sheet: "Модель", address: "A1:D4" }), /занято.*A10/);
});
