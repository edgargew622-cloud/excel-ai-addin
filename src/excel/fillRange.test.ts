import test from "node:test";
import assert from "node:assert/strict";
import { anchorOf, executeFillRangePlan, prepareFillRangePlan } from "./excelTools";
import { PLANNED_TOOLS } from "./plans";

/** Макет столбца F2:F6 рядом с данными D и E.
 * `autoFill` повторяет поведение Excel: формула из первой ячейки протягивается
 * вниз со сдвигом номеров строк в относительных ссылках. */
function fillExcel(options: { autoFill?: boolean; occupied?: boolean; ignoreWrites?: boolean } = {}) {
  const rows = 5;
  const state = { formulas: Array.from({ length: rows }, () => [options.occupied ? "старое" : ""]) };
  const shift = (formula: string, delta: number) =>
    formula.replace(/([A-Z]+)(\d+)/g, (_m, column, row) => `${column}${Number(row) + delta}`);

  const anchor: any = {
    address: "Продажи!F2",
    load: () => undefined,
    set formulas(matrix: any[][]) { if (!options.ignoreWrites) state.formulas[0] = [matrix[0][0]]; },
    set values(matrix: any[][]) { if (!options.ignoreWrites) state.formulas[0] = [matrix[0][0]]; },
    autoFill: options.autoFill === false ? undefined : () => {
      if (options.ignoreWrites) return;
      const first = String(state.formulas[0][0]);
      for (let index = 1; index < rows; index++) state.formulas[index] = [shift(first, index)];
    }
  };
  const range: any = {
    address: "Продажи!F2:F6",
    rowCount: rows,
    columnCount: 1,
    rowIndex: 1,
    columnIndex: 5,
    load: () => undefined,
    get formulas() { return state.formulas; },
    get values() { return state.formulas; },
    format: { protection: { locked: false, load: () => undefined } },
    // В Excel autoFill есть у любого диапазона, не только у первой ячейки.
    autoFill: options.autoFill === false ? undefined : () => undefined
  };
  const sheet: any = {
    id: "sheet-1",
    name: "Продажи",
    load: () => undefined,
    protection: { protected: false, load: () => undefined },
    getRange: (address: string) => (address === "F2" ? anchor : range),
    getRangeByIndexes: () => range
  };
  (globalThis as any).Excel = {
    AutoFillType: { fillDefault: "FillDefault" },
    run: async (fn: any) => fn({
      workbook: {
        application: { calculationMode: "automatic", load: () => undefined },
        worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet }
      },
      sync: async () => undefined
    })
  };
  return state;
}

test("the anchor cell of an area is its top left corner", () => {
  assert.equal(anchorOf("F2:F6"), "F2");
  assert.equal(anchorOf("B3:D10"), "B3");
  assert.equal(anchorOf("AA5:AC9"), "AA5");
  assert.equal(anchorOf("C7"), "C7");
});

test("one formula fills the whole area and Excel adjusts the references", async () => {
  const state = fillExcel();
  const plan = await prepareFillRangePlan({ sheet: "Продажи", address: "F2:F6", value: "=D2*E2", isFormula: true });

  assert.equal(plan.cellCount, 5);
  assert.equal(plan.anchorAddress, "F2");
  assert.equal(plan.occupiedCells, 0, "область была пуста");

  const result = await executeFillRangePlan(plan) as any;
  assert.equal(result.executionState, "verified");
  // Модель передала одну формулу, а в книге они разные — по строкам.
  assert.deepEqual(state.formulas.flat(), ["=D2*E2", "=D3*E3", "=D4*E4", "=D5*E5", "=D6*E6"]);
  assert.equal(result.firstFormula, "=D2*E2");
  assert.equal(result.lastFormula, "=D6*E6");
});

test("a formula is required when isFormula is set", async () => {
  fillExcel();
  await assert.rejects(
    () => prepareFillRangePlan({ sheet: "Продажи", address: "F2:F6", value: "D2*E2", isFormula: true }),
    /начинающейся со знака равенства/
  );
});

test("the preview counts the cells that will be overwritten", async () => {
  fillExcel({ occupied: true });
  const plan = await prepareFillRangePlan({ sheet: "Продажи", address: "F2:F6", value: 0 });
  assert.equal(plan.occupiedCells, 5, "все пять ячеек заняты и будут затёрты");
});

test("an area left unchanged is named as such instead of counted a success", async () => {
  fillExcel({ ignoreWrites: true });
  const plan = await prepareFillRangePlan({ sheet: "Продажи", address: "F2:F6", value: "=D2*E2", isFormula: true });
  await assert.rejects(() => executeFillRangePlan(plan), (error: any) => {
    assert.match(error.message, /не дало эффекта/);
    assert.equal(error.executionState, "applied");
    return true;
  });
});

test("a build without fill support refuses instead of writing one cell", async () => {
  fillExcel({ autoFill: false });
  // Иначе заполнилась бы только первая ячейка, а операция считалась бы удачной.
  await assert.rejects(
    () => prepareFillRangePlan({ sheet: "Продажи", address: "F2:F6", value: "=D2*E2", isFormula: true }),
    /не поддерживает заполнение диапазона/
  );
});

test("filling goes through the plan registry like every other change", () => {
  assert.ok(PLANNED_TOOLS.includes("fill_range"));
});

test("the anchor cell is committed before the fill is asked for", async () => {
  // Excel вернул внутреннюю ошибку, когда запись и протяжка шли одним пакетом.
  const order: string[] = [];
  const state = fillExcel();
  const excel = (globalThis as any).Excel;
  const sheet = () => excel.run(async (ctx: any) => ctx);
  void sheet;
  const originalRun = excel.run;
  excel.run = async (fn: any) => originalRun(async (ctx: any) => {
    const target = ctx.workbook.worksheets.getItem();
    const anchor = target.getRange("F2");
    const originalFill = anchor.autoFill;
    anchor.autoFill = (...args: unknown[]) => { order.push("autoFill"); return originalFill?.(...args); };
    const sync = ctx.sync;
    ctx.sync = async () => { order.push("sync"); return sync(); };
    return fn(ctx);
  });

  const plan = await prepareFillRangePlan({ sheet: "Продажи", address: "F2:F6", value: "=D2*E2", isFormula: true });
  order.length = 0;
  await executeFillRangePlan(plan);

  const fillAt = order.indexOf("autoFill");
  assert.ok(fillAt > 0, "протяжка вообще случилась");
  assert.equal(order[fillAt - 1], "sync", "перед протяжкой первая ячейка уже записана");
  assert.deepEqual(state.formulas.flat(), ["=D2*E2", "=D3*E3", "=D4*E4", "=D5*E5", "=D6*E6"]);
  excel.run = originalRun;
});
