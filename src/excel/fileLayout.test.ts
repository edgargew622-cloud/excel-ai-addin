import test from "node:test";
import assert from "node:assert/strict";
import { layoutWrite } from "./fileLayoutPlans";
import { PLANNED_TOOLS } from "./plans";
import { TOOL_BY_NAME } from "./toolSchemas";

const layout = {
  columnWidths: [30, 75, 120],
  rowHeights: [20, 20, 20],
  cells: [
    { r: 0, c: 0, rowSpan: 1, colSpan: 3, text: "Реестр", value: "Реестр", numberFormat: "@", bold: true, size: 14, align: "Center" as const, wrap: false },
    { r: 1, c: 0, rowSpan: 1, colSpan: 1, text: "0042", value: "0042", numberFormat: "@", bold: false, size: 10, align: "Left" as const, wrap: false },
    { r: 1, c: 1, rowSpan: 2, colSpan: 1, text: "01.09.2026", value: 46266, numberFormat: "dd.mm.yyyy", bold: false, size: 10, align: "Left" as const, wrap: false },
    { r: 1, c: 2, rowSpan: 1, colSpan: 1, text: "12 500,00", value: 12500, numberFormat: "#,##0.00", bold: false, size: 10, align: "Right" as const, wrap: false }
  ],
  hEdges: [[1, 0, "Thin"], [3, 2, "Medium"]] as Array<[number, number, "Thin" | "Medium"]>,
  vEdges: [[1, 0, "Thin"], [1, 3, "Thin"]] as Array<[number, number, "Thin" | "Medium"]>,
  pages: 1,
  pageStarts: [0],
  warnings: []
};

test("import_file_layout is a planned write tool that needs the local server", () => {
  assert.ok(PLANNED_TOOLS.includes("import_file_layout"));
  const spec = TOOL_BY_NAME.get("import_file_layout")!;
  assert.equal(spec.mutating && spec.destructive && spec.needsLocal, true);
});

test("the layout becomes values with their formats, merges and border edges at the right cells", () => {
  const write = layoutWrite(layout);
  assert.equal(write.block, "A1:C3");
  assert.deepEqual(write.values, [["Реестр", "", ""], ["0042", 46266, 12500], ["", "", ""]]);
  assert.deepEqual(write.formats[1], ["@", "dd.mm.yyyy", "#,##0.00"], "код «0042» — текстом, дата и сумма — числами");
  assert.deepEqual(write.merges, ["A1:C1", "B2:B3"]);
  assert.deepEqual(write.borders, [
    { cell: "A2", edge: "EdgeTop", weight: "Thin" },
    { cell: "C3", edge: "EdgeBottom", weight: "Medium" },
    { cell: "A2", edge: "EdgeLeft", weight: "Thin" },
    { cell: "C2", edge: "EdgeRight", weight: "Thin" }
  ]);
});
