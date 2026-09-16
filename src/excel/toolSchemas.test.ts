import test from "node:test";
import assert from "node:assert/strict";
import { TOOL_BY_NAME, toolsForApi } from "./toolSchemas";
import { PLANNED_TOOLS } from "./plans";

function names(analysisOnly: boolean): string[] {
  return toolsForApi(analysisOnly).map((tool) => tool.function.name);
}

test("analysis mode exposes no mutating tools", () => {
  for (const name of names(true)) assert.equal(TOOL_BY_NAME.get(name)?.mutating, false, name);
});

test("only tools that go through a plan may change the workbook", () => {
  const mutating = names(false).filter((name) => TOOL_BY_NAME.get(name)?.mutating);
  // Открыт ровно тот набор, что проходит предпросмотр и сверку результата:
  // одиночная запись, их группа и оформление. Список должен совпадать
  // с реестром планов, иначе инструмент откроется в обход проверок.
  assert.deepEqual(mutating.sort(), [...PLANNED_TOOLS].sort());
});

test("tools written before the plan machinery stay closed", () => {
  const exposed = new Set(names(false));
  for (const name of ["insert_rows", "delete_rows", "sort_range", "apply_filter", "create_pivot_table", "create_chart"]) {
    assert.equal(exposed.has(name), false, name);
  }
});

test("context inspection tools are available in analysis mode", () => {
  const available = names(true);
  for (const name of ["get_active_context", "list_sheets", "get_sheet_overview", "get_range_values", "search_workbook", "get_range_details", "recall_snapshot"]) {
    assert.ok(available.includes(name), name);
  }
});
