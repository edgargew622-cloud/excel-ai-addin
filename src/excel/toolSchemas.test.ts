import test from "node:test";
import assert from "node:assert/strict";
import { TOOL_BY_NAME, toolsForApi } from "./toolSchemas";

function names(analysisOnly: boolean): string[] {
  return toolsForApi(analysisOnly).map((tool) => tool.function.name);
}

test("analysis mode exposes no mutating tools", () => {
  for (const name of names(true)) assert.equal(TOOL_BY_NAME.get(name)?.mutating, false, name);
});

test("writing is limited to the single verified path and a group of the same", () => {
  const mutating = names(false).filter((name) => TOOL_BY_NAME.get(name)?.mutating);
  // Этап 3 открыл одиночную запись, этап 5 добавил группу таких же записей.
  // Всё прочее, что меняет книгу, остаётся закрытым до своей проверки.
  assert.deepEqual(mutating.sort(), ["set_range_values", "set_ranges_values"]);
});

test("structural and formatting tools stay closed to the model", () => {
  const exposed = new Set(names(false));
  for (const name of ["insert_rows", "delete_rows", "sort_range", "apply_filter", "create_pivot_table", "create_chart", "format_range"]) {
    assert.equal(exposed.has(name), false, name);
  }
});

test("context inspection tools are available in analysis mode", () => {
  const available = names(true);
  for (const name of ["get_active_context", "list_sheets", "get_sheet_overview", "get_range_values", "search_workbook", "get_range_details", "recall_snapshot"]) {
    assert.ok(available.includes(name), name);
  }
});
