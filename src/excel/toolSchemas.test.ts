import test from "node:test";
import assert from "node:assert/strict";
import { TOOL_BY_NAME, toolsForApi } from "./toolSchemas";

function names(analysisOnly: boolean): string[] {
  return toolsForApi(analysisOnly).map((tool) => tool.function.name);
}

test("analysis mode exposes no mutating tools", () => {
  for (const name of names(true)) assert.equal(TOOL_BY_NAME.get(name)?.mutating, false, name);
});

test("stage 3 exposes set_range_values as the only write path", () => {
  const mutating = names(false).filter((name) => TOOL_BY_NAME.get(name)?.mutating);
  assert.deepEqual(mutating, ["set_range_values"]);
});

test("context inspection tools are available in analysis mode", () => {
  const available = names(true);
  for (const name of ["get_active_context", "list_sheets", "get_sheet_overview", "get_range_values", "search_workbook", "get_range_details", "recall_snapshot"]) {
    assert.ok(available.includes(name), name);
  }
});
