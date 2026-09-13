import test from "node:test";
import assert from "node:assert/strict";
import { cancellationToolMessages } from "./loop";
import { resolveToolArgs } from "../excel/excelTools";

test("cancelled tool calls are closed with matching tool_call_id", () => {
  const calls = [
    { id: "a", name: "set_range_values", arguments: "{}" },
    { id: "b", name: "insert_rows", arguments: "{}" }
  ];
  const messages = cancellationToolMessages(calls, 1);
  assert.equal(messages.length, 1);
  assert.equal(messages[0].role, "tool");
  assert.equal(messages[0].tool_call_id, "b");
  assert.match(String(messages[0].content), /отменена/i);
});

test("task sheet is stable unless the model explicitly names another sheet", async () => {
  const implicit = (await resolveToolArgs("set_range_values", { address: "A1", values: [[1]] }, "Sheet1")) as any;
  assert.equal(implicit.sheet, "Sheet1");

  const explicit = (await resolveToolArgs(
    "set_range_values",
    { sheet: "Sheet2", address: "A1", values: [[1]] },
    "Sheet1"
  )) as any;
  assert.equal(explicit.sheet, "Sheet2");
});
