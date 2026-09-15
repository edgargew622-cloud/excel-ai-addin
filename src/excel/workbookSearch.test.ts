import test from "node:test";
import assert from "node:assert/strict";
import { decodeSearchCursor, encodeSearchCursor, searchTextMatches, ToolError } from "./excelTools";

test("search matching supports partial, exact and case-sensitive modes", () => {
  assert.equal(searchTextMatches("Москва Север", "москва"), true);
  assert.equal(searchTextMatches("Москва Север", "москва", true), false);
  assert.equal(searchTextMatches("Москва", "москва", false, true), true);
  assert.equal(searchTextMatches("Москва Север", "москва", false, true), false);
  assert.equal(searchTextMatches(null, "null"), false);
});

test("search continuation is opaque, bounded and tied to request fingerprint", () => {
  const cursor = encodeSearchCursor("0123abcd", { sheetIndex: 2, cellOffset: 9001 });
  assert.deepEqual(decodeSearchCursor(cursor, "0123abcd"), { sheetIndex: 2, cellOffset: 9001 });
  assert.throws(() => decodeSearchCursor(cursor, "deadbeef"), ToolError);
  assert.throws(() => decodeSearchCursor("v1:0123abcd:-1:2", "0123abcd"), ToolError);
  assert.throws(() => decodeSearchCursor("not-a-cursor", "0123abcd"), ToolError);
});
