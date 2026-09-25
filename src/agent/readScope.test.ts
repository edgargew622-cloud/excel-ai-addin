import test from "node:test";
import assert from "node:assert/strict";
import { lastUserRequest, mentionsSheet, ReadScope, sheetOfReference, sheetsReadBy, type ScopeIO } from "./readScope";

const io = (names: Record<string, string> = {}, all = ["Public", "Secret", "Итоги"]): ScopeIO => ({
  allSheets: async () => all,
  sheetOfAddress: async (sheet, address) => names[address] ?? sheet
});

test("a sheet counts as named only as a whole word, in any case", () => {
  assert.equal(mentionsSheet("Посчитай итоги на листе Public", "public"), true);
  assert.equal(mentionsSheet("смотри «Итоги»", "Итоги"), true);
  assert.equal(mentionsSheet("лист Лист10", "Лист1"), false);
  assert.equal(mentionsSheet("лист Лист1, пожалуйста", "Лист1"), true);
  assert.equal(mentionsSheet("Publication", "Public"), false);
});

test("the active sheet and sheets named in the request are open; others are not", () => {
  const scope = new ReadScope("Public", "Сравни с листом Итоги");
  assert.deepEqual(scope.outside(["Public", "Итоги", "Secret", "secret"]), ["Secret"]);
  scope.allow(["Secret"]);
  assert.deepEqual(scope.outside(["SECRET"]), []);
});

test("a named range is followed to the sheet it really points at", async () => {
  const sheets = await sheetsReadBy("get_range_values", { sheet: "Public", address: "АудитКод" }, io({ АудитКод: "Secret" }));
  assert.deepEqual(sheets, ["Public", "Secret"]);
  assert.deepEqual(await sheetsReadBy("get_range_values", { sheet: "Public", address: "A1:C4" }, io()), ["Public", "Public"]);
});

test("a search without a sheet list reads every sheet; audit checks name their sheets", async () => {
  assert.deepEqual(await sheetsReadBy("search_workbook", { query: "код" }, io()), ["Public", "Secret", "Итоги"]);
  assert.deepEqual(await sheetsReadBy("search_workbook", { query: "код", sheets: ["Public"] }, io()), ["Public"]);
  assert.deepEqual(await sheetsReadBy("audit_workbook", { sheet: "Public", checks: ["'Secret'!B9:C9"] }, io()), ["Public", "Secret"]);
  assert.equal(sheetOfReference("'Лист 2'!A1"), "Лист 2");
});

test("tools that show no sheet contents need no permission", async () => {
  for (const name of ["list_sheets", "get_active_context", "recall_snapshot", "measure_workbook_export"]) {
    assert.deepEqual(await sheetsReadBy(name, { sheet: "Secret" }, io()), [], name);
  }
});

test("the request is the latest user message", () => {
  assert.equal(lastUserRequest([{ role: "user", content: "первая" }, { role: "assistant", content: "ответ" }, { role: "user", content: "вторая" }]), "вторая");
});
