import test from "node:test";
import assert from "node:assert/strict";
import { supported, TOOL_BY_NAME, TOOL_SPECS, toolsForApi, writableAtCurrentStage } from "./toolSchemas";
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
  // Инструменты, которых нет в этой версии Office, не выдаются вовсе, поэтому
  // сравнение идёт только с поддерживаемой частью реестра.
  const plannedAndSupported = PLANNED_TOOLS.filter((name) => {
    const spec = TOOL_BY_NAME.get(name);
    return spec ? supported(spec) : false;
  });
  assert.deepEqual(mutating.sort(), plannedAndSupported.sort());
  // И обратно: ни один выданный изменяющий инструмент не обходит план.
  for (const name of mutating) assert.ok(PLANNED_TOOLS.includes(name), name);
});

test("every mutating tool now goes through a plan: none is left behind", () => {
  // Этап 6 закрыт: инструментов, написанных до механики планов, не осталось.
  // Сводная выдаётся только там, где есть ExcelApi 1.8, поэтому проверяется
  // допуск к записи, а не выдача в среде без Excel.
  for (const spec of TOOL_SPECS.filter((item) => item.mutating)) {
    assert.ok(writableAtCurrentStage(spec), spec.name);
    assert.ok(PLANNED_TOOLS.includes(spec.name), spec.name);
  }
});

test("row operations are described as irreversible, because they are", () => {
  for (const name of ["insert_rows", "delete_rows"]) {
    const spec = TOOL_BY_NAME.get(name);
    assert.ok(spec?.mutating && spec.destructive, name);
    // Модель обязана знать из описания, что отката нет и нужна копия.
    assert.match(spec.description, /create_workbook_backup/, name);
  }
  assert.match(TOOL_BY_NAME.get("delete_rows")!.description, /необратимо/);
});

test("context inspection tools are available in analysis mode", () => {
  const available = names(true);
  for (const name of ["get_active_context", "list_sheets", "get_sheet_overview", "get_range_values", "search_workbook", "get_range_details", "recall_snapshot"]) {
    assert.ok(available.includes(name), name);
  }
});
