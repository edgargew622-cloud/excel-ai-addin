import test from "node:test";
import assert from "node:assert/strict";
import { supported, TOOL_BY_NAME, TOOL_SPECS, toolsForApi, writableAtCurrentStage } from "./toolSchemas";
import { PLANNED_TOOLS } from "./plans";

// Панель работает только внутри Excel: инструмент выдаётся, если Excel
// поддерживает его набор API (7.1.6). В тестах — Excel с ExcelApi 1.14,
// как на проверочной машине; тесты, которым нужен другой Excel, ставят свой.
(globalThis as any).Office ??= {
  context: { requirements: { isSetSupported: (_: string, version: string) => Number(version.split(".")[1]) <= 14 } }
};


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

test("every tool declares the Excel API it needs, and none needs more than the test machine has", async () => {
  // Этап 7, 7.1.6: прежде набор проверялся у четырёх инструментов по именам.
  const { MIN_EXCEL_API, TOOL_SPECS: TOOLS } = await import("./toolSchemas");
  for (const spec of TOOLS) {
    const version = MIN_EXCEL_API[spec.name];
    assert.ok(version, `${spec.name}: набор ExcelApi не объявлен`);
    const [major, minor] = version.split(".").map(Number);
    assert.ok(major === 1 && minor <= 14, `${spec.name}: ${version} выше 1.14`);
  }
});

test("a tool is not offered when the Excel it runs in lacks the declared API", async () => {
  const { supported, TOOL_SPECS: TOOLS } = await import("./toolSchemas");
  const previous = (globalThis as any).Office;
  (globalThis as any).Office = { context: { requirements: { isSetSupported: (_: string, version: string) => Number(version.split(".")[1]) <= 7 } } };
  try {
    const offered = TOOLS.filter((spec) => supported(spec)).map((spec) => spec.name);
    assert.ok(offered.includes("freeze_panes"), "1.7 есть");
    assert.ok(!offered.includes("create_pivot_table"), "1.8 нет — сводная не выдаётся");
    assert.ok(!offered.includes("apply_filter"), "1.9 нет — фильтр не выдаётся");
  } finally {
    (globalThis as any).Office = previous;
  }
});
