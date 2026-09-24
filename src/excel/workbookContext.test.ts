import test from "node:test";
import assert from "node:assert/strict";
import { getActiveContext } from "./workbookContext";

/**
 * Книга, в которой выделено `selection`: ячейки или диаграмма. Когда
 * выделена диаграмма, Excel отказывает в активной ячейке и выделении
 * с кодом InvalidSelection — так было в проверке 24 сентября 2026 года.
 */
function workbook(selection: "cells" | "chart" | "broken") {
  const fail = (code: string) => Object.assign(new Error(code === "InvalidSelection"
    ? "Текущий выделенный фрагмент недопустим для этой операции."
    : "Внутренняя ошибка."), { code });
  const pending: (() => void)[] = [];
  const lazy = (address: string) => {
    const item: any = { load: () => pending.push(() => {
      if (selection === "chart") throw fail("InvalidSelection");
      if (selection === "broken") throw fail("GeneralException");
      item.address = address;
    }) };
    return item;
  };
  (globalThis as any).Office = { context: { document: { url: "C:/книга.xlsx" }, requirements: { isSetSupported: () => true } } };
  (globalThis as any).Excel = {
    run: async (fn: any) => fn({
      workbook: {
        worksheets: { getActiveWorksheet: () => ({ id: "sheet-1", name: "Продажи", load: () => undefined }) },
        getActiveCell: () => lazy("Продажи!B2"),
        getSelectedRanges: () => lazy("Продажи!B2:C4"),
        getActiveChartOrNullObject: () => ({ isNullObject: selection !== "chart", name: "Диаграмма 1", load: () => undefined })
      },
      sync: async () => { const jobs = pending.splice(0); for (const job of jobs) job(); }
    })
  };
}

test("cells selected: the active cell and the selection are reported", async () => {
  workbook("cells");
  const context: any = await getActiveContext();
  assert.equal(context.activeCell, "Продажи!B2");
  assert.deepEqual(context.selectedAreas, ["Продажи!B2:C4"]);
  assert.equal(context.selectionNote, undefined);
});

test("a selected chart is a fact about the workbook, not a failure of the request", async () => {
  // Проверка в Excel 24 сентября 2026 года: после create_chart диаграмма
  // оставалась выделенной, и каждая следующая просьба падала ещё до модели.
  workbook("chart");
  const context: any = await getActiveContext();
  assert.equal(context.activeSheet.name, "Продажи");
  assert.equal(context.activeCell, null);
  assert.deepEqual(context.selectedAreas, []);
  assert.match(context.selectionNote, /Выделена диаграмма «Диаграмма 1»/);
});

test("other Excel errors are not hidden behind the selection note", async () => {
  workbook("broken");
  await assert.rejects(() => getActiveContext(), /Внутренняя ошибка/);
});
