import test from "node:test";
import assert from "node:assert/strict";
import {
  chartGrid,
  executeArrangeChartsPlan,
  executeEditChartPlan,
  executeFilterPivotsPlan,
  prepareArrangeChartsPlan,
  prepareEditChartPlan,
  prepareFilterPivotsPlan,
  titleCellFormula
} from "./dashboardTools";
import { PLANNED_TOOLS } from "./plans";
import { clear as clearUndo, setUndoMonitorReady, undoLast } from "./undo";

/** Лист «Дашборд» с двумя диаграммами; одна — сводная (у неё есть кнопки полей). */
function dashboard() {
  const cellSize = { width: 50, height: 20 };
  const letterIndex = (letters: string) => [...letters].reduce((total, ch) => total * 26 + ch.charCodeAt(0) - 64, 0);
  const cellBox = (address: string) => {
    const m = /^([A-Z]+)(\d+)$/.exec(address)!;
    return { left: (letterIndex(m[1]) - 1) * cellSize.width, top: (Number(m[2]) - 1) * cellSize.height };
  };
  const makeChart = (name: string, pivot: boolean, top: number) => {
    const chart: any = {
      name, chartType: "ColumnClustered", left: 400, top, width: 360, height: 216, load: () => undefined,
      title: { text: name, visible: true, load: () => undefined, setFormula(formula: string) { chart.title.formula = formula; chart.title.text = "из ячейки"; } },
      legend: { visible: true, position: "Right", load: () => undefined },
      dataLabels: { showValue: false, numberFormat: "General" },
      axes: { valueAxis: { numberFormat: "General" } },
      setPosition(from: string, to: string) {
        const a = cellBox(from); const b = cellBox(to);
        chart.left = a.left; chart.top = a.top; chart.width = b.left + cellSize.width - a.left; chart.height = b.top + cellSize.height - a.top;
      }
    };
    chart.pivotOptions = pivot
      ? { showAxisFieldButtons: true, showLegendFieldButtons: true, showReportFilterFieldButtons: true, showValueFieldButtons: true, load: () => undefined }
      : { load: () => { throw new Error("Не сводная диаграмма"); } };
    return chart;
  };
  const charts = [makeChart("Выручка", true, 300), makeChart("Доли", false, 10)];
  const filters: Record<string, any[]> = { "СвКатегории": [], "СвРегионы": [] };
  const hierarchyList = (pivotName: string, names: string[]) => ({
    load: () => undefined,
    get items() { return names.map((name) => ({ name, fields: { getItem: () => ({
      applyFilter: (filter: any) => filters[pivotName].push(filter.manualFilter.selectedItems),
      clearAllFilters: () => filters[pivotName].push("снят")
    }) } })); },
    add: (name: string) => { names.push(name); }
  });
  const pivot = (name: string, rows: string[]) => {
    const filterNames: string[] = [];
    return {
      name,
      rowHierarchies: hierarchyList(name, rows),
      columnHierarchies: hierarchyList(name, []),
      filterHierarchies: hierarchyList(name, filterNames),
      hierarchies: { load: () => undefined, items: ["Год", "Регион", "Категория", "Выручка"].map((n) => ({ name: n })), getItem: (n: string) => n },
      filterNames
    };
  };
  const pivots = [pivot("СвКатегории", ["Категория"]), pivot("СвРегионы", ["Регион"])];
  const sheet: any = {
    id: "d1", name: "Дашборд", load: () => undefined,
    charts: { load: () => undefined, get items() { return charts; } },
    getRange: (address: string) => ({ ...cellBox(address), load: () => undefined })
  };
  (globalThis as any).Office = { context: { document: { url: "C:/dash.xlsx" }, requirements: { isSetSupported: () => true } } };
  (globalThis as any).Excel = {
    run: async (fn: any) => fn({
      workbook: {
        worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet },
        pivotTables: { load: () => undefined, items: pivots, getItem: (name: string) => pivots.find((p) => p.name === name) }
      },
      application: { cultureInfo: { numberFormat: { numberDecimalSeparator: ",", numberGroupSeparator: " ", load: () => undefined } } },
      sync: async () => undefined
    })
  };
  return { charts, filters, pivots };
}

test("10.8 tools go through the plan registry", () => {
  for (const name of ["edit_chart", "arrange_charts", "filter_pivots"]) assert.ok(PLANNED_TOOLS.includes(name), name);
});

test("a chart grid sits on cell borders: two in a row, same size, one cell apart", () => {
  assert.deepEqual(chartGrid(3, { row: 4, column: 2 }, 2, 8, 15, 1), [
    { from: "B4", to: "I18" },
    { from: "K4", to: "R18" },
    { from: "B20", to: "I34" }
  ]);
  assert.equal(titleCellFormula("Дашборд!b1", "Лист1"), "='Дашборд'!$B$1");
  assert.equal(titleCellFormula("A2", "Мой лист"), "='Мой лист'!$A$2");
  assert.throws(() => titleCellFormula("A1:B2", "Лист1"), /одна ячейка/);
});

test("edit_chart: doughnut, no legend, title from a cell, field buttons hidden — and undo restores", async () => {
  // Инструкция автора: убрать легенду и кнопки полей, заголовок привязать к ячейке, сменить тип на кольцевую.
  const state = dashboard();
  setUndoMonitorReady(true);
  try {
    const plan = await prepareEditChartPlan({ chart: "выручка", chartType: "Doughnut", legend: "None", titleFromCell: "B1", fieldButtons: false });
    assert.equal(plan.titleFormula, "='Дашборд'!$B$1");
    const result = await executeEditChartPlan(plan) as any;
    const chart = state.charts[0];
    assert.equal(result.executionState, "verified");
    assert.equal(chart.chartType, "Doughnut");
    assert.equal(chart.legend.visible, false);
    assert.equal(chart.title.formula, "='Дашборд'!$B$1");
    assert.equal(chart.pivotOptions.showAxisFieldButtons, false);
    await undoLast();
    assert.equal(chart.chartType, "ColumnClustered");
    assert.equal(chart.legend.visible, true);
    assert.equal(chart.pivotOptions.showAxisFieldButtons, true);
  } finally {
    clearUndo();
    setUndoMonitorReady(false);
  }
  await assert.rejects(() => prepareEditChartPlan({ chart: "Доли", fieldButtons: false }), /не сводная диаграмма/);
  await assert.rejects(() => prepareEditChartPlan({ chart: "Нет" , legend: "None" }), /нет диаграммы «Нет»/);
});

test("arrange_charts: charts go top-down into a grid on cell borders; undo puts them back", async () => {
  const state = dashboard();
  setUndoMonitorReady(true);
  try {
    const plan = await prepareArrangeChartsPlan({ startCell: "B4", columns: 2 });
    assert.deepEqual(plan.charts, ["Доли", "Выручка"], "сверху вниз");
    const result = await executeArrangeChartsPlan(plan) as any;
    assert.equal(result.executionState, "verified");
    assert.deepEqual([state.charts[1].left, state.charts[1].top], [50, 60], "«Доли» — в B4");
    assert.equal(state.charts[0].width, state.charts[1].width, "одного размера");
    await undoLast();
    assert.deepEqual([state.charts[0].left, state.charts[0].top], [400, 300]);
  } finally {
    clearUndo();
    setUndoMonitorReady(false);
  }
});

test("filter_pivots: one selection into every pivot with the field — the slicer's missing report connection", async () => {
  const state = dashboard();
  const plan = await prepareFilterPivotsPlan({ field: "Регион", include: ["Север"] });
  assert.deepEqual(plan.pivots.map((item) => [item.name, item.axis]), [["СвКатегории", "none"], ["СвРегионы", "rows"]]);
  const result = await executeFilterPivotsPlan(plan) as any;
  assert.deepEqual(result.filtered, ["СвКатегории", "СвРегионы"]);
  assert.deepEqual(state.pivots[0].filterNames, ["Регион"], "где поля нет — оно встаёт в фильтры сводной");
  assert.deepEqual(state.filters["СвКатегории"], [["Север"]]);
  assert.deepEqual(state.filters["СвРегионы"], [["Север"]]);
  await assert.rejects(() => prepareFilterPivotsPlan({ field: "Заказчик", include: ["А"] }), /нет поля «Заказчик»/);
  await assert.rejects(() => prepareFilterPivotsPlan({ field: "Регион" }), /include/);
});

test("edit_chart: thousands with «к» go to the axis and the labels with a no-break space — Excel's chart needs it", async () => {
  // Замер 07.10.2026: «#,##0,"к"» на диаграмме давал «90000,0,к», «# ##0 "к"» с обычным пробелом — «90 000 к»; верно — с неразрывным.
  const state = dashboard();
  await executeEditChartPlan(await prepareEditChartPlan({ chart: "Выручка", numberFormat: '#,##0,"к"' }));
  const nbsp = String.fromCharCode(0xa0);
  assert.equal(state.charts[0].axes.valueAxis.numberFormat, `#${nbsp}##0${nbsp}"к"`);
  assert.equal(state.charts[0].dataLabels.numberFormat, `#${nbsp}##0${nbsp}"к"`);
});
