import test from "node:test";
import assert from "node:assert/strict";
import { expectChart, placementCell, seriesMismatches } from "./chartModel";
import { executeCreateChartPlan, prepareCreateChartPlan } from "./chartPlans";
import { PLANNED_TOOLS } from "./plans";
import { clear as clearUndo, setUndoMonitorReady, undoLast } from "./undo";

const SALES = [
  ["Месяц", "Выручка", "Расходы"],
  ["Январь", 120, 80],
  ["Февраль", 150, 90],
  ["Март", 170, 95]
];
const origin = { rowIndex: 0, columnIndex: 0 };

/* --- ожидание ------------------------------------------------------------- */

test("a plain table gives one series per numeric column, labels from the first", () => {
  const e = expectChart(SALES, "ColumnClustered", "columns", origin);
  assert.equal(e.headerRow, true);
  assert.equal(e.labelColumn, true);
  assert.deepEqual(e.seriesNames, ["Выручка", "Расходы"]);
  assert.equal(e.pointCount, 3);
  assert.deepEqual(e.categories, ["Январь", "Февраль", "Март"]);
  assert.deepEqual(e.warnings, []);
});

test("series by rows turn the same table around", () => {
  const e = expectChart(SALES, "Line", "rows", origin);
  assert.deepEqual(e.seriesNames, ["Январь", "Февраль", "Март"]);
  assert.equal(e.pointCount, 2);
  assert.deepEqual(e.categories, ["Выручка", "Расходы"]);
});

test("a pie with several series warns that only the first is drawn", () => {
  const e = expectChart(SALES, "Pie", "columns", origin);
  assert.ok(e.warnings.some((text) => /только первый ряд «Выручка»/.test(text)));
});

test("a text column taken into the area is named as a row of zeros", () => {
  const wide = [
    ["Месяц", "Выручка", "Комментарий"],
    ["Январь", 120, "план"],
    ["Февраль", 150, "факт"]
  ];
  const e = expectChart(wide, "ColumnClustered", "columns", origin);
  assert.ok(e.warnings.some((text) => /«Комментарий».*как нули/.test(text)));
});

test("negative shares and a series per point are both caught", () => {
  const negative = [["Статья", "Сумма"], ["Доход", 100], ["Убыток", -30]];
  assert.ok(expectChart(negative, "Pie", "columns", origin).warnings.some((text) => /отрицательные/.test(text)));

  // Одна строка данных при рядах по столбцам — ряды почти наверняка не в ту сторону.
  const oneRow = [["Январь", "Февраль", "Март"], [120, 150, 170]];
  const e = expectChart(oneRow, "ColumnClustered", "columns", origin);
  assert.ok(e.warnings.some((text) => /по одной точке/.test(text)));
});

test("without a header Excel names the series itself, so names are not compared", () => {
  const bare = [[1, 2], [3, 4]];
  const e = expectChart(bare, "Line", "columns", origin);
  assert.equal(e.headerRow, false);
  assert.deepEqual(seriesMismatches(e, { names: ["Ряд1", "Ряд2"], pointCounts: [2, 2] }), []);
  assert.deepEqual(seriesMismatches(e, { names: ["Ряд1"], pointCounts: [4] }), ["рядов 1 вместо 2", "точек в ряду 4 вместо 2"]);
});

test("the chart goes one column past the data so it covers nothing", () => {
  assert.equal(placementCell({ rowIndex: 0, columnIndex: 0, columnCount: 5 }, 0), "G1");
  assert.equal(placementCell(null, 3), "A4");
});

/* --- полный путь ------------------------------------------------------------ */

/**
 * Лист с продажами. `understands` задаёт, как Excel поймёт область:
 * «правильно» — как в ожидании, «шапка как ряд» — первая строка стала данными.
 */
function salesSheet(options: { understands?: "right" | "headerAsData" } = {}) {
  const grid: unknown[][] = SALES.map((row) => [...row]);
  const charts: any[] = [];

  const range: any = {
    address: "Продажи!A1:C4",
    rowIndex: 0,
    columnIndex: 0,
    rowCount: 4,
    columnCount: 3,
    load: () => undefined,
    get values() { return grid.map((row) => [...row]); },
    get formulas() { return grid.map((row) => [...row]); }
  };
  const sheet: any = {
    id: "sheet-1",
    name: "Продажи",
    load: () => undefined,
    protection: { protected: false, load: () => undefined },
    getUsedRangeOrNullObject: () => ({ isNullObject: false, address: "Продажи!A1:C4", rowIndex: 0, columnIndex: 0, rowCount: 4, columnCount: 3, load: () => undefined }),
    getRange: () => range,
    charts: {
      // У диаграмм есть размеры и положение в пунктах: по ним видно наложение.
      get items() { return charts; },
      load: () => undefined,
      add: (type: string, _source: unknown, seriesBy: string) => {
        const headerAsData = options.understands === "headerAsData";
        const names = seriesBy === "Rows" ? ["Январь", "Февраль", "Март"] : ["Выручка", "Расходы"];
        const points = seriesBy === "Rows" ? 2 : headerAsData ? 4 : 3;
        const chart: any = {
          id: `chart-${charts.length + 1}`,
          name: `Диаграмма ${charts.length + 1}`,
          chartType: type,
          position: "",
          top: 0,
          left: 0,
          height: 200,
          width: 300,
          load: () => undefined,
          setPosition(cell: string) {
            this.position = cell;
            // Ячейка — это место на листе: столбец даёт отступ слева, строка сверху.
            this.left = (cell.charCodeAt(0) - 65) * 60;
            this.top = (Number(cell.slice(1)) - 1) * 15;
          },
          title: { text: "", load: () => undefined },
          series: {
            load: () => undefined,
            items: (headerAsData ? ["Ряд1", "Ряд2"] : names).map((name) => ({
              name,
              points: { count: points, load: () => undefined }
            }))
          }
        };
        charts.push(chart);
        return chart;
      },
      getItemOrNullObject: (id: string) => {
        const chart = charts.find((item) => item.id === id);
        return {
          isNullObject: !chart,
          load: () => undefined,
          delete: () => { charts.splice(charts.indexOf(chart), 1); }
        };
      }
    }
  };
  (globalThis as any).Office = {
    context: { document: { url: "C:/charts.xlsx" }, requirements: { isSetSupported: () => true } }
  };
  (globalThis as any).Excel = {
    run: async (fn: any) => fn({
      workbook: { worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet } },
      sync: async () => undefined
    })
  };
  return { grid, charts };
}

test("create_chart goes through the plan registry", () => {
  assert.ok(PLANNED_TOOLS.includes("create_chart"));
});

test("a chart is placed past the data and its series are checked against the plan", async () => {
  const state = salesSheet();
  const plan = await prepareCreateChartPlan({ sheet: "Продажи", address: "A1:C4", chartType: "ColumnClustered", title: "Выручка и расходы" });
  assert.equal(plan.anchorCell, "E1");
  assert.deepEqual(plan.expectation.seriesNames, ["Выручка", "Расходы"]);

  const result = await executeCreateChartPlan(plan) as any;
  assert.equal(result.executionState, "verified");
  assert.deepEqual(result.series, ["Выручка", "Расходы"]);
  assert.equal(result.title, "Выручка и расходы");
  assert.equal(state.charts[0].position, "E1");
});

test("when Excel reads the area differently, the chart is reported as not what was planned", async () => {
  const state = salesSheet({ understands: "headerAsData" });
  setUndoMonitorReady(true);
  try {
    const plan = await prepareCreateChartPlan({ sheet: "Продажи", address: "A1:C4", chartType: "Line" });
    await assert.rejects(() => executeCreateChartPlan(plan), (error: any) => {
      assert.equal(error.executionState, "applied");
      assert.match(error.message, /понял область иначе/);
      assert.match(error.message, /точек в ряду 4\/4 вместо 3/);
      assert.match(error.message, /«Отменить»/);
      return true;
    });
    // Диаграмма уже стоит — отмена обязана её убрать.
    assert.equal(state.charts.length, 1);
    await undoLast();
    assert.equal(state.charts.length, 0);
  } finally {
    clearUndo();
    setUndoMonitorReady(false);
  }
});

test("an anchor inside the data is warned about before the chart covers it", async () => {
  salesSheet();
  const plan = await prepareCreateChartPlan({ sheet: "Продажи", address: "A1:C4", chartType: "Line", anchorCell: "B2" });
  assert.match(plan.anchorWarning ?? "", /закроет их/);
});

test("data edited after the preview stops the chart", async () => {
  const state = salesSheet();
  const plan = await prepareCreateChartPlan({ sheet: "Продажи", address: "A1:C4", chartType: "Line" });
  state.grid[2][1] = 999;
  await assert.rejects(() => executeCreateChartPlan(plan), (error: any) => {
    assert.equal(error.executionState, "failed_before_write");
    return true;
  });
  assert.equal(state.charts.length, 0);
});

test("an area without numbers is refused before anything is built", async () => {
  const state = salesSheet();
  for (const row of state.grid) row[1] = row[2] = "текст";
  await assert.rejects(() => prepareCreateChartPlan({ sheet: "Продажи", address: "A1:C4", chartType: "Line" }), /нет чисел/);
});

test("a second chart does not land on top of the first one", async () => {
  // Проверка в Excel 20 сентября 2026 года: круговая встала в ту же F1,
  // что и построенная до неё столбчатая, и легла поверх.
  const state = salesSheet();
  const first = await prepareCreateChartPlan({ sheet: "Продажи", address: "A1:C4", chartType: "ColumnClustered" });
  assert.equal(first.chartsOnSheet, 0);
  await executeCreateChartPlan(first);

  const second = await prepareCreateChartPlan({ sheet: "Продажи", address: "A1:C4", chartType: "Line" });
  assert.equal(second.chartsOnSheet, 1, "панель знает про уже стоящую диаграмму");
  const result = await executeCreateChartPlan(second) as any;

  assert.equal(result.executionState, "verified");
  assert.match(result.placementNote, /опущена под/);
  const [one, two] = state.charts;
  assert.ok(two.top >= one.top + one.height, "вторая ниже первой");
});
