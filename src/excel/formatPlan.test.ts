import test from "node:test";
import assert from "node:assert/strict";
import {
  executeFormatRangePlan,
  expectedFormatSnapshot,
  formatSnapshotsEqual,
  prepareFormatRangePlan,
  requestedFormatKeys,
  sameFormatValue
} from "./excelTools";
import { PLANNED_TOOLS, planDriverFor } from "./plans";

/** Макет Excel для оформления. Свойства читаются и пишутся честно, поэтому
 * видно и успешный путь, и случаи, когда Excel молча ничего не применил. */
function formatExcel(options: {
  bold?: unknown;
  fillColor?: unknown;
  numberFormat?: unknown;
  protectedSheet?: boolean;
  locked?: boolean | null;
  ignoreWrites?: boolean;
  normalizeTo?: { bold?: unknown; fillColor?: unknown };
  failOnApply?: boolean;
  cells?: number;
} = {}) {
  // Именно "in", а не ??: null — это ответ Office.js о неоднородности области,
  // и подменять его значением по умолчанию нельзя.
  const state: Record<string, unknown> = {
    bold: "bold" in options ? options.bold : false,
    fillColor: "fillColor" in options ? options.fillColor : "#FFFFFF",
    numberFormat: "numberFormat" in options ? options.numberFormat : "General"
  };
  const columns = options.cells ?? 1;
  const range: any = {
    address: "Данные!A1",
    rowCount: 1,
    columnCount: columns,
    rowIndex: 0,
    columnIndex: 0,
    load: () => undefined,
    get numberFormat() { return [[state.numberFormat]]; },
    set numberFormat(value: any) {
      if (options.failOnApply) throw new Error("Excel отказал");
      if (!options.ignoreWrites) state.numberFormat = value[0][0];
    },
    format: {
      protection: { locked: options.locked ?? false, load: () => undefined },
      font: {
        load: () => undefined,
        get bold() { return state.bold; },
        set bold(value: unknown) {
          if (options.failOnApply) throw new Error("Excel отказал");
          if (!options.ignoreWrites) state.bold = options.normalizeTo?.bold ?? value;
        }
      },
      fill: {
        load: () => undefined,
        get color() { return state.fillColor; },
        set color(value: unknown) {
          if (options.failOnApply) throw new Error("Excel отказал");
          if (!options.ignoreWrites) state.fillColor = options.normalizeTo?.fillColor ?? value;
        }
      }
    }
  };
  const sheet: any = {
    id: "sheet-1",
    name: "Данные",
    load: () => undefined,
    protection: { protected: options.protectedSheet ?? false, load: () => undefined },
    getRange: () => range,
    getRangeByIndexes: () => range
  };
  (globalThis as any).Excel = {
    run: async (fn: any) => fn({
      workbook: {
        application: { calculationMode: "automatic", load: () => undefined },
        worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet }
      },
      sync: async () => undefined
    })
  };
  return { state, range };
}

test("only the requested properties are touched", () => {
  assert.deepEqual(requestedFormatKeys({ bold: true }), ["bold"]);
  assert.deepEqual(requestedFormatKeys({ numberFormat: "0.00", fillColor: "#FF0000" }), ["numberFormat", "fillColor"]);
  // false — это осмысленное значение, а не отсутствие просьбы.
  assert.deepEqual(requestedFormatKeys({ bold: false }), ["bold"]);
  assert.deepEqual(requestedFormatKeys({}), []);
});

test("a plan with nothing to change is refused", async () => {
  formatExcel();
  await assert.rejects(
    () => prepareFormatRangePlan({ sheet: "Данные", address: "A1" }),
    /менять нечего/
  );
});

test("the preview shows what is now and what it will become", async () => {
  formatExcel({ bold: false, fillColor: "#FFFFFF" });
  const plan = await prepareFormatRangePlan({ sheet: "Данные", address: "A1", bold: true });

  assert.deepEqual(plan.before, { bold: false });
  assert.deepEqual(plan.expected, { bold: true });
  assert.equal(plan.cellCount, 1);
  // План неизменяем: подменить цель после показа нельзя.
  assert.throws(() => { (plan as any).address = "B2"; });
});

test("mixed formatting is kept as mixed, not as missing", async () => {
  // Office.js отдаёт null, когда свойство различается внутри области.
  formatExcel({ bold: null });
  const plan = await prepareFormatRangePlan({ sheet: "Данные", address: "A1", bold: true });
  assert.equal(plan.before.bold, null);
});

test("a protected target is refused before anything changes", async () => {
  const excel = formatExcel({ protectedSheet: true, locked: true });
  await assert.rejects(
    () => prepareFormatRangePlan({ sheet: "Данные", address: "A1", bold: true }),
    /защищён/
  );
  assert.equal(excel.state.bold, false);
});

test("a manual change between preview and execution stops the operation", async () => {
  const excel = formatExcel({ bold: false });
  const plan = await prepareFormatRangePlan({ sheet: "Данные", address: "A1", bold: true });

  // Пользователь сам сделал ячейку полужирной, пока смотрел предпросмотр.
  excel.state.bold = true;
  await assert.rejects(() => executeFormatRangePlan(plan), (error: any) => {
    assert.match(error.message, /изменилось после предпросмотра/);
    assert.equal(error.executionState, "failed_before_write");
    return true;
  });
});

test("a verified change reports what it did", async () => {
  const excel = formatExcel({ bold: false });
  const plan = await prepareFormatRangePlan({ sheet: "Данные", address: "A1", bold: true });
  const result = await executeFormatRangePlan(plan) as any;

  assert.equal(result.executionState, "verified");
  assert.equal(excel.state.bold, true);
  assert.deepEqual(result.applied, { bold: true });
  assert.deepEqual(result.before, { bold: false });
});

test("a change that Excel silently ignored is named as such", async () => {
  formatExcel({ bold: false, ignoreWrites: true });
  const plan = await prepareFormatRangePlan({ sheet: "Данные", address: "A1", bold: true });

  await assert.rejects(() => executeFormatRangePlan(plan), (error: any) => {
    // Повтор здесь бесполезен, и агенту нужно сказать именно это.
    assert.match(error.message, /не дало эффекта/);
    assert.match(error.message, /Повтор ничего не изменит/);
    assert.equal(error.executionState, "applied");
    return true;
  });
});

test("a result different from the plan is not passed off as success", async () => {
  // Excel принял просьбу, но привёл значение к своему виду.
  formatExcel({ fillColor: "#FFFFFF", normalizeTo: { fillColor: "#FF0001" } });
  const plan = await prepareFormatRangePlan({ sheet: "Данные", address: "A1", fillColor: "#FF0000" });

  await assert.rejects(() => executeFormatRangePlan(plan), (error: any) => {
    assert.match(error.message, /отличается от плана/);
    assert.equal(error.executionState, "applied");
    return true;
  });
});

test("a failure during application is undecided, not a plain error", async () => {
  formatExcel({ failOnApply: true });
  const plan = await prepareFormatRangePlan({ sheet: "Данные", address: "A1", bold: true });

  await assert.rejects(() => executeFormatRangePlan(plan), (error: any) => {
    // Доказать, что Excel ничего не успел, нельзя.
    assert.equal(error.executionState, "unknown");
    assert.match(error.message, /Перечитайте оформление/);
    return true;
  });
});

test("colours are compared the way Excel writes them back", () => {
  assert.equal(sameFormatValue("#ff0000", "#FF0000"), true, "регистр цвета не важен");
  assert.equal(sameFormatValue(true, true), true);
  assert.equal(sameFormatValue(null, false), false, "разное в области — не то же, что выключено");
  assert.equal(formatSnapshotsEqual({ bold: true }, { bold: true }), true);
  assert.equal(formatSnapshotsEqual({ bold: true }, { bold: null }), false);
  assert.equal(expectedFormatSnapshot({ bold: false }).bold, false);
});

test("every planned tool is registered, and formatting is now one of them", () => {
  assert.ok(PLANNED_TOOLS.includes("format_range"));
  for (const name of PLANNED_TOOLS) {
    const driver = planDriverFor(name);
    assert.ok(driver, name);
    // Отказ и остановка должны уметь отпустить удержанное для любой операции.
    assert.equal(typeof driver!.release, "function", name);
  }
  assert.equal(planDriverFor("get_range_values"), undefined, "чтение через план не идёт");
});

test("an escaped literal is the same number format, not a mismatch", async () => {
  // Excel сохраняет код 0.00 ₽ как 0.00 \₽. Прежде это останавливало задачу.
  const { canonicalFormatText } = await import("./excelTools");
  assert.equal(sameFormatValue("0.00 ₽", "0.00 \₽"), true);
  assert.equal(sameFormatValue('0.00 "₽"', "0.00 \₽"), true, "кавычки тоже литерал");
  assert.equal(sameFormatValue("0.00", "0.000"), false, "разные форматы остаются разными");
  assert.equal(canonicalFormatText('#,##0 "руб."'), "#,##0 руб.");
});

test("a normalised format no longer stops a successful operation", async () => {
  formatExcel({ numberFormat: "General" });
  // Excel переписывает код по-своему, но формат тот же.
  const excel = (globalThis as any).Excel;
  const plan = await prepareFormatRangePlan({ sheet: "Данные", address: "A1", numberFormat: "0.00 ₽" });
  (globalThis as any).Excel = excel;
  const state = formatExcel({ numberFormat: "General" });
  Object.defineProperty(state.range, "numberFormat", {
    get: () => [[state.state.numberFormat]],
    set: () => { state.state.numberFormat = "0.00 \₽"; },
    configurable: true
  });
  const result = await executeFormatRangePlan(plan) as any;
  assert.equal(result.executionState, "verified");
});

test("the report is grounded in the changed cells themselves", async () => {
  const excel = formatExcel({ bold: false });
  excel.range.values = [[2]];
  excel.range.text = [["2"]];
  const plan = await prepareFormatRangePlan({ sheet: "Данные", address: "A1", bold: true });
  const result = await executeFormatRangePlan(plan) as any;
  // Модель пишет отчёт по значениям из ответа, а не по памяти о соседних столбцах.
  assert.deepEqual(result.sampleValues, [[2]]);
  assert.deepEqual(result.sampleText, [["2"]]);
});

test("a currency symbol with a locale code is the same symbol", async () => {
  const { canonicalFormatText } = await import("./excelTools");
  assert.equal(sameFormatValue("#,##0 ₽", "#,##0 [$₽-419]"), true);
  assert.equal(sameFormatValue("#,##0 ₽", "#,##0\ [$₽-419]"), true, "вместе с экранированным пробелом");
  assert.equal(canonicalFormatText("[$€-407] #,##0.00"), "€ #,##0.00");
  assert.equal(sameFormatValue("#,##0 ₽", "#,##0 $"), false, "другая валюта остаётся другой");
});
