import test from "node:test";
import assert from "node:assert/strict";
import { executeConvertTablePlan, literalMentionsTable, prepareConvertTablePlan, tableImpact, usesOwnColumns, usesTable } from "./tablePlans";
import { PLANNED_TOOLS } from "./plans";
import { parseA1Rect } from "./a1";

test("structured references are told apart from text and from other names", () => {
  assert.equal(usesTable("=SUM(ТПроба[Сумма])", "ТПроба"), true);
  assert.equal(usesTable("=COUNTA(тпроба[#All])", "ТПроба"), true);
  assert.equal(usesTable("=SUM(ТПроба2[Сумма])", "ТПроба"), false);
  assert.equal(usesTable('=INDIRECT("ТПроба[Сумма]")', "ТПроба"), false);
  assert.equal(literalMentionsTable('=INDIRECT("ТПроба[Сумма]")', "ТПроба"), true);
  assert.equal(usesOwnColumns("=[@Сумма]*2"), true);
  assert.equal(usesOwnColumns("=SUBTOTAL(109,[Сумма])"), true);
  assert.equal(usesOwnColumns("=ТПроба[[#Headers],[Город]]"), true);
  // Внешняя книга — не столбец таблицы.
  assert.equal(usesOwnColumns("=[Книга.xlsx]Лист!A1"), false);
  assert.equal(usesOwnColumns("='[Книга.xlsx]Лист 2'!A1"), false);
  assert.equal(usesOwnColumns("=A1+B1"), false);
});

test("the impact lists formulas Excel rewrites and formulas that break", () => {
  const impact = tableImpact(
    [
      { name: "Лист", rowIndex: 0, columnIndex: 0, formulas: [["Город", "Сумма"], ["Москва", 100], ["Итог", "=SUBTOTAL(109,[Сумма])"]], values: [] },
      { name: "Итоги", rowIndex: 0, columnIndex: 0, formulas: [["=SUM(ТПроба[Сумма])"], ['=INDIRECT("ТПроба[Сумма]")'], ["=A1*2"]], values: [] }
    ],
    "ТПроба",
    "Лист",
    parseA1Rect("A1:B3")!
  );
  assert.deepEqual(impact.structured.map((item) => `${item.sheet}!${item.cell}`), ["Лист!B3", "Итоги!A1"]);
  assert.deepEqual(impact.literal.map((item) => `${item.sheet}!${item.cell}`), ["Итоги!A2"]);
});

test("table conversion goes through the plan registry", () => {
  assert.ok(PLANNED_TOOLS.includes("convert_table_to_range"));
});

/* --- книга, которая ведёт себя как Excel в замере ------------------------------------ */

type Cell = { formula: unknown; value: unknown };

function workbook(options: { filtered?: boolean; unexpected?: boolean; keepStructured?: boolean } = {}) {
  const grids: Record<string, Record<string, Cell>> = {
    Лист: {
      A1: { formula: "Город", value: "Город" }, B1: { formula: "Сумма", value: "Сумма" },
      A2: { formula: "Москва", value: "Москва" }, B2: { formula: 100, value: 100 },
      A3: { formula: "Казань", value: "Казань" }, B3: { formula: 200, value: 200 },
      A4: { formula: "Итог", value: "Итог" }, B4: { formula: "=SUBTOTAL(109,[Сумма])", value: options.filtered ? 100 : 300 }
    },
    Итоги: {
      A1: { formula: "=SUM(ТПроба[Сумма])", value: 300 },
      A2: { formula: '=INDIRECT("ТПроба[Сумма]")', value: 300 }
    }
  };
  const state = { table: true };
  const col = (n: number) => String.fromCharCode(64 + n);
  function rangeFor(sheetName: string, address: string): any {
    const rect = parseA1Rect(address.replace(/^.*!/, ""))!;
    const cells = () => Array.from({ length: rect.rowEnd - rect.rowStart + 1 }, (_, r) =>
      Array.from({ length: rect.columnEnd - rect.columnStart + 1 }, (_, c) => grids[sheetName][`${col(rect.columnStart + c)}${rect.rowStart + r}`] ?? { formula: "", value: "" }));
    return {
      isNullObject: false,
      load: () => undefined,
      get address() { return `${sheetName}!${address.replace(/^.*!/, "")}`; },
      rowIndex: rect.rowStart - 1,
      columnIndex: rect.columnStart - 1,
      rowCount: rect.rowEnd - rect.rowStart + 1,
      columnCount: rect.columnEnd - rect.columnStart + 1,
      get formulas() { return cells().map((row) => row.map((cell) => cell.formula)); },
      get values() { return cells().map((row) => row.map((cell) => cell.value)); },
      getRow: (index: number) => ({ load: () => undefined, rowHidden: Boolean(options.filtered && index === 2) })
    };
  }
  const used: Record<string, string> = { Лист: "A1:B4", Итоги: "A1:A2" };
  const sheets: Record<string, any> = {};
  for (const name of Object.keys(grids)) {
    sheets[name] = {
      id: `id-${name}`, name, load: () => undefined,
      protection: { protected: false, load: () => undefined },
      getRange: (address: string) => rangeFor(name, address),
      getUsedRangeOrNullObject: () => rangeFor(name, used[name]),
      tables: { items: [], load: () => undefined }
    };
  }
  const table: any = {
    id: "{T1}", name: "ТПроба", style: "TableStyleMedium2", showTotals: true, showHeaders: true,
    get isNullObject() { return !state.table; },
    load: () => undefined,
    worksheet: sheets.Лист,
    autoFilter: { isDataFiltered: Boolean(options.filtered), load: () => undefined },
    getRange: () => rangeFor("Лист", "A1:B4"),
    convertToRange: () => {
      state.table = false;
      grids.Лист.B4 = { formula: "=SUBTOTAL(109,Лист!$B$2:$B$3)", value: 300 };
      grids.Итоги.A1 = { formula: options.keepStructured ? "=SUM(ТПроба[Сумма])" : "=SUM(Лист!$B$2:$B$3)", value: 300 };
      grids.Итоги.A2 = { formula: '=INDIRECT("ТПроба[Сумма]")', value: "#ССЫЛКА!" };
      if (options.unexpected) grids.Лист.B2 = { formula: 100, value: 999 };
    }
  };
  const byName = (key: string) => sheets[key] ?? Object.values(sheets).find((sheet: any) => sheet.id === key);
  const wb = {
    worksheets: {
      getActiveWorksheet: () => sheets.Лист,
      getItem: byName,
      load: () => undefined,
      get items() { return Object.values(sheets); }
    },
    tables: { getItemOrNullObject: () => table, items: [table], load: () => undefined },
    names: { items: [{ name: "ИмяТ", formula: "=ТПроба[Сумма]" }], load: () => undefined }
  };
  (globalThis as any).Office = { context: { document: { url: "C:/t.xlsx" }, requirements: { isSetSupported: () => true } } };
  (globalThis as any).Excel = { run: async (fn: any) => fn({ workbook: wb, sync: async () => undefined }) };
  return { grids, state };
}

test("the preview names the style left on cells, rewritten references and formulas that break", async () => {
  workbook({ filtered: true });
  const plan = await prepareConvertTablePlan({ table: "ТПроба" });
  assert.equal(plan.resolvedAddress, "A1:B4");
  assert.match(plan.styleNote, /TableStyleMedium2.*останется на ячейках/);
  assert.deepEqual(plan.structured.map((item) => `${item.sheet}!${item.cell}`), ["Лист!B4", "Итоги!A1"]);
  assert.deepEqual(plan.names.map((item) => item.name), ["ИмяТ"]);
  assert.deepEqual(plan.literal.map((item) => item.cell), ["A2"]);
  assert.equal(plan.hiddenRows, 1);
  assert.ok(plan.warnings.some((warning) => /фильтр/.test(warning)));
  assert.ok(plan.warnings.some((warning) => /#ССЫЛКА!/.test(warning)));
  assert.equal(plan.undoAvailable, false);
});

test("conversion is verified: table gone, values kept, subtotal change under a filter is expected", async () => {
  workbook({ filtered: true });
  const result = await executeConvertTablePlan(await prepareConvertTablePlan({ table: "ТПроба" })) as any;
  assert.equal(result.executionState, "verified");
  assert.deepEqual(result.brokenFormulas, ["Итоги!A2"]);
  assert.match(result.filterNote, /B4: 100 → 300/);
  assert.equal(result.rewrittenReferences, 3);
});

test("a value that changed for no reason is reported as applied, not verified", async () => {
  workbook({ unexpected: true });
  await assert.rejects(
    async () => executeConvertTablePlan(await prepareConvertTablePlan({ table: "ТПроба" })),
    (error: any) => error.executionState === "applied" && /B2: 100 → 999/.test(error.message)
  );
});

test("a structured reference left behind is not reported as converted", async () => {
  workbook({ keepStructured: true });
  await assert.rejects(
    async () => executeConvertTablePlan(await prepareConvertTablePlan({ table: "ТПроба" })),
    (error: any) => error.executionState === "applied" && /Итоги!A1/.test(error.message)
  );
});

test("a table edited after the preview is left alone", async () => {
  const { grids } = workbook();
  const plan = await prepareConvertTablePlan({ table: "ТПроба" });
  grids.Лист.B2 = { formula: 150, value: 150 };
  await assert.rejects(() => executeConvertTablePlan(plan), (error: any) => error.executionState === "failed_before_write");
});
