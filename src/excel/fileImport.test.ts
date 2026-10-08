import test from "node:test";
import assert from "node:assert/strict";
import { convertTable, executeImportFilePlan, plainNumber, prepareImportFilePlan } from "./fileImportPlans";
import { PLANNED_TOOLS } from "./plans";
import { clear as clearUndo, setUndoMonitorReady, undoLast } from "./undo";

test("only unambiguous numbers are taken from text; the rest stays text for convert_values", () => {
  assert.equal(plainNumber("1200"), 1200);
  assert.equal(plainNumber("-12.5"), -12.5);
  assert.equal(plainNumber("12,50"), 12.5);
  assert.equal(plainNumber("0"), 0);
  assert.equal(plainNumber("007"), null, "код с ведущим нулём");
  assert.equal(plainNumber("1,500"), null, "полторы или тысяча пятьсот — не угадываем");
  assert.equal(plainNumber("1 200,50"), null);
  assert.equal(plainNumber("01.02.2026"), null);
  const out = convertTable("csv", [["Город", "Сумма", "Код"], ["Москва", "1 200,50", "007"], ["Казань", "800", "=1+1"]]);
  assert.deepEqual(out.values, [["Город", "Сумма", "Код"], ["Москва", "1 200,50", "007"], ["Казань", 800, "=1+1"]]);
  assert.equal(out.numbersFromText, 1);
  // «Код» — реквизит: текстом без предупреждения.
  assert.deepEqual(out.keptAsText, { count: 1, examples: ["1 200,50"] });
  // ИНН и счёт из CSV/TXT остаются текстом целиком («Книга11»: часть ИНН становилась числом).
  const bank = convertTable("txt", [["ИНН контрагента", "Счёт контрагента", "Сумма"], ["4385708398", "40702810000000001001", "1500"], ["0274051582", "", "20"]]);
  assert.deepEqual(bank.values.slice(1), [["4385708398", "40702810000000001001", 1500], ["0274051582", "", 20]]);
});

test("XLSX values keep their type, and dates keep their format", () => {
  const out = convertTable("xlsx", [["Дата", "Сумма"], [46037, 1200.5], ["007", null]], { "1,0": "dd.mm.yyyy" });
  assert.deepEqual(out.values, [["Дата", "Сумма"], [46037, 1200.5], ["007", ""]], "текст из Excel не превращается в число");
  assert.deepEqual(out.dateFormats, { "1,0": "dd.mm.yyyy" });
});

/** Лист, который ведёт себя как Excel: текст с апострофом — текст, «=…» без него — формула. */
function sheetMock(prefilled: Record<string, unknown> = {}) {
  const grid = new Map<string, unknown>(Object.entries(prefilled));
  const formats = new Map<string, string>();
  const letter = (index: number) => String.fromCharCode(65 + index);
  // Текст, записанный с апострофом, хранится текстом: { text }.
  const raw = (key: string) => { const value = grid.get(key) ?? ""; return value && typeof value === "object" ? (value as any).text : value; };
  const range = (row: number, column: number, rows: number, columns: number): any => ({
    get formulas() { return Array.from({ length: rows }, (_, r) => Array.from({ length: columns }, (_, c) => raw(`${letter(column + c)}${row + r + 1}`))); },
    get values() {
      return Array.from({ length: rows }, (_, r) => Array.from({ length: columns }, (_, c) => {
        const key = `${letter(column + c)}${row + r + 1}`;
        const value = grid.get(key) ?? "";
        if (value && typeof value === "object") return (value as any).text;
        return typeof value === "string" && value.startsWith("=") ? 2 : value;
      }));
    },
    set values(matrix: unknown[][]) {
      matrix.forEach((line, r) => line.forEach((value, c) => {
        const key = `${letter(column + c)}${row + r + 1}`;
        if (value === "") grid.delete(key);
        else grid.set(key, typeof value === "string" && value.startsWith("'") ? { text: value.slice(1) } : value);
      }));
    },
    getCell: (r: number, c: number) => ({ set numberFormat(value: string[][]) { formats.set(`${letter(column + c)}${row + r + 1}`, value[0][0]); } }),
    clear: () => { for (let r = 0; r < rows; r++) for (let c = 0; c < columns; c++) grid.delete(`${letter(column + c)}${row + r + 1}`); },
    load: () => undefined
  });
  const parse = (address: string) => {
    const [a, b = a] = address.split(":");
    const cell = (text: string) => ({ column: text.charCodeAt(0) - 65, row: Number(text.slice(1)) - 1 });
    return { start: cell(a), end: cell(b) };
  };
  const sheet: any = {
    id: "sheet-1", name: "Данные", load: () => undefined,
    protection: { protected: false, load: () => undefined },
    tables: { items: [], load: () => undefined },
    getRange: (address: string) => { const { start, end } = parse(address); return range(start.row, start.column, end.row - start.row + 1, end.column - start.column + 1); },
    getRangeByIndexes: range,
    getUsedRangeOrNullObject: () => ({ isNullObject: grid.size === 0, columnIndex: 0, columnCount: 3, rowIndex: 0, address: "Данные!A1:C3", load: () => undefined })
  };
  (globalThis as any).Office = { context: { document: { url: "C:/import.xlsx" }, requirements: { isSetSupported: () => true } } };
  (globalThis as any).Excel = {
    run: async (fn: any) => fn({
      workbook: { worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet, items: [sheet], load: () => undefined } },
      sync: async () => undefined
    })
  };
  return { grid, formats };
}

function serveTable(table: unknown) {
  (globalThis as any).fetch = async () => new Response(JSON.stringify(table));
}

const CSV_TABLE = {
  file: "sales.csv", kind: "csv", name: "sales.csv", rows: 3, columns: 3,
  cells: [["Город", "Сумма", "Ссылка"], ["Москва", "1200", '=HYPERLINK("http://evil")'], ["Казань", "1 200,50", ""]]
};

test("import_file_table goes through the plan registry", () => {
  assert.ok(PLANNED_TOOLS.includes("import_file_table"));
});

test("a CSV table lands in an empty place, a formula from the file stays text, and undo clears it", async (t) => {
  const previousFetch = globalThis.fetch;
  t.after(() => { globalThis.fetch = previousFetch; clearUndo(); setUndoMonitorReady(false); });
  const state = sheetMock();
  serveTable(CSV_TABLE);
  setUndoMonitorReady(true);
  const plan = await prepareImportFilePlan({ fileId: "f1", sheet: "Данные" });
  assert.equal(plan.destArea, "A1:C3");
  assert.equal(plan.numbersFromText, 1);
  assert.deepEqual(plan.keptAsText.examples, ["1 200,50"]);
  const result = await executeImportFilePlan(plan) as any;
  assert.equal(result.executionState, "verified");
  assert.equal(state.grid.get("B2"), 1200);
  assert.deepEqual(state.grid.get("C2"), { text: '=HYPERLINK("http://evil")' }, "записано с апострофом — текст");
  await undoLast();
  assert.equal(state.grid.size, 0);
});

test("an occupied place is refused before any card, with a free place named", async (t) => {
  const previousFetch = globalThis.fetch;
  t.after(() => { globalThis.fetch = previousFetch; });
  sheetMock({ B2: "занято" });
  serveTable(CSV_TABLE);
  await assert.rejects(() => prepareImportFilePlan({ fileId: "f1", sheet: "Данные" }), /занято: там 1 непустых ячеек.*destAddress: "E1"/);
});

test("rows can be taken in parts, and a too-large table is refused with how to split it", async (t) => {
  const previousFetch = globalThis.fetch;
  t.after(() => { globalThis.fetch = previousFetch; });
  sheetMock();
  serveTable(CSV_TABLE);
  const plan = await prepareImportFilePlan({ fileId: "f1", sheet: "Данные", firstRow: 2, lastRow: 2, destAddress: "D5" });
  assert.equal(plan.destArea, "D5:F5");
  assert.equal(plan.sourceRows, "2–2");
  serveTable({ ...CSV_TABLE, rows: 60_000, columns: 2, cells: Array.from({ length: 60_000 }, () => ["a", "b"]) });
  await assert.rejects(() => prepareImportFilePlan({ fileId: "f1", sheet: "Данные" }), /до 100000.*firstRow и lastRow/);
});
