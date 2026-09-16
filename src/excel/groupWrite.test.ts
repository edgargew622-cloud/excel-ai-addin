import test from "node:test";
import assert from "node:assert/strict";
import { executeSetRangesPlan, findOverlappingWrites, prepareSetRangesPlan, ToolError } from "./excelTools";

/** Макет одного листа: диапазоны адресуются как есть, запись фиксируется.
 * failAt — номер операции (с единицы), на которой Excel «ломается». */
function singleSheetExcel(options: { failAt?: number; failKind?: "before" | "after" } = {}) {
  const written: string[] = [];
  let writes = 0;
  const makeRange = (address: string) => {
    const range: any = {
      address: `Лист1!${address}`,
      rowCount: 1,
      columnCount: 1,
      rowIndex: 0,
      columnIndex: 0,
      load: () => undefined,
      values: [[0]],
      formulas: [[0]]
    };
    Object.defineProperty(range, "values", {
      get: () => range._v ?? [[0]],
      set: (value: unknown[][]) => {
        writes += 1;
        if (options.failAt === writes) throw new Error("Excel отказал на записи");
        range._v = value;
        written.push(address);
      }
    });
    return range;
  };
  const ranges = new Map<string, any>();
  const sheet: any = {
    id: "sheet-1",
    name: "Лист1",
    load: () => undefined,
    getRange: (address: string) => {
      if (!ranges.has(address)) ranges.set(address, makeRange(address));
      return ranges.get(address);
    }
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
  return { written };
}

const write = (address: string, value: unknown) => ({ sheet: "Лист1", address, values: [[value]] });

test("overlapping writes are refused before anything is written", async () => {
  const excel = singleSheetExcel();
  await assert.rejects(
    () => prepareSetRangesPlan({ writes: [write("A1", 1), write("B1", 2), write("A1", 3)] }),
    (error: unknown) => {
      assert.ok(error instanceof ToolError);
      // Названы обе конфликтующие операции, чтобы модель знала, что разделять.
      assert.match((error as Error).message, /Операции 1 и 3 пересекаются/);
      assert.match((error as Error).message, /отклонена до записи/);
      return true;
    }
  );
  assert.deepEqual(excel.written, []);
});

test("overlap is detected on resolved rectangles, not on address strings", () => {
  const plan = (a: string, b: string) => [
    { target: { sheetId: "s1", sheetName: "Лист1" }, resolvedAddress: a },
    { target: { sheetId: "s1", sheetName: "Лист1" }, resolvedAddress: b }
  ] as any;
  // Разные записи одних и тех же ячеек.
  assert.deepEqual(findOverlappingWrites(plan("A1:C3", "B2")), [0, 1]);
  // Соседние диапазоны не пересекаются.
  assert.equal(findOverlappingWrites(plan("A1:C3", "D1:F3")), null);
  // Одинаковые адреса на разных листах независимы.
  const otherSheet = [
    { target: { sheetId: "s1", sheetName: "Лист1" }, resolvedAddress: "A1" },
    { target: { sheetId: "s2", sheetName: "Лист2" }, resolvedAddress: "A1" }
  ] as any;
  assert.equal(findOverlappingWrites(otherSheet), null);
});

test("a failure stops the rest and separates done from not started", async () => {
  const excel = singleSheetExcel({ failAt: 2 });
  const plan = await prepareSetRangesPlan({ writes: [write("A1", 1), write("B1", 2), write("C1", 3)] });
  const result = await executeSetRangesPlan(plan) as any;

  assert.equal(result.ok, false);
  assert.equal(result.appliedCount, 1);
  // Первая выполнена и проверена. Вторая сорвалась на самой записи, и доказать,
  // что запись не началась, нельзя — поэтому «unknown», а не «не выполнено».
  assert.equal(result.operations[0].executionState, "verified");
  assert.equal(result.operations[1].executionState, "unknown");
  assert.equal(result.operations[2].executionState, "not_started");
  assert.match(result.operations[2].note, /группа остановлена/i);
  // Неопределённость одной операции делает неопределённым итог всей группы.
  assert.equal(result.executionState, "unknown");
  // Отчёт не обещает отката и требует перечитать, а не повторять.
  assert.match(result.note, /не откатываются/);
  assert.match(result.note, /перечитайте её диапазон/);
  assert.deepEqual(excel.written, ["A1"]);
});

test("a clean group reports every operation as verified", async () => {
  const excel = singleSheetExcel();
  const plan = await prepareSetRangesPlan({ writes: [write("A1", 1), write("B1", 2)] });
  const result = await executeSetRangesPlan(plan) as any;

  assert.equal(result.ok, true);
  assert.equal(result.executionState, "verified");
  assert.equal(result.appliedCount, 2);
  assert.deepEqual(excel.written, ["A1", "B1"]);
});
