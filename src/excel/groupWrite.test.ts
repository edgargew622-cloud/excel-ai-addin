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

/** Макет объединения: запись в неугловую ячейку Excel принимает молча,
 * но значение никуда не попадает — так ведёт себя настоящий Excel. */
function mergedCellExcel() {
  const range: any = {
    address: "Данные!M1",
    rowCount: 1, columnCount: 1, rowIndex: 0, columnIndex: 12,
    load: () => undefined,
    values: [[""]],
    formulas: [[""]]
  };
  Object.defineProperty(range, "values", {
    get: () => [[""]],
    set: () => { /* Excel молча игнорирует запись в неугловую ячейку */ }
  });
  const sheet: any = {
    id: "sheet-1", name: "Данные", load: () => undefined,
    getRange: () => range
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
}

test("a write that changes nothing says so instead of blaming the read-back", async () => {
  mergedCellExcel();
  const { executeSetRangePlan, prepareSetRangePlan } = await import("./excelTools");
  const plan = await prepareSetRangePlan({ sheet: "Данные", address: "M1", values: [[5]] });

  await assert.rejects(() => executeSetRangePlan(plan), (error: any) => {
    // Повторять бессмысленно, и агенту нужно сказать именно это.
    assert.match(error.message, /не дала эффекта/);
    assert.match(error.message, /объединённой области/);
    assert.match(error.message, /Повтор ничего не изменит/);
    // Доказано лишь то, что цель не изменилась, а не вся книга.
    assert.equal(error.executionState, "applied");
    return true;
  });
});

/** Макет защищённого листа: Excel сорвал бы запись, но мы отказываем раньше. */
function protectedSheetExcel(options: { protectedSheet: boolean; locked: boolean | null }) {
  let writes = 0;
  const range: any = {
    address: "Защищённый!A2",
    rowCount: 1, columnCount: 1, rowIndex: 1, columnIndex: 0,
    load: () => undefined,
    formulas: [[""]],
    format: { protection: { locked: options.locked, load: () => undefined } }
  };
  Object.defineProperty(range, "values", { get: () => [[""]], set: () => { writes += 1; } });
  const sheet: any = {
    id: "sheet-p", name: "Защищённый", load: () => undefined,
    protection: { protected: options.protectedSheet, load: () => undefined },
    getRange: () => range
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
  return { writeCount: () => writes };
}

test("a protected target is refused while nothing has been written yet", async () => {
  const excel = protectedSheetExcel({ protectedSheet: true, locked: true });
  const { prepareSetRangePlan } = await import("./excelTools");

  await assert.rejects(
    () => prepareSetRangePlan({ sheet: "Защищённый", address: "A2", values: [[22]] }),
    (error: any) => {
      // Провал до записи доказуем, в отличие от сорвавшейся попытки.
      assert.match(error.message, /защищён/);
      assert.match(error.message, /не выполнялась/);
      return true;
    }
  );
  assert.equal(excel.writeCount(), 0);
});

test("mixed locking on a protected sheet is refused too: proof is impossible", async () => {
  protectedSheetExcel({ protectedSheet: true, locked: null });
  const { prepareSetRangePlan } = await import("./excelTools");
  await assert.rejects(
    () => prepareSetRangePlan({ sheet: "Защищённый", address: "A2", values: [[22]] }),
    /заблокированы не все одинаково/
  );
});

test("unlocked cells on a protected sheet are still writable", async () => {
  // Защита листа не запрещает запись в явно разблокированные ячейки.
  protectedSheetExcel({ protectedSheet: true, locked: false });
  const { prepareSetRangePlan } = await import("./excelTools");
  const plan = await prepareSetRangePlan({ sheet: "Защищённый", address: "A2", values: [[22]] });
  assert.equal(plan.cellCount, 1);
});
