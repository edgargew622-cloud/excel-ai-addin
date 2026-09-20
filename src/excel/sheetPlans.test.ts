import test from "node:test";
import assert from "node:assert/strict";
import { checkSheetName, freeSheetName } from "./sheetRules";
import { executeCreateSheetPlan, prepareCreateSheetPlan } from "./sheetPlans";
import { PLANNED_TOOLS } from "./plans";
import { clear as clearUndo, setUndoMonitorReady, undoLast } from "./undo";

/* --- имя листа -------------------------------------------------------------- */

test("a sheet name is checked by Excel's own rules before the workbook is touched", () => {
  assert.equal(checkSheetName("  Сводка  ", ["Данные"]), "Сводка", "лишние пробелы по краям убираются");
  assert.throws(() => checkSheetName("", ["Данные"]), /не может быть пустым/);
  assert.throws(() => checkSheetName("Отчёт: январь", []), /нельзя использовать/);
  assert.throws(() => checkSheetName("a".repeat(32), []), /не длиннее 31/);
  assert.throws(() => checkSheetName("'Сводка", []), /апостроф/);
  assert.throws(() => checkSheetName("История", []), /служебное имя/);
  // Имена листов в Excel не различаются регистром.
  assert.throws(() => checkSheetName("данные", ["Данные"]), /уже есть/);
});

test("a free name is suggested instead of a bare refusal", () => {
  assert.equal(freeSheetName("Сводка", ["Сводка"]), "Сводка 2");
  assert.equal(freeSheetName("Сводка", ["Сводка", "Сводка 2"]), "Сводка 3");
});

/* --- полный путь ------------------------------------------------------------ */

function workbook(options: { renamesTo?: string; refuses?: boolean; filledAfter?: boolean } = {}) {
  const sheets: any[] = [];
  const add = (name: string, position: number) => {
    const sheet: any = {
      id: `sheet-${position + 1}`,
      name,
      position,
      isNullObject: false,
      filled: false,
      load: () => undefined,
      getUsedRangeOrNullObject: () => ({
        isNullObject: !sheet.filled,
        address: `${sheet.name}!A1:B2`,
        load: () => undefined
      }),
      delete: () => { sheets.splice(sheets.indexOf(sheet), 1); }
    };
    sheets.push(sheet);
    return sheet;
  };
  add("Данные", 0);
  add("Отчёт", 1);

  (globalThis as any).Office = { context: { document: { url: "C:/книга.xlsx" }, requirements: { isSetSupported: () => true } } };
  (globalThis as any).Excel = {
    run: async (fn: any) => fn({
      workbook: {
        worksheets: {
          get items() { return sheets; },
          load: () => undefined,
          add: (name: string) => {
            if (options.refuses) throw new Error("Во время обработки запроса произошла внутренняя ошибка.");
            const sheet = add(options.renamesTo ?? name, sheets.length);
            sheet.filled = options.filledAfter === true;
            return sheet;
          },
          getItemOrNullObject: (id: string) => sheets.find((item) => item.id === id) ?? { isNullObject: true, load: () => undefined }
        }
      },
      sync: async () => undefined
    })
  };
  return { sheets };
}

test("create_sheet goes through the plan registry", () => {
  assert.ok(PLANNED_TOOLS.includes("create_sheet"));
});

test("a new sheet is created last and verified by its name", async () => {
  const state = workbook();
  const plan = await prepareCreateSheetPlan({ name: "Сводка" });
  assert.equal(plan.positionText, "последним в книге");
  assert.deepEqual(plan.sheetsBefore, ["Данные", "Отчёт"]);

  const result = await executeCreateSheetPlan(plan) as any;
  assert.equal(result.executionState, "verified");
  assert.equal(result.sheet, "Сводка");
  assert.deepEqual(result.sheetsAfter, ["Данные", "Отчёт", "Сводка"]);
  assert.equal(state.sheets.length, 3);
});

test("a sheet can be placed right after a named one", async () => {
  workbook();
  const plan = await prepareCreateSheetPlan({ name: "Сводка", after: "Данные" });
  assert.match(plan.positionText, /после листа «Данные»/);
  const result = await executeCreateSheetPlan(plan) as any;
  assert.equal(result.placedAfter, "Данные");
  assert.equal(result.position, 2, "сразу за первым листом");
});

test("an existing name is refused with a free one named", async () => {
  workbook();
  await assert.rejects(() => prepareCreateSheetPlan({ name: "Отчёт" }), (error: any) => {
    assert.match(error.message, /Лист «Отчёт» в книге уже есть/);
    assert.match(error.message, /Свободно, например, «Отчёт 2»/);
    return true;
  });
});

test("a missing anchor sheet is refused with the real sheet names", async () => {
  workbook();
  await assert.rejects(
    () => prepareCreateSheetPlan({ name: "Сводка", after: "Итоги" }),
    /Листа «Итоги» нет.*«Данные», «Отчёт»/s
  );
});

test("a name taken between preview and confirmation stops the operation", async () => {
  const state = workbook();
  const plan = await prepareCreateSheetPlan({ name: "Сводка" });
  state.sheets.push({ id: "manual", name: "Сводка", position: 2, load: () => undefined });

  await assert.rejects(() => executeCreateSheetPlan(plan), (error: any) => {
    assert.equal(error.executionState, "failed_before_write");
    assert.match(error.message, /появился в книге после предпросмотра/);
    return true;
  });
});

test("a refusal from Excel leaves the outcome unknown, not successful", async () => {
  workbook({ refuses: true });
  const plan = await prepareCreateSheetPlan({ name: "Сводка" });
  await assert.rejects(() => executeCreateSheetPlan(plan), (error: any) => {
    assert.equal(error.executionState, "unknown");
    return true;
  });
});

test("a name Excel changed on its own is reported as a mismatch", async () => {
  workbook({ renamesTo: "Сводка1" });
  const plan = await prepareCreateSheetPlan({ name: "Сводка" });
  await assert.rejects(() => executeCreateSheetPlan(plan), (error: any) => {
    assert.equal(error.executionState, "applied");
    assert.match(error.message, /назвал его «Сводка1»/);
    return true;
  });
});

test("undo removes the sheet while it is empty", async () => {
  const state = workbook();
  setUndoMonitorReady(true);
  try {
    const plan = await prepareCreateSheetPlan({ name: "Сводка" });
    const result = await executeCreateSheetPlan(plan) as any;
    assert.equal(result.undoable, true);
    assert.match(await undoLast(), /создание листа «Сводка»/);
    assert.deepEqual(state.sheets.map((sheet) => sheet.name), ["Данные", "Отчёт"]);
  } finally {
    clearUndo();
    setUndoMonitorReady(false);
  }
});

test("undo refuses to delete a sheet somebody has already used", async () => {
  const state = workbook({ filledAfter: true });
  setUndoMonitorReady(true);
  try {
    const plan = await prepareCreateSheetPlan({ name: "Сводка" });
    await executeCreateSheetPlan(plan);
    // Удаление листа с данными унесло бы чужую работу.
    await assert.rejects(() => undoLast(), /уже есть данные/);
    assert.equal(state.sheets.length, 3, "лист остался");
  } finally {
    clearUndo();
    setUndoMonitorReady(false);
  }
});
