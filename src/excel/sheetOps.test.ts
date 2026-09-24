import test from "node:test";
import assert from "node:assert/strict";
import { executeDeleteSheetPlan, executeRenameSheetPlan, prepareDeleteSheetPlan, prepareRenameSheetPlan } from "./sheetPlans";
import { PLANNED_TOOLS } from "./plans";
import { clear as clearUndo, setUndoMonitorReady, undoLast } from "./undo";

const letters = (index: number) => String.fromCharCode(64 + index);

/**
 * Книга, которая ведёт себя как Excel в замере 24 сентября 2026 года:
 * переименование переписывает обычные ссылки и именованные диапазоны,
 * но не имя внутри текста формулы; удаление превращает ссылки в #ССЫЛКА!.
 */
function book(options: { hidden?: string[]; unexpectedRef?: boolean } = {}) {
  const sheets: { id: string; name: string; cells: Map<string, unknown>; deleted?: boolean }[] = [
    { id: "s1", name: "Продажи", cells: new Map<string, unknown>([["A1", 10], ["A2", 20]]) },
    { id: "s2", name: "Отчёт", cells: new Map<string, unknown>([["A1", "=Продажи!A1"], ["A2", '=INDIRECT("Продажи!A1")'], ["A3", "=Ставка*2"], ["A4", "см. лист Продажи"], ["A5", "=A1+1"]]) },
    { id: "s3", name: "Прочее", cells: new Map<string, unknown>([["A1", 1]]) }
  ];
  const names: { name: string; formula: string }[] = [{ name: "Ставка", formula: "=Продажи!$A$2" }];
  const live = () => sheets.filter((sheet) => !sheet.deleted);
  const byName = (name: string) => live().find((sheet) => sheet.name.toLowerCase() === name.toLowerCase());

  /** Значение ячейки — как посчитал бы Excel. */
  const evaluate = (sheet: (typeof sheets)[number], cell: string): unknown => {
    const content = sheet.cells.get(cell) ?? "";
    if (typeof content !== "string" || !content.startsWith("=")) return content;
    if (content.includes("#REF!")) return "#ССЫЛКА!";
    const indirect = /INDIRECT\("(.+)!([A-Z]\d+)"\)/.exec(content);
    if (indirect) { const target = byName(indirect[1]); return target ? evaluate(target, indirect[2]) : "#ССЫЛКА!"; }
    const plain = /=(.+)!([A-Z]\d+)/.exec(content);
    if (plain) { const target = byName(plain[1].replace(/^'|'$/g, "")); return target ? evaluate(target, plain[2]) : "#ССЫЛКА!"; }
    if (content.includes("Ставка")) {
      const named = names.find((item) => item.name === "Ставка")!;
      if (named.formula.includes("#REF!")) return "#ССЫЛКА!";
      const [, sheetName, target] = /=(.+)!\$([A-Z])\$(\d+)/.exec(named.formula)!;
      const value = evaluate(byName(sheetName)!, `${target}${/\$(\d+)$/.exec(named.formula)![1]}`);
      return typeof value === "number" ? value * 2 : value;
    }
    return 11;
  };

  const worksheetObject = (sheet: (typeof sheets)[number]): any => ({
    get id() { return sheet.id; },
    get name() { return sheet.name; },
    set name(value: string) {
      const old = sheet.name;
      for (const other of sheets) for (const [cell, content] of other.cells) {
        if (typeof content === "string" && content.startsWith("=") && !content.includes("INDIRECT")) other.cells.set(cell, content.split(`${old}!`).join(`${value}!`));
      }
      for (const item of names) item.formula = item.formula.split(`${old}!`).join(`${value}!`);
      sheet.name = value;
    },
    get visibility() { return options.hidden?.includes(sheet.name) ? "Hidden" : "Visible"; },
    position: sheets.indexOf(sheet),
    isNullObject: false,
    load: () => undefined,
    delete: () => {
      sheet.deleted = true;
      for (const other of live()) for (const [cell, content] of other.cells) {
        if (typeof content === "string" && content.includes(`${sheet.name}!`) && !content.includes("INDIRECT")) other.cells.set(cell, content.split(`${sheet.name}!`).join("#REF!"));
      }
      for (const item of names) item.formula = item.formula.split(`${sheet.name}!`).join("#REF!");
      if (options.unexpectedRef) byName("Прочее")!.cells.set("B1", "=#REF!A1");
    },
    getUsedRangeOrNullObject: () => {
      const keys = [...sheet.cells.keys()].filter((key) => sheet.cells.get(key) !== "");
      if (!keys.length) return { isNullObject: true, load: () => undefined };
      const rows = keys.map((key) => Number(key.slice(1)));
      const cols = keys.map((key) => key.charCodeAt(0) - 64);
      const [r0, r1, c0, c1] = [Math.min(...rows), Math.max(...rows), Math.min(...cols), Math.max(...cols)];
      const grid = (read: (cell: string) => unknown) => Array.from({ length: r1 - r0 + 1 }, (_, r) => Array.from({ length: c1 - c0 + 1 }, (_, c) => read(`${letters(c0 + c)}${r0 + r}`)));
      return {
        isNullObject: false, load: () => undefined,
        address: `${sheet.name}!${letters(c0)}${r0}:${letters(c1)}${r1}`,
        rowIndex: r0 - 1, columnIndex: c0 - 1, rowCount: r1 - r0 + 1, columnCount: c1 - c0 + 1,
        get formulas() { return grid((cell) => sheet.cells.get(cell) ?? ""); },
        get values() { return grid((cell) => evaluate(sheet, cell)); }
      };
    },
    getRange: (cell: string) => ({ load: () => undefined, get values() { return [[evaluate(sheet, cell)]]; } }),
    charts: { items: [], load: () => undefined },
    tables: { items: [], load: () => undefined },
    pivotTables: { items: [], load: () => undefined }
  });

  const worksheets: any = {
    get items() { return live().map(worksheetObject); },
    load: () => undefined,
    getItem: (key: string) => worksheetObject(live().find((sheet) => sheet.id === key || sheet.name === key)!),
    getItemOrNullObject: (key: string) => {
      const found = live().find((sheet) => sheet.id === key || sheet.name === key);
      return found ? worksheetObject(found) : { isNullObject: true, load: () => undefined };
    }
  };
  (globalThis as any).Office = { context: { document: { url: "C:/sheets.xlsx" }, requirements: { isSetSupported: () => true } } };
  (globalThis as any).Excel = {
    run: async (fn: any) => fn({
      workbook: {
        worksheets,
        names: { get items() { return names; }, load: () => undefined },
        protection: { protected: false, load: () => undefined }
      },
      sync: async () => undefined
    })
  };
  return { sheets, names, cell: (sheet: string, cell: string) => byName(sheet)?.cells.get(cell) };
}

test("renaming and deleting a sheet go through the plan registry", () => {
  assert.ok(PLANNED_TOOLS.includes("rename_sheet"));
  assert.ok(PLANNED_TOOLS.includes("delete_sheet"));
});

test("a rename names what Excel will not rewrite, and the referencing values are checked", async () => {
  const state = book();
  const plan = await prepareRenameSheetPlan({ sheet: "Продажи", newName: "Продажи 2026" });
  assert.deepEqual(plan.referencing.map((item) => `${item.sheet}!${item.cell}`), ["Отчёт!A1"]);
  assert.deepEqual(plan.literal.map((item) => item.cell), ["A2"], "INDIRECT с прежним именем");
  assert.deepEqual(plan.textMentions.map((item) => item.cell), ["A4"]);
  const result = await executeRenameSheetPlan(plan) as any;
  assert.equal(result.executionState, "verified");
  assert.equal(state.cell("Отчёт", "A1"), "=Продажи 2026!A1", "обычную ссылку Excel переписал");
  assert.equal(result.brokenLiteralFormulas.length, 1);
  assert.equal(state.names[0].formula, "=Продажи 2026!$A$2");
});

test("a rename is undone back to the old name", async () => {
  const state = book();
  setUndoMonitorReady(true);
  try {
    await executeRenameSheetPlan(await prepareRenameSheetPlan({ sheet: "Продажи", newName: "Выручка" }));
    await undoLast();
    assert.ok(state.sheets.some((sheet) => sheet.name === "Продажи"));
  } finally {
    clearUndo();
    setUndoMonitorReady(false);
  }
});

test("a taken name, the same name, and a rename done meanwhile are refused", async () => {
  const state = book();
  await assert.rejects(() => prepareRenameSheetPlan({ sheet: "Продажи", newName: "отчёт" }), /уже есть.*Свободно, например, «отчёт 2»/);
  await assert.rejects(() => prepareRenameSheetPlan({ sheet: "Продажи", newName: "Продажи" }), /менять нечего/);
  await assert.rejects(() => prepareRenameSheetPlan({ sheet: "Нет", newName: "Да" }), /Листа «Нет» нет/);
  const plan = await prepareRenameSheetPlan({ sheet: "Продажи", newName: "Выручка" });
  state.sheets[0].name = "Вручную";
  await assert.rejects(() => executeRenameSheetPlan(plan), (error: any) => error.executionState === "failed_before_write");
});

test("a delete lists what disappears and what breaks, and the new errors match the prediction", async () => {
  const state = book();
  const plan = await prepareDeleteSheetPlan({ sheet: "Продажи" });
  assert.equal(plan.filledCells, 2);
  assert.deepEqual(plan.brokenNames, ["Ставка"]);
  assert.deepEqual([...plan.referencing, ...plan.literal, ...plan.viaNames].map((item) => item.cell).sort(), ["A1", "A2", "A3"]);
  assert.equal(plan.undoAvailable, false);
  const result = await executeDeleteSheetPlan(plan) as any;
  assert.equal(result.executionState, "verified");
  assert.equal(result.refErrorsAfter - result.refErrorsBefore, 3);
  assert.ok(!state.sheets.some((sheet) => sheet.name === "Продажи" && !sheet.deleted));
});

test("an unforeseen broken reference after a delete is not passed over", async () => {
  book({ unexpectedRef: true });
  const plan = await prepareDeleteSheetPlan({ sheet: "Продажи" });
  await assert.rejects(() => executeDeleteSheetPlan(plan), (error: any) => {
    assert.equal(error.executionState, "applied");
    assert.match(error.message, /новых ошибок ссылок в книге 4, а предсказано 3/);
    return true;
  });
});

test("the last visible sheet is not deleted, and an edit after the preview stops the delete", async () => {
  book({ hidden: ["Отчёт", "Прочее"] });
  await assert.rejects(() => prepareDeleteSheetPlan({ sheet: "Продажи" }), /последний видимый лист/);
  const state = book();
  const plan = await prepareDeleteSheetPlan({ sheet: "Прочее" });
  state.sheets[2].cells.set("B2", "новое");
  await assert.rejects(() => executeDeleteSheetPlan(plan), (error: any) => error.executionState === "failed_before_write");
  assert.ok(state.sheets[2].deleted !== true, "лист не удалён");
});
