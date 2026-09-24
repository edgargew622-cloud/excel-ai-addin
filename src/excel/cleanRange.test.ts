import test from "node:test";
import assert from "node:assert/strict";
import { executeCleanPlan, executeRemoveDuplicatesPlan, prepareConvertValuesPlan, prepareRemoveDuplicatesPlan, prepareTrimTextPlan } from "./dataCleaning";
import { PLANNED_TOOLS } from "./plans";
import { clear as clearUndo, setUndoMonitorReady, undoLast } from "./undo";

const NBSP = String.fromCharCode(160);
const letters = (index: number) => String.fromCharCode(64 + index);

/**
 * Лист, который ведёт себя как Excel ru-RU из замера 24 сентября 2026 года:
 * текст с апострофом хранится текстом; текст без апострофа, похожий
 * на число, Excel превращает в число — поэтому очистка обязана писать
 * текст с апострофом. Дата — число, её вид задаёт формат.
 */
function cleanSheet(
  cells: Record<string, unknown>,
  options: {
    ignore?: string[];
    date1904?: boolean;
    /** Как removeDuplicates сравнивает текст: как Excel (без регистра) или с регистром. */
    duplicatesCaseSensitive?: boolean;
    tables?: { name: string; address: string }[];
  } = {}
) {
  const grid = new Map<string, unknown>();
  const formats = new Map<string, string>();
  const key = (row: number, column: number) => `${letters(column)}${row}`;
  const excelDate = (serial: number) => {
    const shifted = options.date1904 ? serial + 1462 : serial;
    const date = new Date(Date.UTC(1899, 11, 30) + shifted * 86_400_000);
    return `${String(date.getUTCDate()).padStart(2, "0")}.${String(date.getUTCMonth() + 1).padStart(2, "0")}.${date.getUTCFullYear()}`;
  };
  const input = (value: unknown) => {
    if (typeof value !== "string") return value;
    if (value.startsWith("'")) return { text: value.slice(1) };
    return /^-?\d+(,\d+)?$/.test(value) ? Number(value.replace(",", ".")) : value;
  };
  // Начальное содержимое вводится так же, как его ввёл бы человек или запись.
  for (const [name, value] of Object.entries(cells)) grid.set(name, input(value));
  const stored = (name: string) => grid.get(name) ?? "";
  const valueOf = (name: string) => {
    const raw = stored(name);
    return raw && typeof raw === "object" && "text" in (raw as any) ? (raw as any).text : raw;
  };
  const isFormula = (name: string) => typeof stored(name) === "string" && String(stored(name)).startsWith("=");

  function makeRange(address: string): any {
    const [start, end = start] = address.replace(/^.*!/, "").split(":");
    const parse = (text: string) => ({ column: text.charCodeAt(0) - 64, row: Number(text.slice(1)) });
    const from = parse(start);
    const to = parse(end);
    const matrix = (read: (name: string) => unknown) =>
      Array.from({ length: to.row - from.row + 1 }, (_, r) => Array.from({ length: to.column - from.column + 1 }, (_, c) => read(key(from.row + r, from.column + c))));
    const write = (source: unknown[][], convert: (value: unknown) => unknown) => source.forEach((row, r) => row.forEach((value, c) => {
      const name = key(from.row + r, from.column + c);
      if (!options.ignore?.includes(name)) grid.set(name, convert(value));
    }));
    return {
      address: `Данные!${address}`,
      rowIndex: from.row - 1,
      columnIndex: from.column - 1,
      rowCount: to.row - from.row + 1,
      columnCount: to.column - from.column + 1,
      load: () => undefined,
      format: { protection: { locked: false, load: () => undefined } },
      get formulas() { return matrix((name) => (isFormula(name) ? stored(name) : valueOf(name))); },
      set formulas(source: unknown[][]) { write(source, input); },
      get values() { return matrix((name) => (isFormula(name) ? 5 : valueOf(name))); },
      set values(source: unknown[][]) { write(source, input); },
      get valueTypes() {
        return matrix((name) => {
          const value = isFormula(name) ? 5 : valueOf(name);
          return value === "" ? "Empty" : typeof value === "number" ? "Double" : "String";
        });
      },
      get numberFormat() { return matrix((name) => formats.get(name) ?? "General"); },
      set numberFormat(source: unknown[][]) {
        source.forEach((row, r) => row.forEach((value, c) => formats.set(key(from.row + r, from.column + c), String(value))));
      },
      get text() {
        return matrix((name) => {
          const value = valueOf(name);
          return typeof value === "number" && /d/.test(formats.get(name) ?? "") ? excelDate(value) : String(value);
        });
      },
      getCell(r: number, c: number) { return makeRange(key(from.row + r, from.column + c)); },
      // Так удалял дубликаты Excel в замере: первое вхождение остаётся,
      // значения ниже поднимаются внутри области, соседние столбцы стоят.
      removeDuplicates(columns: number[], includesHeader: boolean) {
        const rows = this.values as unknown[][];
        const body = rows.slice(includesHeader ? 1 : 0);
        const seen = new Set<string>();
        const kept = body.filter((row) => {
          const id = JSON.stringify(columns.map((column) => {
            const value = row[column];
            return typeof value === "string" && !options.duplicatesCaseSensitive ? `s:${value.toLowerCase()}` : `${typeof value}:${value}`;
          }));
          if (seen.has(id)) return false;
          seen.add(id);
          return true;
        });
        const removed = body.length - kept.length;
        const width = rows[0].length;
        const filled = [...kept, ...Array.from({ length: removed }, () => Array.from({ length: width }, () => ""))];
        filled.forEach((row, r) => row.forEach((value, c) => grid.set(key(from.row + (includesHeader ? 1 : 0) + r, from.column + c), value)));
        return { removed, uniqueRemaining: kept.length, load: () => undefined };
      }
    };
  }
  const sheet: any = {
    id: "sheet-1",
    name: "Данные",
    load: () => undefined,
    protection: { protected: false, load: () => undefined },
    getRange: (address: string) => makeRange(address),
    getRangeByIndexes: (row: number, column: number, rows: number, columns: number) =>
      makeRange(`${letters(column + 1)}${row + 1}:${letters(column + columns)}${row + rows}`),
    tables: { items: (options.tables ?? []).map((table) => ({ name: table.name, getRange: () => ({ address: `Данные!${table.address}`, load: () => undefined }) })), load: () => undefined },
    // Занятая область — для обхода формул книги.
    getUsedRangeOrNullObject: () => {
      const names = [...grid.keys()].filter((name) => stored(name) !== "");
      if (!names.length) return { isNullObject: true, load: () => undefined };
      const rows = names.map((name) => Number(name.slice(1)));
      const columns = names.map((name) => name.charCodeAt(0) - 64);
      const address = `${letters(Math.min(...columns))}${Math.min(...rows)}:${letters(Math.max(...columns))}${Math.max(...rows)}`;
      return Object.assign(makeRange(address), { isNullObject: false });
    }
  };
  (globalThis as any).Office = { context: { document: { url: "C:/clean.xlsx" }, requirements: { isSetSupported: () => true } } };
  (globalThis as any).Excel = {
    run: async (fn: any) => fn({
      workbook: {
        worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet, items: [sheet], load: () => undefined },
        application: {
          cultureInfo: {
            name: "ru-RU",
            load: () => undefined,
            numberFormat: { numberDecimalSeparator: ",", numberGroupSeparator: " ", load: () => undefined },
            datetimeFormat: { dateSeparator: ".", shortDatePattern: "ДД.ММ.ГГГГ", load: () => undefined }
          }
        }
      },
      sync: async () => undefined
    })
  };
  return { grid, formats, valueOf };
}

test("both cleaning tools go through the plan registry", () => {
  assert.ok(PLANNED_TOOLS.includes("trim_text"));
  assert.ok(PLANNED_TOOLS.includes("convert_values"));
});

test("extra spaces are removed, and a code stays text with its zeros", async () => {
  const sheet = cleanSheet({ A1: "' Москва ", A2: `'Омск${NBSP}`, A3: "'Казань   Север", A4: "' 007 ", A5: 42, A6: "=A5*2", A7: "'Тула" });
  const plan = await prepareTrimTextPlan({ sheet: "Данные", address: "A1:A7" });
  assert.equal(plan.changes.length, 4, "Тула, число и формула не меняются");
  assert.deepEqual(plan.sample.slice(0, 2), ["A1: « Москва » → «Москва»", `A2: «Омск${NBSP}» → «Омск»`]);
  assert.ok(Object.keys(plan.skipped).some((reason) => /формула/.test(reason)));
  const result = await executeCleanPlan(plan) as any;
  assert.equal(result.executionState, "verified");
  assert.equal(sheet.valueOf("A3"), "Казань Север");
  // Без апострофа Excel сделал бы из «007» число 7.
  assert.equal(sheet.valueOf("A4"), "007");
  assert.equal(sheet.valueOf("A5"), 42);
});

test("numbers are made from text only where the separators are certain", async () => {
  const sheet = cleanSheet({ B1: "'1 200", B2: "'1,500", B3: "'1.5", B4: "'007", B5: "'2 300,50", B6: "'Москва" });
  const plan = await prepareConvertValuesPlan({ sheet: "Данные", address: "B1:B6", to: "number" });
  assert.deepEqual(plan.changes.map((item) => `${item.cell}=${item.after}`), ["B1=1200", "B5=2300.5"]);
  const reasons = Object.keys(plan.skipped).join(" | ");
  assert.match(reasons, /дробь или тысячи/);
  assert.match(reasons, /не как в книге/);
  assert.match(reasons, /ведущими нулями/);
  await executeCleanPlan(plan);
  assert.equal(sheet.valueOf("B1"), 1200);
  assert.equal(sheet.valueOf("B2"), "1,500", "неоднозначное не тронуто");
});

test("once the user names the decimal separator, the ambiguous numbers are converted", async () => {
  const sheet = cleanSheet({ B1: "'1,500", B2: "'1.5", B3: "'007" });
  await executeCleanPlan(await prepareConvertValuesPlan({ sheet: "Данные", address: "B1:B3", to: "number", decimalSeparator: "." }));
  assert.equal(sheet.valueOf("B1"), 1500);
  assert.equal(sheet.valueOf("B2"), 1.5);
  assert.equal(sheet.valueOf("B3"), "007", "код не становится числом и по прямой просьбе");
});

test("dates get their serial and the workbook's date format; ambiguous ones wait for the order", async () => {
  const sheet = cleanSheet({ C1: "'25.02.2026", C2: "'01.02.2026", C3: "'2026-03-05" });
  const plan = await prepareConvertValuesPlan({ sheet: "Данные", address: "C1:C3", to: "date" });
  assert.equal(plan.numberFormat, "dd.mm.yyyy");
  assert.deepEqual(plan.changes.map((item) => item.cell), ["C1", "C3"]);
  assert.match(Object.keys(plan.skipped)[0], /день и месяц не различить/);
  await executeCleanPlan(plan);
  assert.equal(sheet.valueOf("C1"), 46078);
  assert.equal(sheet.formats.get("C1"), "dd.mm.yyyy");

  await executeCleanPlan(await prepareConvertValuesPlan({ sheet: "Данные", address: "C2", to: "date", dateOrder: "DMY" }));
  assert.equal(sheet.valueOf("C2"), 46054, "1 февраля 2026");
});

test("a workbook counting dates from 1904 is caught by how the date looks", async () => {
  cleanSheet({ C1: "'25.02.2026" }, { date1904: true });
  const plan = await prepareConvertValuesPlan({ sheet: "Данные", address: "C1", to: "date" });
  await assert.rejects(() => executeCleanPlan(plan), (error: any) => {
    assert.equal(error.executionState, "applied");
    assert.match(error.message, /C1.*25\.02\.2026/);
    return true;
  });
});

test("data edited after the preview stops the cleaning, and a cell Excel ignored is named", async () => {
  const edited = cleanSheet({ A1: "' Москва " });
  const plan = await prepareTrimTextPlan({ sheet: "Данные", address: "A1" });
  edited.grid.set("A1", { text: "Москва!" });
  await assert.rejects(() => executeCleanPlan(plan), (error: any) => error.executionState === "failed_before_write");

  cleanSheet({ A1: "' Москва ", A2: "' Омск " }, { ignore: ["A2"] });
  await assert.rejects(async () => executeCleanPlan(await prepareTrimTextPlan({ sheet: "Данные", address: "A1:A2" })), (error: any) => {
    assert.equal(error.executionState, "applied");
    assert.match(error.message, /A2/);
    return true;
  });
});

test("undo brings back the text and, for dates, the old number format", async () => {
  const sheet = cleanSheet({ C1: "'25.02.2026", C2: "'1 200" });
  setUndoMonitorReady(true);
  try {
    await executeCleanPlan(await prepareConvertValuesPlan({ sheet: "Данные", address: "C1", to: "date" }));
    assert.equal(sheet.formats.get("C1"), "dd.mm.yyyy");
    await undoLast();
    assert.equal(sheet.valueOf("C1"), "25.02.2026", "вернулся текст, а не число");
    assert.equal(sheet.formats.get("C1"), "General", "и прежний формат");
  } finally {
    clearUndo();
    setUndoMonitorReady(false);
  }
});

test("nothing to change is refused before any card", async () => {
  cleanSheet({ A1: "'Москва", A2: 5 });
  await assert.rejects(() => prepareTrimTextPlan({ sheet: "Данные", address: "A1:A2" }), /менять нечего/);
});

/* --- дубликаты (7.2.4) --------------------------------------------------------------- */

/** Таблица из замера removeDuplicates 24 сентября 2026 года, со шапкой. */
const ORDERS: Record<string, unknown> = {
  A1: "'Город", B1: "'Сумма", C1: "'Номер",
  A2: "'Москва", B2: 100, C2: 1,
  A3: "'москва", B3: 100, C3: 2,
  A4: "'Москва ", B4: 100, C4: 3,
  A5: "'Омск", B5: 200, C5: 4,
  A6: "'Омск", B6: 200, C6: 5,
  A7: "'Казань", B7: 300, C7: 6
};

test("duplicates are removed as Excel removes them, and the result is checked row by row", async () => {
  const sheet = cleanSheet({ ...ORDERS, E2: "=C3", E3: "=SUM(C2:C7)" });
  const plan = await prepareRemoveDuplicatesPlan({ sheet: "Данные", address: "A1:C7", columns: ["Город", "Сумма"] });
  assert.deepEqual(plan.removed, [{ row: 3, duplicateOf: 2 }, { row: 6, duplicateOf: 5 }]);
  assert.deepEqual(plan.keyNames, ["«Город»", "«Сумма»"]);
  // =C3 после сдвига покажет другую строку; итог по всей высоте — нет.
  assert.deepEqual(plan.risks.map((risk) => `${risk.cell} ${risk.formula}`), ["E2 =C3"]);
  assert.equal(plan.undoAvailable, false);
  const result = await executeRemoveDuplicatesPlan(plan) as any;
  assert.equal(result.executionState, "verified");
  assert.equal(result.removedRows, 2);
  assert.deepEqual([2, 3, 4, 5, 6, 7].map((row) => sheet.valueOf(`C${row}`)), [1, 3, 4, 6, "", ""]);
});

test("an Excel that compared text differently is caught, not reported as done", async () => {
  cleanSheet(ORDERS, { duplicatesCaseSensitive: true });
  const plan = await prepareRemoveDuplicatesPlan({ sheet: "Данные", address: "A1:C7", columns: ["Город", "Сумма"] });
  await assert.rejects(() => executeRemoveDuplicatesPlan(plan), (error: any) => {
    assert.equal(error.executionState, "applied");
    assert.match(error.message, /удалил дубликатов: 1, а по расчёту панели — 2/);
    return true;
  });
});

test("rows would drift apart from neighbouring data, so the operation refuses", async () => {
  cleanSheet({ ...ORDERS, D2: "'заметка" });
  await assert.rejects(() => prepareRemoveDuplicatesPlan({ sheet: "Данные", address: "A1:C7" }), /в столбцах D.*разъехались/);
});

test("formulas inside the area, a table, an unknown key column are refused before any card", async () => {
  cleanSheet({ ...ORDERS, C7: "=C6+1" });
  await assert.rejects(() => prepareRemoveDuplicatesPlan({ sheet: "Данные", address: "A1:C7" }), /есть формулы/);
  cleanSheet(ORDERS, { tables: [{ name: "Заказы", address: "A1:C7" }] });
  await assert.rejects(() => prepareRemoveDuplicatesPlan({ sheet: "Данные", address: "A1:C7" }), /таблицу Excel «Заказы»/);
  cleanSheet(ORDERS);
  await assert.rejects(() => prepareRemoveDuplicatesPlan({ sheet: "Данные", address: "A1:C7", columns: ["Регион"] }), /Столбца «Регион».*Заголовки: «Город»/);
});

test("no duplicates is said plainly, with a hint when spaces are what keeps rows apart", async () => {
  cleanSheet({ A1: "'Город", A2: "'Москва", A3: "'Москва ", A4: "'Омск" });
  await assert.rejects(() => prepareRemoveDuplicatesPlan({ sheet: "Данные", address: "A1:A4" }), /дубликатов по ключу нет.*trim_text/);
});

/* --- результат на отдельный лист (7.2.5) ------------------------------------------- */

function withSecondSheet(state: ReturnType<typeof cleanSheet>, filled = false) {
  // Второй лист — отдельная сетка ячеек; getItemOrNullObject находит его по имени.
  const other = new Map<string, unknown>(filled ? [["A1", "занято"]] : []);
  const range = (address: string): any => {
    const [start, end = start] = address.split(":");
    const from = { column: start.charCodeAt(0) - 64, row: Number(start.slice(1)) };
    const to = { column: end.charCodeAt(0) - 64, row: Number(end.slice(1)) };
    const cells = () => Array.from({ length: to.row - from.row + 1 }, (_, r) => Array.from({ length: to.column - from.column + 1 }, (_, c) => `${letters(from.column + c)}${from.row + r}`));
    const read = (name: string) => { const value = other.get(name) ?? ""; return typeof value === "string" && value.startsWith("'") ? value.slice(1) : value; };
    return {
      address: `Итог!${address}`,
      load: () => undefined,
      get values() { return cells().map((row) => row.map(read)); },
      set values(source: unknown[][]) { cells().forEach((row, r) => row.forEach((name, c) => other.set(name, source[r][c]))); },
      get formulas() { return this.values; },
      set formulas(source: unknown[][]) { this.values = source; },
      get valueTypes() { return cells().map((row) => row.map((name) => (read(name) === "" ? "Empty" : typeof read(name) === "number" ? "Double" : "String"))); }
    };
  };
  const second: any = {
    id: "sheet-2", name: "Итог", isNullObject: false, load: () => undefined,
    getRange: range,
    getUsedRangeOrNullObject: () => ({ isNullObject: other.size === 0 || [...other.values()].every((value) => value === ""), address: "Итог!A1", load: () => undefined })
  };
  const run = (globalThis as any).Excel.run;
  (globalThis as any).Excel.run = async (fn: any) => run(async (ctx: any) => {
    const worksheets = ctx.workbook.worksheets;
    const first = worksheets.getItem();
    ctx.workbook.worksheets = {
      ...worksheets,
      getItem: (name: string) => (name === "Итог" || name === "sheet-2" ? second : first),
      getItemOrNullObject: (name: string) => (name === "Итог" ? second : { isNullObject: true, load: () => undefined })
    };
    return fn(ctx);
  });
  return { state, other };
}

test("with a result sheet the source stays as it was, and the copy can be undone", async () => {
  const { state, other } = withSecondSheet(cleanSheet({ ...ORDERS, D2: "'заметка", C7: "=C6+1" }));
  setUndoMonitorReady(true);
  try {
    // Соседние данные и формулы источнику не страшны: он не меняется.
    const plan = await prepareRemoveDuplicatesPlan({ sheet: "Данные", address: "A1:C7", columns: ["Город", "Сумма"], destSheet: "Итог" });
    assert.equal(plan.undoAvailable, true);
    assert.equal(plan.dest?.address, "A1:C5", "шапка и 4 уникальные строки");
    const result = await executeRemoveDuplicatesPlan(plan) as any;
    assert.equal(result.executionState, "verified");
    assert.equal(result.sourceUnchanged, true);
    assert.equal(state.valueOf("A3"), "москва", "источник не тронут");
    assert.equal(other.get("A1"), "'Город");
    assert.equal(other.get("C3"), 3);
    await undoLast();
    assert.ok([...other.values()].every((value) => value === ""), "отмена очистила копию");
  } finally {
    clearUndo();
    setUndoMonitorReady(false);
  }
});

test("a result sheet that is missing, the same, or not empty is refused before any card", async () => {
  withSecondSheet(cleanSheet(ORDERS));
  await assert.rejects(() => prepareRemoveDuplicatesPlan({ sheet: "Данные", address: "A1:C7", columns: ["Город", "Сумма"], destSheet: "Нет такого" }), /Листа «Нет такого» нет.*create_sheet/);
  withSecondSheet(cleanSheet(ORDERS), true);
  await assert.rejects(() => prepareRemoveDuplicatesPlan({ sheet: "Данные", address: "A1:C7", columns: ["Город", "Сумма"], destSheet: "Итог" }), /не пуст/);
  // Лист источника как лист результата — нельзя: копия легла бы на данные.
  const same = withSecondSheet(cleanSheet(ORDERS));
  const run = (globalThis as any).Excel.run;
  (globalThis as any).Excel.run = async (fn: any) => run(async (ctx: any) => {
    const first = ctx.workbook.worksheets.getItem("Данные");
    ctx.workbook.worksheets.getItemOrNullObject = () => ({ ...first, isNullObject: false });
    return fn(ctx);
  });
  await assert.rejects(() => prepareRemoveDuplicatesPlan({ sheet: "Данные", address: "A1:C7", columns: ["Город", "Сумма"], destSheet: "Данные" }), /совпадает с листом источника/);
  assert.equal(same.state.valueOf("A3"), "москва");
});
