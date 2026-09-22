import test from "node:test";
import assert from "node:assert/strict";
import { anchorOf, executeFillRangePlan, prepareFillRangePlan, tableExpansionWarning } from "./excelTools";
import { PLANNED_TOOLS } from "./plans";

/* ---------------------------------------------------------------------------
 * Макет листа, который ведёт себя как Excel с формулами.
 *
 * Ячейки хранят формулы в привычном виде A1. Свойство `formulasR1C1` отдаёт
 * и принимает их в виде R1C1 — относительно самой ячейки, как это делает
 * Excel: `=A1` в F2 читается как `=R[-1]C[-5]`, а та же строка R1C1, записанная
 * в F3, превращается в `=A2`. Имена листов, текст в кавычках и ссылки
 * на столбцы таблиц Excel при этом не трогает — макет тоже.
 *
 * Это сторона Excel, а не проверяемый код: ожидаемые формулы в тестах
 * написаны готовыми строками, как требует план стабилизации.
 * ------------------------------------------------------------------------- */

const MAX_ROW = 1_048_576;
const letters = (index: number) => {
  let value = "";
  for (let left = index; left > 0; left = Math.floor((left - 1) / 26)) value = String.fromCharCode(65 + ((left - 1) % 26)) + value;
  return value;
};
const columnNumber = (text: string) => [...text].reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0);

/** Проход по формуле с пропуском текста, имён листов и скобок таблиц. */
function walk(formula: string, onToken: (match: RegExpExecArray, at: number) => string | null, token: RegExp): string {
  let result = "";
  let index = 0;
  while (index < formula.length) {
    const char = formula[index];
    const skip = (end: number) => { result += formula.slice(index, end); index = end; };
    if (char === '"') { skip(formula.indexOf('"', index + 1) + 1 || formula.length); continue; }
    if (char === "'") { skip(formula.indexOf("'!", index + 1) + 2 || formula.length); continue; }
    if (char === "[" && !/[RC]$/.test(formula.slice(0, index))) { skip(formula.indexOf("]", index) + 1 || formula.length); continue; }
    token.lastIndex = index;
    const match = token.exec(formula);
    const before = formula[index - 1] ?? "";
    if (match && match.index === index && !/[A-Za-z0-9_.\u0400-\u04FF]/.test(before)) {
      const after = formula[index + match[0].length] ?? "";
      if (!/[A-Za-z0-9_(!\[\u0400-\u04FF]/.test(after)) {
        const replaced = onToken(match, index);
        if (replaced !== null) { result += replaced; index += match[0].length; continue; }
      }
    }
    const word = /[A-Za-z_\u0400-\u04FF][\w.\u0400-\u04FF]*/y;
    word.lastIndex = index;
    const name = word.exec(formula);
    if (name && name.index === index) { skip(index + name[0].length); continue; }
    result += char;
    index += 1;
  }
  return result;
}

function toR1C1(formula: unknown, row: number, column: number): unknown {
  if (typeof formula !== "string" || !formula.startsWith("=")) return formula;
  return walk(formula, (m) => {
    const [, colAbs, col, rowAbs, rowText] = m;
    const r = Number(rowText);
    const c = columnNumber(col);
    const rPart = rowAbs ? `R${r}` : r === row ? "R" : `R[${r - row}]`;
    const cPart = colAbs ? `C${c}` : c === column ? "C" : `C[${c - column}]`;
    return rPart + cPart;
  }, /(\$?)([A-Z]{1,3})(\$?)(\d+)/y);
}

function fromR1C1(formula: unknown, row: number, column: number): unknown {
  if (typeof formula !== "string" || !formula.startsWith("=")) return formula;
  return walk(formula, (m) => {
    const [, rAbs, rRel, cAbs, cRel] = m;
    const r = rAbs ? Number(rAbs) : row + (rRel ? Number(rRel) : 0);
    const c = cAbs ? Number(cAbs) : column + (cRel ? Number(cRel) : 0);
    if (r < 1 || r > MAX_ROW || c < 1) return "#REF!";
    return `${cAbs ? "$" : ""}${letters(c)}${rAbs ? "$" : ""}${r}`;
  }, /R(?:(\d+)|\[(-?\d+)\])?C(?:(\d+)|\[(-?\d+)\])?/y);
}

interface SheetOptions {
  /** Старое содержимое ячеек, например { F4: "старое" }. */
  cells?: Record<string, unknown>;
  /** Ячейки, запись в которые Excel молча игнорирует. */
  ignore?: string[];
  /** Excel отказывает в записи всей области после первой ячейки. */
  areaWriteThrows?: boolean;
  /** Возврат первой ячейки после отказа тоже не удаётся. */
  restoreThrows?: boolean;
}

function fillSheet(options: SheetOptions = {}) {
  const grid = new Map<string, unknown>(Object.entries(options.cells ?? {}));
  const key = (row: number, column: number) => `${letters(column)}${row}`;
  const calls = { autoFill: 0, areaWrites: 0, anchorWrites: 0 };
  const sync = { count: 0, afterAnchor: -1, beforeR1C1Read: -1 };

  const evaluate = (content: unknown) => {
    if (typeof content !== "string" || !content.startsWith("=")) return content ?? "";
    return content.includes("#REF!") ? "#REF!" : 1;
  };
  const literal = (value: unknown) => (typeof value === "string" && value.startsWith("'") ? value.slice(1) : value);

  function makeRange(address: string): any {
    const [start, end = start] = address.replace(/^.*!/, "").split(":");
    const cell = (text: string) => { const [, c, r] = /([A-Z]+)(\d+)/.exec(text)!; return { row: Number(r), column: columnNumber(c) }; };
    const from = cell(start);
    const to = cell(end);
    const rows = to.row - from.row + 1;
    const columns = to.column - from.column + 1;
    const matrix = (read: (row: number, column: number) => unknown) =>
      Array.from({ length: rows }, (_, r) => Array.from({ length: columns }, (_, c) => read(from.row + r, from.column + c)));
    const write = (source: unknown[][], convert: (value: unknown, row: number, column: number) => unknown, isArea: boolean) => {
      source.forEach((line, r) => line.forEach((value, c) => {
        const row = from.row + r;
        const column = from.column + c;
        if (options.ignore?.includes(key(row, column))) return;
        grid.set(key(row, column), convert(value, row, column));
      }));
      if (isArea) calls.areaWrites += 1;
    };
    const single = rows === 1 && columns === 1;

    return {
      address: `Продажи!${address}`,
      rowIndex: from.row - 1,
      columnIndex: from.column - 1,
      rowCount: rows,
      columnCount: columns,
      load: () => undefined,
      format: { protection: { locked: false, load: () => undefined } },
      get formulas() { return matrix((row, column) => grid.get(key(row, column)) ?? ""); },
      set formulas(source: unknown[][]) {
        if (single && calls.anchorWrites > 0 && options.restoreThrows) throw new Error("Excel отказал");
        if (single) calls.anchorWrites += 1;
        write(source, literal, !single);
      },
      get values() { return matrix((row, column) => evaluate(grid.get(key(row, column)))); },
      set values(source: unknown[][]) { write(source, literal, !single); },
      get valueTypes() {
        return matrix((row, column) => {
          const value = evaluate(grid.get(key(row, column)));
          return value === "" ? "Empty" : typeof value === "number" ? "Double" : "String";
        });
      },
      get formulasR1C1() {
        if (sync.beforeR1C1Read < 0) sync.beforeR1C1Read = sync.count;
        return matrix((row, column) => toR1C1(grid.get(key(row, column)) ?? "", row, column));
      },
      set formulasR1C1(source: unknown[][]) {
        if (options.areaWriteThrows) throw new Error("Во время обработки запроса произошла внутренняя ошибка.");
        write(source, (value, row, column) => fromR1C1(value, row, column), true);
      },
      // Так Excel повёл себя на проверке 18 сентября 2026 года рядом с таблицей.
      autoFill: () => { calls.autoFill += 1; throw new Error("Во время обработки запроса произошла внутренняя ошибка."); }
    };
  }

  const sheet: any = {
    id: "sheet-1",
    name: "Продажи",
    load: () => undefined,
    protection: { protected: false, load: () => undefined },
    tables: { items: [], load: () => undefined },
    getRange: (address: string) => makeRange(address),
    getRangeByIndexes: (row: number, column: number, rows: number, columns: number) =>
      makeRange(`${letters(column + 1)}${row + 1}:${letters(column + columns)}${row + rows}`)
  };
  (globalThis as any).Office = { context: { document: { url: "C:/fill.xlsx" }, requirements: { isSetSupported: () => true } } };
  (globalThis as any).Excel = {
    AutoFillType: { fillDefault: "FillDefault" },
    run: async (fn: any) => fn({
      workbook: {
        application: { calculationMode: "automatic", load: () => undefined },
        worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet }
      },
      sync: async () => {
        sync.count += 1;
        if (calls.anchorWrites > 0 && sync.afterAnchor < 0) sync.afterAnchor = sync.count;
      }
    })
  };
  const column = (name: string, from: number, to: number) =>
    Array.from({ length: to - from + 1 }, (_, index) => grid.get(`${name}${from + index}`) ?? "");
  return { grid, calls, sync, column };
}

const fill = async (address: string, value: string | number, isFormula = true) =>
  executeFillRangePlan(await prepareFillRangePlan({ sheet: "Продажи", address, value, isFormula })) as Promise<any>;

/* --- основное ---------------------------------------------------------------- */

test("the anchor cell of an area is its top left corner", () => {
  assert.equal(anchorOf("F2:F6"), "F2");
  assert.equal(anchorOf("B3:D10"), "B3");
  assert.equal(anchorOf("AA5:AC9"), "AA5");
  assert.equal(anchorOf("C7"), "C7");
});

test("filling goes through the plan registry like every other change", () => {
  assert.ok(PLANNED_TOOLS.includes("fill_range"));
});

test("one formula fills the whole area with references moved row by row", async () => {
  const state = fillSheet();
  const result = await fill("F2:F6", "=D2*E2");
  assert.equal(result.executionState, "verified");
  assert.deepEqual(state.column("F", 2, 6), ["=D2*E2", "=D3*E3", "=D4*E4", "=D5*E5", "=D6*E6"]);
});

/* --- S1: примеры из плана стабилизации ---------------------------------------- */

test("a sheet name that looks like a cell is left alone", async () => {
  // План стабилизации, S1: собственный сдвиг формул превращал =Q1!A1 в =Q2!A2.
  const state = fillSheet();
  await fill("F2:F4", "=Q1!A1");
  assert.deepEqual(state.column("F", 2, 4), ["=Q1!A1", "=Q1!A2", "=Q1!A3"]);
});

test("a quoted sheet name is left alone", async () => {
  const state = fillSheet();
  await fill("F2:F3", "='Q1 2026'!A1");
  assert.deepEqual(state.column("F", 2, 3), ["='Q1 2026'!A1", "='Q1 2026'!A2"]);
});

test("a table column named like a cell is left alone", async () => {
  const state = fillSheet();
  await fill("F2:F3", "=SUM(Sales[Q1])");
  assert.deepEqual(state.column("F", 2, 3), ["=SUM(Sales[Q1])", "=SUM(Sales[Q1])"]);
});

test("dollars pin what they stand in front of, and filling right moves columns", async () => {
  const down = fillSheet();
  await fill("F2:F3", "=$A1*B$1");
  assert.deepEqual(down.column("F", 2, 3), ["=$A1*B$1", "=$A2*B$1"]);

  const right = fillSheet();
  await fill("F2:H2", "=A2");
  assert.deepEqual(["F2", "G2", "H2"].map((cell) => right.grid.get(cell)), ["=A2", "=B2", "=C2"]);
});

test("text in quotes that looks like a reference is left alone", async () => {
  const state = fillSheet();
  await fill("F2:F3", '=IF(A1>0,"B2 штук","")');
  assert.deepEqual(state.column("F", 2, 3), ['=IF(A1>0,"B2 штук","")', '=IF(A2>0,"B2 штук","")']);
});

/* --- S1: сверка результата ------------------------------------------------------ */

test("a cell Excel left with its old content is not verified", async () => {
  // Частично проигнорированная запись: F4 сохранила старое значение.
  fillSheet({ cells: { F4: "старое" }, ignore: ["F4"] });
  await assert.rejects(() => fill("F2:F6", "=D2*E2"), (error: any) => {
    assert.equal(error.executionState, "applied");
    assert.match(error.message, /F4/);
    return true;
  });
});

test("a formula that points off the sheet is not verified", async () => {
  // Протяжка вниз уводит ссылку за последнюю строку — Excel ставит #ССЫЛКА!.
  fillSheet();
  await assert.rejects(() => fill("F2:F3", "=A1048576"), (error: any) => {
    assert.equal(error.executionState, "applied");
    assert.match(error.message, /#REF!|ссылк/i);
    return true;
  });
});

test("an area already holding the requested fill is verified, not a failure", async () => {
  fillSheet({ cells: { F2: "=D2*E2", F3: "=D3*E3" } });
  const result = await fill("F2:F3", "=D2*E2");
  assert.equal(result.executionState, "verified");
});

test("a plain value is written into every cell, not continued as a series", async () => {
  // autoFill продолжает «Товар 1» рядом «Товар 2, Товар 3»; значение — это значение.
  const state = fillSheet();
  const result = await fill("F2:F4", "Товар 1", false);
  assert.equal(result.executionState, "verified");
  assert.deepEqual(state.column("F", 2, 4), ["Товар 1", "Товар 1", "Товар 1"]);
});

/* --- S1: отказ после записи первой ячейки ------------------------------------ */

test("the formula is read back only after the first cell is committed", async () => {
  // Свойство formulasR1C1 читается у ячейки: новой формулы до записи в ней нет.
  const state = fillSheet();
  await fill("F2:F4", "=D2*E2");
  assert.ok(state.sync.afterAnchor > 0, "первая ячейка записана и отправлена");
  assert.ok(state.sync.beforeR1C1Read >= state.sync.afterAnchor, "чтение R1C1 — после этой отправки");
  assert.equal(state.calls.autoFill, 0, "протяжка Excel больше не используется");
});

test("if the area write fails, the first cell is put back and nothing is reported as done", async () => {
  const state = fillSheet({ cells: { F2: "прежнее" }, areaWriteThrows: true });
  await assert.rejects(() => fill("F2:F4", "=D2*E2"), (error: any) => {
    assert.equal(error.executionState, "failed_before_write");
    assert.match(error.message, /первая ячейка возвращена/i);
    return true;
  });
  assert.equal(state.grid.get("F2"), "прежнее");
  assert.equal(state.grid.get("F3"), undefined, "остальное не тронуто");
});

test("if even putting the first cell back fails, the outcome is unknown", async () => {
  fillSheet({ cells: { F2: "прежнее" }, areaWriteThrows: true, restoreThrows: true });
  await assert.rejects(() => fill("F2:F4", "=D2*E2"), (error: any) => {
    assert.equal(error.executionState, "unknown");
    return true;
  });
});

/* --- прочее ------------------------------------------------------------------------ */

test("a formula is required when isFormula is set", async () => {
  fillSheet();
  await assert.rejects(
    () => prepareFillRangePlan({ sheet: "Продажи", address: "F2:F6", value: "D2*E2", isFormula: true }),
    /начинающейся со знака равенства/
  );
});

test("the preview counts the cells that will be overwritten", async () => {
  fillSheet({ cells: { F2: 1, F3: 2 } });
  const plan = await prepareFillRangePlan({ sheet: "Продажи", address: "F2:F6", value: 0 });
  assert.equal(plan.occupiedCells, 2);
});

test("writing next to a table is flagged before it silently grows it", () => {
  const sales = [{ name: "SalesTable", address: "Продажи!A1:G6" }];
  assert.match(tableExpansionWarning("Продажи!H2:H6", sales) ?? "", /вплотную примыкает/);
  assert.match(tableExpansionWarning("Продажи!A7:G7", sales) ?? "", /вплотную примыкает/);
  assert.match(tableExpansionWarning("Продажи!C3", sales) ?? "", /внутри таблицы/);
  assert.equal(tableExpansionWarning("Продажи!J2:J6", sales), null);
  assert.equal(tableExpansionWarning("Продажи!H20:H25", sales), null);
  assert.equal(tableExpansionWarning("Продажи!H2:H6", []), null);
});
