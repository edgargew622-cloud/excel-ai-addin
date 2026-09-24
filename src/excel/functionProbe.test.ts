import test from "node:test";
import assert from "node:assert/strict";
import {
  errorCells,
  errorKind,
  errorNote,
  functionNamesIn,
  knownMissing,
  probeCellFor,
  probeFunctions,
  resetFunctionAvailability,
  uncheckedFunctions
} from "./functionProbe";

const letters = (index: number) => {
  let value = "";
  for (let left = index; left > 0; left = Math.floor((left - 1) / 26)) value = String.fromCharCode(65 + ((left - 1) % 26)) + value;
  return value;
};

test("errors are recognised in both languages, as Excel reported them", () => {
  // Замер 24 сентября 2026 года: русский Excel отдаёт ошибки по-русски.
  assert.equal(errorKind("#ИМЯ?"), "name");
  assert.equal(errorKind("#NAME?"), "name");
  assert.equal(errorKind("#ДЕЛ/0!"), "div0");
  assert.equal(errorKind("#Н/Д"), "na");
  assert.equal(errorKind("#ЗНАЧ!"), "value");
  assert.equal(errorKind("#ЧИСЛО!"), "num");
  assert.equal(errorKind("#ПУСТО!"), "null");
  assert.equal(errorKind("#ССЫЛКА!"), "ref");
  assert.equal(errorKind("#ПЕРЕНОС!"), "spill");
  assert.equal(errorKind("что-то"), "other");
});

test("an error cell is found by its type, whatever the language of its text", () => {
  // Прежде поиск сравнивал текст с английскими подписями и на русском
  // Excel не находил ни одной ошибки.
  const cells = errorCells(
    [[1, "#ДЕЛ/0!"], ["#ИМЯ?", "#ЧТО-ТО?"]],
    [["Double", "Error"], ["Error", "Error"]],
    (r, c) => `${letters(c + 1)}${r + 1}`
  );
  assert.deepEqual(cells.map((item) => `${item.cell} ${item.kind}`), ["B1 div0", "A2 name", "B2 other"]);
  // Без valueTypes — по подписям на обоих языках; текст «#хэштег» ошибкой не считается.
  assert.deepEqual(errorCells([["#Н/Д", "#хэштег"]], undefined, (r, c) => `${r}:${c}`).map((item) => item.kind), ["na"]);
});

test("the note explains every kind of new error and asks for an explicit replacement of a missing function", () => {
  const note = errorNote([{ cell: "E2", text: "#ИМЯ?", kind: "name" }, { cell: "E3", text: "#ДЕЛ/0!", kind: "div0" }])!;
  assert.match(note, /#ИМЯ\? — функция или имя/);
  assert.match(note, /#ДЕЛ\/0! — деление на ноль/);
  assert.match(note, /предложи замену явно/);
  assert.equal(errorNote([]), undefined);
});

test("function names are taken from the formula, not from text or sheet names", () => {
  assert.deepEqual(functionNamesIn('=IFERROR(XLOOKUP(A2,B:B,C:C),"нет (данных)")'), ["IFERROR", "XLOOKUP"]);
  assert.deepEqual(functionNamesIn("='Отчёт (Q1)'!A1+sum(B1:B3)"), ["SUM"]);
  assert.deepEqual(functionNamesIn("=СУММ(A1:A3)"), ["СУММ"]);
  assert.deepEqual(functionNamesIn("=_xlfn.XLOOKUP(1,A:A,B:B)"), ["_XLFN.XLOOKUP"]);
  assert.deepEqual(functionNamesIn("=A1*2"), []);
  assert.deepEqual(functionNamesIn("просто текст (с скобкой)"), []);
});

test("the probe cell stands past both the data and the target, on the first data row", () => {
  assert.equal(probeCellFor({ rowIndex: 0, columnIndex: 0, columnCount: 4 }, { rowIndex: 1, columnIndex: 4, columnCount: 1 }, letters), "G1");
  // Цель шире данных — ячейка правее цели.
  assert.equal(probeCellFor({ rowIndex: 2, columnIndex: 0, columnCount: 3 }, { rowIndex: 2, columnIndex: 5, columnCount: 4 }, letters), "K3");
  assert.equal(probeCellFor(null, { rowIndex: 0, columnIndex: 16_380, columnCount: 3 }, letters), null);
});

/** Ячейка, которая ведёт себя как Excel из замера. */
function probeExcel(known: string[], missing: string[]) {
  let formula = "";
  const writes: string[] = [];
  const cell: any = {
    get formulas() { return [[formula]]; },
    set formulas(value: unknown[][]) {
      const text = String(value[0][0]);
      const name = text.slice(1, text.indexOf("("));
      if (known.includes(name)) throw Object.assign(new Error("Аргумент недопустим"), { code: "InvalidArgument" });
      writes.push(text);
      formula = text;
    },
    get values() { return [[missing.includes(formula.slice(1, formula.indexOf("("))) ? "#ИМЯ?" : 3.14]]; },
    get valueTypes() { return [[missing.includes(formula.slice(1, formula.indexOf("("))) ? "Error" : "Double"]]; },
    load: () => undefined,
    clear: () => { formula = ""; }
  };
  const ctx: any = { sync: async () => undefined };
  const sheet: any = { getRange: () => cell };
  return { ctx, sheet, writes, cell };
}

test("the probe tells a missing function from a present one and remembers the answer", async () => {
  resetFunctionAvailability();
  const excel = probeExcel(["XLOOKUP"], ["TEXTSPLIT"]);
  const result = await probeFunctions(excel.ctx, excel.sheet, "G1", ["XLOOKUP", "TEXTSPLIT", "PI"]);
  assert.deepEqual(result, { available: ["XLOOKUP", "PI"], unavailable: ["TEXTSPLIT"], unchecked: [] });
  assert.equal(excel.cell.formulas[0][0], "", "временная ячейка очищена");
  // Запомнено: повторно эти функции не проверяются, недоступная отклоняется сразу.
  assert.deepEqual(uncheckedFunctions(["=XLOOKUP(1,A:A,B:B)+PI()"]), []);
  assert.deepEqual(knownMissing(['=TEXTSPLIT(A1,",")']), ["TEXTSPLIT"]);
});

test("a probe cell that does not come clean is not passed over in silence", async () => {
  resetFunctionAvailability();
  const excel = probeExcel([], []);
  excel.cell.clear = () => undefined;
  await assert.rejects(() => probeFunctions(excel.ctx, excel.sheet, "G1", ["PI"]), /не очистилась/);
});
