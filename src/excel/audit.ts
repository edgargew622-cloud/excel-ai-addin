/**
 * Аудит книги — только чтение (этап 7, 7.5).
 *
 * Отчёт различает три вида выводов, и у каждого есть адрес и прочитанная
 * формула или значение:
 * - доказанное — Excel сам показывает ошибку, в формуле ссылка #ССЫЛКА!,
 *   контрольная ячейка не равна нулю;
 * - подозрение — формула не такая, как у соседей по ряду, число вместо
 *   формулы в ряду формул, ссылка на пустую ячейку, число внутри формулы;
 * - непроверенное — внешние книги, INDIRECT/OFFSET (адрес вычисляется),
 *   слишком большие листы.
 *
 * Ошибки делятся на исходные и следствия: формула, которая ссылается на
 * ячейку с ошибкой, ошибается из-за неё, и чинить надо исходную.
 *
 * Разбор ничего не меняет в книге. Исправления — отдельная просьба.
 */

import { parseA1Rect } from "./a1";
import { errorKind, ERROR_MEANING } from "./functionProbe";
import { columnLetters } from "./formulaFill";
import { formulaReferences } from "./rowOps";

export interface AuditSheet {
  name: string;
  rowIndex: number;
  columnIndex: number;
  formulas: readonly (readonly unknown[])[];
  formulasR1C1: readonly (readonly unknown[])[];
  values: readonly (readonly unknown[])[];
  valueTypes: readonly (readonly unknown[])[];
}

export interface Finding {
  sheet: string;
  cell: string;
  /** Формула или значение, как их прочитала панель. */
  content: string;
  value?: string;
  reason: string;
}

export interface AuditReport {
  proven: Finding[];
  suspicions: Finding[];
  unverified: Finding[];
  totals: { proven: number; suspicions: number; unverified: number; errorsRoot: number; errorsConsequence: number; formulas: number };
}

const MAX_LISTED = 30;
const MAX_CONSTANTS = 15;

const isFormula = (value: unknown): value is string => typeof value === "string" && value.startsWith("=");
const same = (a: string, b: string) => a.trim().toLowerCase() === b.trim().toLowerCase();
const shown = (value: unknown) => (typeof value === "string" ? value : JSON.stringify(value));

/** Числа внутри формулы, кроме 0, 1 и адресов: `=B2*1.05` → [1.05]. */
export function constantsIn(formula: string): number[] {
  const bare = formula.replace(/"(?:[^"]|"")*"/g, '""').replace(/'[^']*'!/g, "");
  const found: number[] = [];
  const pattern = /(?<![A-Za-z0-9_.$Ѐ-ӿ:])(\d+(?:\.\d+)?)(?![A-Za-z0-9_(:!Ѐ-ӿ])/g;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(bare))) {
    // Номер строки в ссылке «$B$2» или «B2» отсекается просмотром назад; «2:2» — двоеточием.
    const number = Number(match[1]);
    if (number !== 0 && number !== 1) found.push(number);
  }
  return found;
}

export function auditSheets(sheets: readonly AuditSheet[], options: { checks?: { sheet: string; address: string }[]; unscanned?: readonly string[] } = {}): AuditReport {
  const report: AuditReport = {
    proven: [], suspicions: [], unverified: [],
    totals: { proven: 0, suspicions: 0, unverified: 0, errorsRoot: 0, errorsConsequence: 0, formulas: 0 }
  };
  const add = (list: "proven" | "suspicions" | "unverified", item: Finding) => {
    report.totals[list] += 1;
    if (report[list].length < MAX_LISTED) report[list].push(item);
  };
  const byName = (name: string) => sheets.find((sheet) => same(sheet.name, name));
  const at = (sheet: AuditSheet, row: number, column: number) => {
    const r = row - sheet.rowIndex - 1;
    const c = column - sheet.columnIndex - 1;
    return { formula: sheet.formulas[r]?.[c], value: sheet.values[r]?.[c], type: sheet.valueTypes[r]?.[c], r1c1: sheet.formulasR1C1[r]?.[c] };
  };
  const name = (sheet: AuditSheet, r: number, c: number) => `${columnLetters(sheet.columnIndex + c + 1)}${sheet.rowIndex + r + 1}`;

  // Ячейки с ошибкой на каждом листе — для деления на исходные и следствия.
  const errors = new Map<string, { row: number; column: number }[]>();
  for (const sheet of sheets) {
    const list: { row: number; column: number }[] = [];
    sheet.valueTypes.forEach((row, r) => row.forEach((type, c) => {
      if (type === "Error") list.push({ row: sheet.rowIndex + r + 1, column: sheet.columnIndex + c + 1 });
    }));
    errors.set(sheet.name.toLowerCase(), list);
  }
  const touchesError = (sheetName: string, reference: { rowStart: number; rowEnd: number; columnStart: number; columnEnd: number }) =>
    (errors.get(sheetName.toLowerCase()) ?? []).some((cell) =>
      cell.row >= reference.rowStart && cell.row <= reference.rowEnd && cell.column >= reference.columnStart && cell.column <= reference.columnEnd);

  for (const sheet of sheets) {
    sheet.formulas.forEach((row, r) => row.forEach((formula, c) => {
      const cell = name(sheet, r, c);
      const value = sheet.values[r]?.[c];
      const type = sheet.valueTypes[r]?.[c];
      if (!isFormula(formula)) return;
      report.totals.formulas += 1;
      const references = formulaReferences(formula);
      const bare = formula.replace(/"(?:[^"]|"")*"/g, '""');

      // Доказанное: ссылка, которую Excel уже потерял.
      if (/#REF!|#ССЫЛКА!/i.test(bare)) {
        add("proven", { sheet: sheet.name, cell, content: formula, value: shown(value), reason: "в формуле потерянная ссылка #ССЫЛКА!: ячейка или лист, на который она указывала, удалены" });
      } else if (type === "Error") {
        const consequence = references.some((reference) => {
          const target = reference.sheet ?? sheet.name;
          return byName(target) && touchesError(byName(target)!.name, reference) &&
            !(same(target, sheet.name) && reference.rowStart === reference.rowEnd && reference.columnStart === reference.columnEnd &&
              reference.rowStart === sheet.rowIndex + r + 1 && reference.columnStart === sheet.columnIndex + c + 1);
        });
        if (consequence) {
          report.totals.errorsConsequence += 1;
        } else {
          report.totals.errorsRoot += 1;
          add("proven", { sheet: sheet.name, cell, content: formula, value: shown(value), reason: `ошибка Excel: ${ERROR_MEANING[errorKind(value)]}` });
        }
      }

      // Непроверенное: адрес вычисляется или лежит в другой книге.
      if (/\[[^\]]+\][^!]*!/.test(bare)) add("unverified", { sheet: sheet.name, cell, content: formula, reason: "ссылка на другую книгу: её данные панель не видит" });
      else if (/\b(INDIRECT|OFFSET|ДВССЫЛ|СМЕЩ)\s*\(/i.test(bare)) add("unverified", { sheet: sheet.name, cell, content: formula, reason: "адрес вычисляется (INDIRECT/OFFSET): куда формула смотрит, по тексту не проверить" });

      // Подозрение: одиночная ссылка на пустую ячейку.
      for (const reference of references) {
        if (reference.rowStart !== reference.rowEnd || reference.columnStart !== reference.columnEnd || reference.wholeColumns || reference.wholeRows) continue;
        const target = byName(reference.sheet ?? sheet.name);
        if (!target) continue;
        const pointed = at(target, reference.rowStart, reference.columnStart);
        // Вне занятой области ячейка тоже пуста: занятая область включает все непустые.
        if (pointed.formula === "" || pointed.formula === null || pointed.formula === undefined) {
          add("suspicions", { sheet: sheet.name, cell, content: formula, reason: `ссылается на пустую ячейку ${reference.sheet ? `${target.name}!` : ""}${columnLetters(reference.columnStart)}${reference.rowStart}: вход не заполнен?` });
          break;
        }
      }
    }));

    // Подозрение: формула не такая, как у обоих соседей, или число среди одинаковых формул.
    const checkNeighbours = (r: number, c: number, dr: number, dc: number, direction: string) => {
      const here = sheet.formulasR1C1[r]?.[c];
      const before = sheet.formulasR1C1[r - dr]?.[c - dc];
      const after = sheet.formulasR1C1[r + dr]?.[c + dc];
      if (!isFormula(before) || !isFormula(after) || before !== after) return;
      const cell = name(sheet, r, c);
      const content = sheet.formulas[r]?.[c];
      if (isFormula(here) && here !== before) {
        add("suspicions", { sheet: sheet.name, cell, content: String(content), reason: `формула не такая, как у соседей ${direction} (${String(sheet.formulas[r - dr]?.[c - dc])} и ${String(sheet.formulas[r + dr]?.[c + dc])})` });
      } else if (typeof here === "number") {
        add("suspicions", { sheet: sheet.name, cell, content: String(content), reason: `число вместо формулы: соседи ${direction} считают одинаковой формулой — значение вписано вручную?` });
      }
    };
    const flagged = new Set<string>();
    sheet.formulasR1C1.forEach((row, r) => row.forEach((_, c) => {
      const before = report.totals.suspicions;
      checkNeighbours(r, c, 0, 1, "слева и справа");
      if (report.totals.suspicions === before) checkNeighbours(r, c, 1, 0, "сверху и снизу");
      if (report.totals.suspicions !== before) flagged.add(`${r},${c}`);
    }));

    // Подозрение: число внутри формулы — допущение, спрятанное от входов.
    let constants = 0;
    sheet.formulas.forEach((row, r) => row.forEach((formula, c) => {
      if (!isFormula(formula) || flagged.has(`${r},${c}`) || constants >= MAX_CONSTANTS) return;
      const numbers = constantsIn(formula);
      if (!numbers.length) return;
      constants += 1;
      add("suspicions", { sheet: sheet.name, cell: name(sheet, r, c), content: formula, reason: `число ${numbers.join(", ")} внутри формулы: допущение лучше держать во входной ячейке` });
    }));
  }

  // Контрольные ячейки: пользователь сказал, что там должен быть ноль.
  for (const check of options.checks ?? []) {
    const sheet = byName(check.sheet);
    const rect = parseA1Rect(check.address);
    if (!sheet || !rect) { add("unverified", { sheet: check.sheet, cell: check.address, content: "", reason: "контрольная область не прочитана" }); continue; }
    for (let row = rect.rowStart; row <= rect.rowEnd; row++) {
      for (let column = rect.columnStart; column <= rect.columnEnd; column++) {
        const cell = at(sheet, row, column);
        const address = `${columnLetters(column)}${row}`;
        if (cell.type === "Error") continue; // уже в списке ошибок
        if (typeof cell.value !== "number") {
          add("unverified", { sheet: sheet.name, cell: address, content: shown(cell.formula ?? ""), value: shown(cell.value ?? ""), reason: "контрольная ячейка не содержит числа" });
        } else if (Math.abs(cell.value) > 0.005) {
          add("proven", { sheet: sheet.name, cell: address, content: shown(cell.formula), value: shown(cell.value), reason: `контрольное равенство не сходится: ${cell.value} вместо 0` });
        }
      }
    }
  }
  for (const name of options.unscanned ?? []) add("unverified", { sheet: name, cell: "", content: "", reason: "лист слишком большой для разбора: не проверен" });
  return report;
}

/* --- чтение книги ------------------------------------------------------------ */

const MAX_AUDIT_SHEETS = 30;

/** Контрольная область «Лист!B30:F30» или «B30:F30» на листе по умолчанию. */
function parseCheck(text: string, fallbackSheet: string): { sheet: string; address: string } {
  const bang = text.lastIndexOf("!");
  if (bang === -1) return { sheet: fallbackSheet, address: text.replace(/\$/g, "") };
  return { sheet: text.slice(0, bang).replace(/^'|'$/g, "").replace(/''/g, "'"), address: text.slice(bang + 1).replace(/\$/g, "") };
}

export async function auditWorkbook(args: { sheet?: string; checks?: string[] }) {
  return Excel.run(async (ctx) => {
    const collection = ctx.workbook.worksheets;
    collection.load("items/name");
    const active = ctx.workbook.worksheets.getActiveWorksheet();
    active.load("name");
    await ctx.sync();
    let items = [...collection.items];
    if (args.sheet) {
      items = items.filter((item) => same(item.name, String(args.sheet)));
      if (!items.length) throw new Error(`Листа «${args.sheet}» в книге нет.`);
    }
    const unscanned: string[] = [];
    if (items.length > MAX_AUDIT_SHEETS) {
      unscanned.push(...items.slice(MAX_AUDIT_SHEETS).map((item) => item.name));
      items = items.slice(0, MAX_AUDIT_SHEETS);
    }
    const used = items.map((sheet) => {
      const range = sheet.getUsedRangeOrNullObject(true);
      range.load(["isNullObject", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
      return { sheet, range };
    });
    await ctx.sync();
    const wanted = used.filter(({ sheet, range }) => {
      if (range.isNullObject) return false;
      if (range.rowCount * range.columnCount > 20_000) { unscanned.push(sheet.name); return false; }
      range.load(["formulas", "formulasR1C1", "values", "valueTypes"]);
      return true;
    });
    if (wanted.length) await ctx.sync();
    const sheets: AuditSheet[] = wanted.map(({ sheet, range }) => ({
      name: sheet.name,
      rowIndex: range.rowIndex,
      columnIndex: range.columnIndex,
      formulas: range.formulas as unknown[][],
      formulasR1C1: range.formulasR1C1 as unknown[][],
      values: range.values as unknown[][],
      valueTypes: range.valueTypes as unknown[][]
    }));
    const checks = (args.checks ?? []).map((text) => parseCheck(text, args.sheet ?? active.name));
    const report = auditSheets(sheets, { checks, unscanned });
    return {
      scannedSheets: sheets.map((sheet) => sheet.name),
      ...report,
      note:
        "Отчёт только читает книгу. proven — доказано самим Excel или контрольным равенством; suspicions — подозрения, их подтверждает человек; " +
        "unverified — места, которые панель проверить не может. Единицы, валюты, периоды и допущения панель не проверяет. " +
        "Исправления — отдельная просьба и отдельное подтверждение." +
        (report.totals.errorsConsequence ? ` Ещё ${report.totals.errorsConsequence} ячеек с ошибкой — следствия исходных: они исправятся вместе с ними.` : "")
    };
  });
}
