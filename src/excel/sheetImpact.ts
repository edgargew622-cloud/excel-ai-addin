/**
 * Что заденет переименование или удаление листа (этап 7, 7.3.1).
 *
 * Замер в Excel 24 сентября 2026 года:
 * - переименование сохраняет ID листа; обычные ссылки (`=Лист!A1`) и
 *   именованные диапазоны Excel переписывает сам, а имя внутри текста —
 *   `INDIRECT("Лист!A1")` — нет: такая формула становится `#ССЫЛКА!`;
 * - удаление превращает ссылки на лист в `#REF!`, именованные диапазоны
 *   на нём — тоже, и все формулы, которые ими пользуются.
 *
 * Модуль не обращается к Excel и проверяется без него.
 */

import { formulaReferences } from "./rowOps";

export interface ScannedSheetLike {
  name: string;
  rowIndex: number;
  columnIndex: number;
  formulas: readonly (readonly unknown[])[];
  values: readonly (readonly unknown[])[];
}

const same = (a: string, b: string) => a.trim().toLowerCase() === b.trim().toLowerCase();

/** Формула ссылается на лист обычной ссылкой: `Лист!A1`, `'Лист 2'!A1:B3`. */
export function referencesSheet(formula: unknown, sheet: string): boolean {
  return formulaReferences(formula).some((reference) => reference.sheet !== null && same(reference.sheet, sheet));
}

/** Имя листа стоит внутри текста формулы — `INDIRECT("Лист!A1")`: Excel его не перепишет. */
export function literalMentionsSheet(formula: unknown, sheet: string): boolean {
  if (typeof formula !== "string" || !formula.startsWith("=")) return false;
  const wanted = sheet.trim().toLowerCase();
  return (formula.match(/"(?:[^"]|"")*"/g) ?? []).some((literal) => literal.toLowerCase().includes(wanted));
}

/** Формула пользуется именованным диапазоном — целое слово, не часть другого имени. */
export function usesName(formula: unknown, name: string): boolean {
  if (typeof formula !== "string" || !formula.startsWith("=")) return false;
  const bare = formula.replace(/"(?:[^"]|"")*"/g, '""');
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return new RegExp(`(?<![A-Za-z0-9_.\\u0400-\\u04FF])${escaped}(?![A-Za-z0-9_.\\u0400-\\u04FF(])`, "i").test(bare);
}

export interface CellMention {
  sheet: string;
  cell: string;
  formula: string;
}

export interface SheetImpact {
  /** Формулы других листов с обычными ссылками на лист. */
  referencing: CellMention[];
  /** Формулы с именем листа внутри текста — INDIRECT, HYPERLINK. */
  literal: CellMention[];
  /** Ячейки, где имя листа упомянуто просто текстом: их никто не перепишет. */
  textMentions: CellMention[];
  /** Формулы, которые пользуются именованными диапазонами на этом листе. */
  viaNames: CellMention[];
  /** Сколько не вошло в списки сверх предела. */
  overflow: number;
}

const MAX_LISTED = 20;

export function sheetImpact(
  sheets: readonly ScannedSheetLike[],
  target: string,
  namesOnSheet: readonly string[],
  columnName: (index: number) => string
): SheetImpact {
  const impact: SheetImpact = { referencing: [], literal: [], textMentions: [], viaNames: [], overflow: 0 };
  const add = (list: CellMention[], item: CellMention) => {
    if (list.length < MAX_LISTED) list.push(item);
    else impact.overflow += 1;
  };
  const wanted = target.trim().toLowerCase();
  for (const sheet of sheets) {
    const own = same(sheet.name, target);
    sheet.formulas.forEach((row, r) => row.forEach((formula, c) => {
      const cell = `${columnName(sheet.columnIndex + c + 1)}${sheet.rowIndex + r + 1}`;
      const text = String(formula ?? "");
      if (typeof formula === "string" && formula.startsWith("=")) {
        if (!own && referencesSheet(formula, target)) add(impact.referencing, { sheet: sheet.name, cell, formula: text });
        if (literalMentionsSheet(formula, target)) add(impact.literal, { sheet: sheet.name, cell, formula: text });
        if (!own && namesOnSheet.some((name) => usesName(formula, name))) add(impact.viaNames, { sheet: sheet.name, cell, formula: text });
        return;
      }
      if (typeof formula === "string" && formula.toLowerCase().includes(wanted)) add(impact.textMentions, { sheet: sheet.name, cell, formula: text });
    }));
  }
  return impact;
}
