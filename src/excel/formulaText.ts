/**
 * Когда две формулы — одна и та же формула.
 *
 * Excel хранит формулу не в том виде, в каком её записали, а в своём.
 * Проверка в Excel 23 сентября 2026 года: `=Q1!A1` он сохранил как
 * `='Q1'!A1` — имя листа похоже на адрес ячейки, и Excel взял его
 * в кавычки. Сверка записи сравнивала текст побуквенно, объявила
 * расхождение и остановила задачу, хотя записано было ровно то, что просили.
 *
 * Правило намеренно узкое, это не разбор формул:
 *
 * - кавычки вокруг имени листа без пробелов и особых знаков не важны —
 *   `'Q1'!A1` и `Q1!A1` Excel считает одним и тем же;
 * - регистр вне текстовых строк не важен — имена функций, листов, таблиц
 *   и адреса Excel не различает по регистру и сам пишет `sum` как `SUM`;
 * - всё внутри двойных кавычек сравнивается как есть: `"да"` и `"ДА"` —
 *   разные строки, и формулы с ними разные.
 */

/** Имя листа в кавычках, которому кавычки не нужны: буквы, цифры, точка, подчёркивание. */
const NEEDLESSLY_QUOTED_SHEET = /'([A-Za-z0-9_.Ѐ-ӿ]+)'!/g;

export function canonicalFormula(formula: string): string {
  let result = "";
  let index = 0;
  while (index < formula.length) {
    if (formula[index] === '"') {
      // Строка Excel: удвоенная кавычка внутри — это сама кавычка.
      let end = index + 1;
      while (end < formula.length) {
        if (formula[end] === '"' && formula[end + 1] === '"') { end += 2; continue; }
        if (formula[end] === '"') break;
        end += 1;
      }
      result += formula.slice(index, end + 1);
      index = end + 1;
      continue;
    }
    let next = formula.indexOf('"', index);
    if (next < 0) next = formula.length;
    result += formula.slice(index, next).replace(NEEDLESSLY_QUOTED_SHEET, "$1!").toUpperCase();
    index = next;
  }
  return result;
}

/**
 * Совпадает ли содержимое ячейки с записанным.
 *
 * Для формул — по правилу выше, для всего остального — строго, как раньше:
 * значение Excel не переписывает, и любое расхождение в нём настоящее.
 */
export function sameCellContent(actual: unknown, expected: unknown): boolean {
  if (typeof actual === "string" && typeof expected === "string" && actual.startsWith("=") && expected.startsWith("=")) {
    return canonicalFormula(actual) === canonicalFormula(expected);
  }
  return JSON.stringify(actual ?? "") === JSON.stringify(expected ?? "");
}

export function sameCellMatrix(actual: readonly (readonly unknown[])[], expected: readonly (readonly unknown[])[]): boolean {
  if (actual.length !== expected.length) return false;
  return expected.every((row, r) =>
    (actual[r]?.length ?? -1) === row.length && row.every((cell, c) => sameCellContent(actual[r][c], cell)));
}
