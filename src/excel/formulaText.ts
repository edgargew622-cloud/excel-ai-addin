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
export function sameCellContent(actual: unknown, expected: unknown, typed = false): boolean {
  if (typeof actual === "string" && typeof expected === "string" && actual.startsWith("=") && expected.startsWith("=")) {
    return canonicalFormula(actual) === canonicalFormula(expected);
  }
  if (typed && typeof expected === "string" && typeof actual !== "string" && typedLiteral(expected, actual)) return true;
  return JSON.stringify(actual ?? "") === JSON.stringify(expected ?? "");
}

/**
 * Запись через range.formulas — это ввод: текст «2027» Excel хранит числом,
 * «5%» — долей, «31.12.2026» — датой, «TRUE» — логическим значением.
 * Книга 1567 (26.10.2026): заголовок столбца «2027» дважды остановил задачу
 * «обратное чтение отличается от плана», хотя записано было то, что просили.
 * Такое превращение — не расхождение; число при этом должно совпасть.
 */
function typedLiteral(expected: string, actual: unknown): boolean {
  const text = expected.trim();
  if (typeof actual === "boolean") return /^(true|false|истина|ложь)$/i.test(text) && /^(true|истина)$/i.test(text) === actual;
  if (typeof actual !== "number") return false;
  // \s покрывает и неразрывный пробел — разделитель тысяч в русской записи.
  const number = /^[+-]?\d[\d\s]*(?:[.,]\d+)?(%?)$/.exec(text);
  if (number) {
    const parsed = Number(text.replace(/[\s%]/g, "").replace(",", ".")) / (number[1] ? 100 : 1);
    return Math.abs(parsed - actual) <= 1e-9 * Math.max(1, Math.abs(actual));
  }
  // Дата или время: серийный номер Excel проверить по тексту без локали нельзя, поэтому достаточно числа.
  return /^\d{1,4}[./-]\d{1,2}[./-]\d{1,4}(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?$|^\d{1,2}:\d{2}(?::\d{2})?$/.test(text);
}

export function sameCellMatrix(actual: readonly (readonly unknown[])[], expected: readonly (readonly unknown[])[], typed = false): boolean {
  if (actual.length !== expected.length) return false;
  return expected.every((row, r) =>
    (actual[r]?.length ?? -1) === row.length && row.every((cell, c) => sameCellContent(actual[r][c], cell, typed)));
}
