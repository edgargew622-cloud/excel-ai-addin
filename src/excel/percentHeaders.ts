/**
 * Доли без процентного формата под шапкой «… %» (26.10.2026, «Книга1567»).
 *
 * Агент оформил проценты у первого блока листа и забыл у второго: под
 * шапкой «Доля, %» стояло 0,0789… вместо 7,9 %. Уборка в конце задачи
 * находит такие ячейки сама. Правило узкое:
 *
 * - шапка — текст со знаком «%» («Доля, %», «Темп, %», «Доля в выручке, %»);
 * - ниже неё подряд, до первой пустой или текстовой ячейки, — формулы
 *   с числом в формате «Общий» и по модулю не больше 10;
 * - числа, введённые руками, не трогаются: это могут быть данные
 *   пользователя в своих единицах.
 */
const PERCENT_HEADER = /%/;
const MAX_SHARE = 10;

export function percentCellsUnderHeaders(
  values: readonly (readonly unknown[])[],
  formulas: readonly (readonly unknown[])[],
  formats: readonly (readonly unknown[])[]
): [number, number][] {
  const found: [number, number][] = [];
  values.forEach((line, r) => line.forEach((header, c) => {
    if (typeof header !== "string" || !PERCENT_HEADER.test(header) || String(formulas[r]?.[c] ?? "").startsWith("=")) return;
    for (let row = r + 1; row < values.length; row++) {
      const value = values[row]?.[c];
      if (typeof value !== "number") break;
      const formula = String(formulas[row]?.[c] ?? "");
      const format = String(formats[row]?.[c] ?? "");
      if (formula.startsWith("=") && /^(General|Общий)$/i.test(format) && Math.abs(value) <= MAX_SHARE) found.push([row, c]);
    }
  }));
  return found;
}
