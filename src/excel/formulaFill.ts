/**
 * Построение формул для заполнения области, когда протяжка Excel недоступна.
 *
 * Обычно область заполняет сам Excel через `Range.autoFill`: он подстраивает
 * относительные ссылки так же, как при протяжке за угол. Но проверка
 * 18 сентября 2026 года показала, что на записи в `Продажи!H2:H6` рядом
 * с таблицей `SalesTable` Excel дважды ответил внутренней ошибкой. Поэтому
 * нужен запасной путь: те же формулы строятся здесь и пишутся обычной записью.
 *
 * Правила подстановки — как у Excel. Относительная часть ссылки сдвигается
 * на смещение ячейки от первой, часть со знаком доллара остаётся на месте.
 */

const MAX_COLUMN = 16_384;
const MAX_ROW = 1_048_576;

export function columnLetters(index: number): string {
  let value = "";
  let left = index;
  while (left > 0) {
    const remainder = (left - 1) % 26;
    value = String.fromCharCode(65 + remainder) + value;
    left = Math.floor((left - 1) / 26);
  }
  return value;
}

export function columnIndex(letters: string): number {
  let index = 0;
  for (const letter of letters.toUpperCase()) index = index * 26 + (letter.charCodeAt(0) - 64);
  return index;
}

/** Ссылка на ячейку: буквы столбца и номер строки, каждая часть со своим $. */
const REFERENCE = /(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/g;

/**
 * Сдвигает ссылки формулы на заданное смещение.
 *
 * Текст в кавычках не трогается: в нём могут быть похожие на ссылки строки.
 * Имена функций тоже: `LOG10(` выглядит как ссылка `LOG10`, поэтому совпадение
 * отбрасывается, если сразу за ним идёт открывающая скобка или если перед ним
 * стоит буква, цифра или подчёркивание — часть более длинного имени.
 */
export function shiftFormula(formula: string, rowDelta: number, columnDelta: number): string {
  if (!formula.startsWith("=")) return formula;
  let result = "";
  let index = 0;
  let inQuotes = false;

  while (index < formula.length) {
    const char = formula[index];
    if (char === '"') {
      inQuotes = !inQuotes;
      result += char;
      index += 1;
      continue;
    }
    if (inQuotes) {
      result += char;
      index += 1;
      continue;
    }

    REFERENCE.lastIndex = index;
    const match = REFERENCE.exec(formula);
    if (!match || match.index !== index) {
      result += char;
      index += 1;
      continue;
    }

    const [whole, columnAbsolute, letters, rowAbsolute, digits] = match;
    const previous = index > 0 ? formula[index - 1] : "";
    const next = formula[index + whole.length] ?? "";
    const partOfName = /[A-Za-z0-9_.]/.test(previous);
    const functionCall = next === "(";
    if (partOfName || functionCall) {
      result += whole;
      index += whole.length;
      continue;
    }

    const column = columnAbsolute ? columnIndex(letters) : columnIndex(letters) + columnDelta;
    const row = rowAbsolute ? Number(digits) : Number(digits) + rowDelta;
    // Выход за границы листа Excel показывает как #ССЫЛКА!; строим то же самое.
    const shifted = column < 1 || column > MAX_COLUMN || row < 1 || row > MAX_ROW
      ? "#REF!"
      : `${columnAbsolute}${columnLetters(column)}${rowAbsolute}${digits === "" ? "" : row}`;
    result += shifted;
    index += whole.length;
  }
  return result;
}

/** Матрица формул для всей области: первая ячейка как есть, остальные со сдвигом. */
export function fillFormulaMatrix(formula: string, rows: number, columns: number): string[][] {
  return Array.from({ length: rows }, (_, row) =>
    Array.from({ length: columns }, (_, column) => shiftFormula(formula, row, column))
  );
}
