/**
 * Есть ли функция в этом Excel и что значит ошибка в ячейке (этап 7,
 * пункты 7.1.4 и 7.1.5).
 *
 * Замер 24 сентября 2026 года на Office 2021 (16.0.14334):
 *
 * - ошибки Excel отдаёт на языке интерфейса — `#ИМЯ?`, `#ДЕЛ/0!`, `#Н/Д`;
 *   признак, не зависящий от языка, — `valueTypes = "Error"`;
 * - набор функций зависит от версии: XLOOKUP, FILTER, LET здесь есть,
 *   LAMBDA, TEXTSPLIT, VSTACK — дают `#ИМЯ?`; русские имена (`СУММ`) через
 *   API не понимаются вовсе;
 * - запись `=ИМЯ()` различает одно от другого без ссылок: известная функция
 *   либо отклоняется Excel (`InvalidArgument` — мало аргументов), либо даёт
 *   значение; неизвестная записывается и даёт `#ИМЯ?`.
 *
 * Поэтому функции проверяются именно так, во временной ячейке на пустом
 * месте вне данных: без ссылок формула не может зацепить саму ячейку
 * (циклическая ссылка вызвала бы окно Excel), а результат запоминается
 * до конца сеанса — повторно та же функция не проверяется.
 */

/** Ошибка Excel по смыслу: подписи зависят от языка интерфейса. */
export type ExcelErrorKind = "name" | "div0" | "na" | "value" | "ref" | "num" | "null" | "spill" | "calc" | "other";

const ERROR_TEXT: Record<Exclude<ExcelErrorKind, "other">, readonly string[]> = {
  name: ["#NAME?", "#ИМЯ?", "#NOM?", "#NOME?", "#NAZWA?", "#¿NOMBRE?"],
  div0: ["#DIV/0!", "#ДЕЛ/0!"],
  na: ["#N/A", "#Н/Д"],
  value: ["#VALUE!", "#ЗНАЧ!"],
  ref: ["#REF!", "#ССЫЛКА!"],
  num: ["#NUM!", "#ЧИСЛО!"],
  null: ["#NULL!", "#ПУСТО!"],
  spill: ["#SPILL!", "#ПЕРЕНОС!"],
  calc: ["#CALC!", "#ВЫЧИСЛ!"]
};

export const ERROR_MEANING: Record<ExcelErrorKind, string> = {
  name: "функция или имя, которых этот Excel не знает",
  div0: "деление на ноль в данных",
  na: "искомое значение не найдено",
  value: "аргумент не того типа, например текст вместо числа",
  ref: "ссылка на несуществующую ячейку",
  num: "недопустимое число, например корень из отрицательного",
  null: "пересечение диапазонов пусто",
  spill: "результату-массиву некуда разлиться: соседние ячейки заняты",
  calc: "вычисление массива не удалось",
  other: "ошибка Excel"
};

export function errorKind(text: unknown): ExcelErrorKind {
  const value = String(text ?? "").trim().toUpperCase();
  for (const [kind, texts] of Object.entries(ERROR_TEXT)) {
    if (texts.some((item) => item.toUpperCase() === value)) return kind as ExcelErrorKind;
  }
  return "other";
}

export interface ExcelErrorCell {
  cell: string;
  text: string;
  kind: ExcelErrorKind;
}

/**
 * Ячейки с ошибками. Где Excel отдал `valueTypes`, признак — тип `Error`:
 * он не зависит от языка. Без него — по известным подписям на русском
 * и английском.
 */
export function errorCells(
  values: readonly (readonly unknown[])[],
  valueTypes: readonly (readonly unknown[])[] | undefined,
  cellName: (row: number, column: number) => string
): ExcelErrorCell[] {
  const found: ExcelErrorCell[] = [];
  values.forEach((row, r) => row.forEach((value, c) => {
    const type = valueTypes?.[r]?.[c];
    const isError = type !== undefined ? type === "Error" : typeof value === "string" && errorKind(value) !== "other";
    if (isError) found.push({ cell: cellName(r, c), text: String(value), kind: errorKind(value) });
  }));
  return found;
}

/** Пояснение к новым ошибкам для ответа инструмента: что значит каждая. */
export function errorNote(errors: readonly ExcelErrorCell[]): string | undefined {
  if (!errors.length) return undefined;
  const kinds = [...new Set(errors.map((item) => item.kind))];
  const meaning = kinds.map((kind) => `${errors.find((item) => item.kind === kind)!.text} — ${ERROR_MEANING[kind]}`).join("; ");
  return `После записи Excel показывает ошибки в ${errors.length} ячейках: ${meaning}. ` +
    "Запись выполнена как просили, но результат — ошибка: обязательно скажи это пользователю, назови ячейки и причину" +
    (kinds.includes("name") ? "; для неизвестной функции предложи замену явно и дождись согласия, не заменяй молча." : ".");
}

/* --- функции формулы -------------------------------------------------------- */

/**
 * Имена функций в формуле, заглавными: всё, за чем идёт открывающая скобка.
 * Текст в кавычках и имена листов в апострофах пропускаются — «(» внутри
 * них функцией не является.
 */
export function functionNamesIn(formula: unknown): string[] {
  if (typeof formula !== "string" || !formula.startsWith("=")) return [];
  const bare = formula.replace(/"(?:[^"]|"")*"/g, '""').replace(/'(?:[^']|'')*'!/g, "X!");
  const names = new Set<string>();
  // Просмотр назад не поглощает символ: иначе скобка внешней функции
  // съедалась бы, и вложенная — IFERROR(XLOOKUP(…)) — терялась.
  const pattern = /(?<![A-Za-z0-9_.\u0400-\u04FF])([A-Za-z_\u0400-\u04FF][A-Za-z0-9_.\u0400-\u04FF]*)\s*\(/g;
  for (let match = pattern.exec(bare); match; match = pattern.exec(bare)) names.add(match[1].toUpperCase());
  return [...names];
}

/** Что уже известно о функциях в этом сеансе Excel: true — есть, false — нет. */
const availability = new Map<string, boolean>();

export function resetFunctionAvailability(): void {
  availability.clear();
}

export function knownAvailability(name: string): boolean | undefined {
  return availability.get(name.toUpperCase());
}

/** Имена из формул, которых в этом Excel заведомо нет — проверено раньше. */
export function knownMissing(formulas: readonly unknown[]): string[] {
  return [...new Set(formulas.flatMap(functionNamesIn))].filter((name) => availability.get(name) === false);
}

/** Имена из формул, о которых ещё ничего не известно. */
export function uncheckedFunctions(formulas: readonly unknown[]): string[] {
  return [...new Set(formulas.flatMap(functionNamesIn))].filter((name) => !availability.has(name));
}

export function missingFunctionsMessage(names: readonly string[]): string {
  return `В этом Excel нет функций ${names.join(", ")}: формула с ними дала бы #ИМЯ?. Запись не выполнялась. ` +
    "Имена функций пишутся по-английски; если функция есть только в новых версиях Excel, предложи пользователю замену явно " +
    "(например, INDEX и MATCH вместо XLOOKUP) и дождись согласия.";
}

/** Ячейка для проверки: через столбец правее и данных листа, и цели, в строке начала данных. */
export function probeCellFor(
  used: { rowIndex: number; columnIndex: number; columnCount: number } | null,
  target: { rowIndex: number; columnIndex: number; columnCount: number },
  columnName: (index: number) => string
): string | null {
  const right = Math.max(used ? used.columnIndex + used.columnCount : 0, target.columnIndex + target.columnCount);
  const column = right + 1;
  if (column >= 16_384) return null;
  return `${columnName(column + 1)}${(used ? used.rowIndex : target.rowIndex) + 1}`;
}

export interface FunctionCheck {
  names: string[];
  /** Пустая ячейка вне данных; null — места нет, функции не проверяются. */
  probeCell: string | null;
  note: string;
}

export interface ProbeResult {
  available: string[];
  unavailable: string[];
  /** Ответ Excel не удалось истолковать — о функции ничего не утверждается. */
  unchecked: string[];
}

/**
 * Проверяет функции в пустой ячейке: пишет `=ИМЯ()`, читает, очищает.
 * Вызывается только внутри исполнения подтверждённого плана.
 */
export async function probeFunctions(ctx: Excel.RequestContext, sheet: Excel.Worksheet, cell: string, names: readonly string[]): Promise<ProbeResult> {
  const result: ProbeResult = { available: [], unavailable: [], unchecked: [] };
  const target = sheet.getRange(cell);
  for (const name of names) {
    try {
      target.formulas = [[`=${name}()`]] as any[][];
      await ctx.sync();
      target.load(["values", "valueTypes"]);
      await ctx.sync();
      const value = (target.values as unknown[][])[0][0];
      const type = (target.valueTypes as unknown[][])[0][0];
      if (type === "Error" && errorKind(value) === "name") result.unavailable.push(name);
      else if (type === "Error" && errorKind(value) === "other") result.unchecked.push(name);
      else result.available.push(name);
    } catch (error: any) {
      // Мало аргументов у существующей функции — Excel не даёт её записать.
      if (error?.code === "InvalidArgument") result.available.push(name);
      else result.unchecked.push(name);
    }
  }
  target.clear("Contents" as any);
  target.load("formulas");
  await ctx.sync();
  const left = (target.formulas as unknown[][])[0][0];
  if (left !== "" && left !== null && left !== undefined) {
    throw new Error(`временная ячейка ${cell} не очистилась: в ней осталось ${String(left)}`);
  }
  for (const name of result.available) availability.set(name, true);
  for (const name of result.unavailable) availability.set(name, false);
  return result;
}
