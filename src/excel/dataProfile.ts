/**
 * Профиль данных: что в таблице мешает считать (этап 7, 7.2.1).
 *
 * Только чтение. Панель считает сама по прочитанным значениям и называет
 * проверенную область; содержимое таблицы в ответ не выгружается — только
 * счётчики и несколько примеров на каждую находку.
 *
 * Замер в Excel 24 сентября 2026 года (ru-RU):
 * - число или дата, записанные текстом, остаются текстом (`valueTypes =
 *   String`), в том числе с пробелом или неразрывным пробелом внутри —
 *   в расчёт они не входят, и это главное, что ищет профиль;
 * - разделители книги Excel сообщает сам (`application.cultureInfo`):
 *   десятичная запятая, пробел между тысячами, дата `ДД.ММ.ГГГГ`;
 * - коды числовых форматов API отдаёт по-английски (`dd.mm.yyyy`).
 *
 * Правило: неоднозначное не угадывается. `01.02.2026` может быть и первым
 * февраля, и вторым января; `1,500` — и полутора, и тысячей пятьсот. Такие
 * значения считаются отдельно, и преобразовывать их можно только после
 * явного ответа пользователя.
 */

export interface NumberCulture {
  /** Десятичный разделитель книги, например «,». */
  decimal: string;
  /** Разделитель тысяч, например « ». */
  group: string;
}

export type DateOrder = "DMY" | "MDY" | "YMD";

const SPACES = /[ \u00a0\u202f\u2007]/;
const SPACE_CLASS = "\\u0020\\u00a0\\u202f\\u2007";
const escape = (text: string) => text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

/** Классы разделителя тысяч: пробел книги означает любой вид пробела. */
function groupClass(group: string): string {
  return SPACES.test(group) || group === "" ? `[${SPACE_CLASS}]` : escape(group);
}

/**
 * Текст как число по заданным разделителям — строго: тысячи группами по три,
 * без валюты, процентов и экспоненты. Иначе null.
 */
export function parseNumberText(text: string, decimal: string, group: string): number | null {
  const trimmed = text.replace(new RegExp(`^[${SPACE_CLASS}]+|[${SPACE_CLASS}]+$`, "g"), "");
  const g = groupClass(group);
  const d = escape(decimal);
  const pattern = new RegExp(`^([+-]?)(\\d{1,3}(?:${g}\\d{3})+|\\d+)(?:${d}(\\d+))?$`);
  const match = pattern.exec(trimmed);
  if (!match) return null;
  const whole = match[2].replace(new RegExp(g, "g"), "");
  const value = Number(`${match[1]}${whole}${match[3] ? `.${match[3]}` : ""}`);
  return Number.isFinite(value) ? value : null;
}

export type NumberTextKind =
  /** Число по разделителям книги — однозначно. */
  | "number"
  /** Число только по чужим разделителям (например «1.5» в русской книге). */
  | "foreignNumber"
  /** Читается и как дробь, и как тысячи: «1,500», «2.300». */
  | "ambiguousNumber"
  /** Ведущие нули — это код, а не число: «007». */
  | "leadingZeros";

export function classifyNumberText(text: string, culture: NumberCulture): { kind: NumberTextKind; value?: number } | null {
  const trimmed = text.trim();
  if (/^[+-]?0\d/.test(trimmed)) return parseNumberText(trimmed, culture.decimal, culture.group) !== null ? { kind: "leadingZeros" } : null;
  // Один разделитель и ровно три цифры после него — дробь или тысячи?
  if (/^[+-]?\d{1,3}[.,]\d{3}$/.test(trimmed)) return { kind: "ambiguousNumber" };
  const own = parseNumberText(trimmed, culture.decimal, culture.group);
  if (own !== null) return { kind: "number", value: own };
  const otherDecimal = culture.decimal === "," ? "." : ",";
  const foreign = parseNumberText(trimmed, otherDecimal, otherDecimal === "." ? "," : " ");
  if (foreign !== null) return { kind: "foreignNumber", value: foreign };
  return null;
}

export interface DateTextInfo {
  /** Однозначная дата: год, месяц, день. */
  date?: { year: number; month: number; day: number };
  /** Какие порядки подходят; больше одного — неоднозначно. */
  orders: DateOrder[];
}

const validDate = (year: number, month: number, day: number) => {
  if (month < 1 || month > 12 || day < 1) return false;
  const days = new Date(Date.UTC(year, month, 0)).getUTCDate();
  return day <= days;
};

/**
 * Текст как дата. Четырёхзначный год обязателен: двузначный не говорит
 * о веке, и такой текст датой не считается. Порядок «день-месяц» против
 * «месяц-день» определяется, только если одно из чисел больше 12.
 */
export function parseDateText(text: string): DateTextInfo | null {
  const trimmed = text.trim();
  const ymd = /^(\d{4})([-./])(\d{1,2})\2(\d{1,2})$/.exec(trimmed);
  if (ymd) {
    const [year, month, day] = [Number(ymd[1]), Number(ymd[3]), Number(ymd[4])];
    return validDate(year, month, day) ? { date: { year, month, day }, orders: ["YMD"] } : null;
  }
  const other = /^(\d{1,2})([-./])(\d{1,2})\2(\d{4})$/.exec(trimmed);
  if (!other) return null;
  const [a, b, year] = [Number(other[1]), Number(other[3]), Number(other[4])];
  const orders: DateOrder[] = [];
  if (validDate(year, b, a)) orders.push("DMY");
  if (validDate(year, a, b)) orders.push("MDY");
  if (!orders.length) return null;
  if (orders.length === 1) {
    const [month, day] = orders[0] === "DMY" ? [b, a] : [a, b];
    return { date: { year, month, day }, orders };
  }
  // Оба порядка возможны. Если числа равны (05.05.2026), порядок неважен.
  return a === b ? { date: { year, month: a, day: a }, orders } : { orders };
}

/** Дата по заданному порядку — когда пользователь его назвал. */
export function dateByOrder(text: string, order: DateOrder): { year: number; month: number; day: number } | null {
  const info = parseDateText(text);
  if (!info) return null;
  // Год впереди — порядок однозначен сам по себе.
  if (info.orders.includes("YMD")) return info.date ?? null;
  if (!info.orders.includes(order)) return null;
  const match = /^(\d{1,2})[-./](\d{1,2})[-./](\d{4})$/.exec(text.trim());
  if (!match) return null;
  const [a, b, year] = [Number(match[1]), Number(match[2]), Number(match[3])];
  return order === "DMY" ? { year, month: b, day: a } : { year, month: a, day: b };
}

/** Похож ли код формата на дату (коды форматов API — английские). */
export function isDateFormat(format: unknown): boolean {
  const text = String(format ?? "").replace(/"[^"]*"|\[[^\]]*\]/g, "");
  return /[dy]/i.test(text) || /m{3,}/i.test(text);
}

/* --- профиль ------------------------------------------------------------------ */

const MAX_EXAMPLES = 3;

export interface Finding {
  count: number;
  /** Адреса и значения первых случаев — для отчёта, не вся таблица. */
  examples: string[];
}

const finding = (): Finding => ({ count: 0, examples: [] });
const note = (item: Finding, example: string) => {
  item.count += 1;
  if (item.examples.length < MAX_EXAMPLES) item.examples.push(example);
};

export interface ColumnProfile {
  column: string;
  header: string | null;
  cells: number;
  empty: number;
  numbers: number;
  dates: number;
  texts: number;
  booleans: number;
  errors: number;
  formulas: number;
  distinct: number;
  /** Пробелы по краям текста. */
  edgeSpaces: Finding;
  /** Два и больше пробела подряд внутри текста. */
  innerSpaces: Finding;
  /** Неразрывный пробел в тексте — невидим и мешает сравнению. */
  nonBreakingSpaces: Finding;
  numbersAsText: Finding;
  foreignNumbersAsText: Finding;
  ambiguousNumbersAsText: Finding;
  codesWithLeadingZeros: Finding;
  datesAsText: Finding;
  ambiguousDatesAsText: Finding;
}

export interface DataProfile {
  address: string;
  hasHeaders: boolean;
  rows: number;
  columns: ColumnProfile[];
  emptyRows: number;
  /** Строки, полностью совпавшие с одной из выше. */
  duplicateRows: Finding;
  /** Совпавшие после обрезки пробелов и без учёта регистра — кандидаты. */
  duplicateRowsNormalized: Finding;
  culture: NumberCulture & { dateOrder: DateOrder | null };
}

const isBlank = (value: unknown) => value === "" || value === null || value === undefined;
const normalizeText = (value: unknown) =>
  typeof value === "string" ? value.replace(new RegExp(`[${SPACE_CLASS}]+`, "g"), " ").trim().toLowerCase() : value;

/** Порядок даты книги по её шаблону, например «ДД.ММ.ГГГГ» → DMY. */
export function cultureDateOrder(pattern: string): DateOrder | null {
  const text = pattern.toUpperCase().replace(/Д/g, "D").replace(/М/g, "M").replace(/Г/g, "Y");
  const order = [...text].filter((ch) => ch === "D" || ch === "M" || ch === "Y");
  const first = [...new Set(order)].join("");
  return first === "DMY" || first === "MDY" || first === "YMD" ? first : null;
}

export function profileData(input: {
  values: readonly (readonly unknown[])[];
  formulas: readonly (readonly unknown[])[];
  valueTypes: readonly (readonly unknown[])[];
  numberFormat: readonly (readonly unknown[])[];
  hasHeaders: boolean;
  origin: { rowIndex: number; columnIndex: number };
  address: string;
  culture: NumberCulture & { dateOrder: DateOrder | null };
  columnName: (index: number) => string;
}): DataProfile {
  const { values, formulas, valueTypes, numberFormat, hasHeaders, origin, culture, columnName } = input;
  const width = values[0]?.length ?? 0;
  const bodyStart = hasHeaders ? 1 : 0;
  const cellName = (r: number, c: number) => `${columnName(origin.columnIndex + c + 1)}${origin.rowIndex + r + 1}`;

  const columns: ColumnProfile[] = Array.from({ length: width }, (_, c) => ({
    column: columnName(origin.columnIndex + c + 1),
    header: hasHeaders && !isBlank(values[0]?.[c]) ? String(values[0][c]) : null,
    cells: 0, empty: 0, numbers: 0, dates: 0, texts: 0, booleans: 0, errors: 0, formulas: 0, distinct: 0,
    edgeSpaces: finding(), innerSpaces: finding(), nonBreakingSpaces: finding(),
    numbersAsText: finding(), foreignNumbersAsText: finding(), ambiguousNumbersAsText: finding(), codesWithLeadingZeros: finding(),
    datesAsText: finding(), ambiguousDatesAsText: finding()
  }));
  const distinct = columns.map(() => new Set<string>());

  let emptyRows = 0;
  const duplicateRows = finding();
  const duplicateRowsNormalized = finding();
  const seenExact = new Map<string, number>();
  const seenNormalized = new Map<string, number>();

  for (let r = bodyStart; r < values.length; r++) {
    const row = values[r];
    if (row.every(isBlank)) { emptyRows += 1; continue; }
    const exact = JSON.stringify(row);
    const normalized = JSON.stringify(row.map(normalizeText));
    const rowName = `строка ${origin.rowIndex + r + 1}`;
    if (seenExact.has(exact)) note(duplicateRows, `${rowName} = строка ${seenExact.get(exact)}`);
    else seenExact.set(exact, origin.rowIndex + r + 1);
    if (seenNormalized.has(normalized)) note(duplicateRowsNormalized, `${rowName} ≈ строка ${seenNormalized.get(normalized)}`);
    else seenNormalized.set(normalized, origin.rowIndex + r + 1);

    row.forEach((value, c) => {
      const column = columns[c];
      column.cells += 1;
      const type = valueTypes[r]?.[c];
      const formula = formulas[r]?.[c];
      if (typeof formula === "string" && formula.startsWith("=")) column.formulas += 1;
      if (isBlank(value) || type === "Empty") { column.empty += 1; return; }
      distinct[c].add(JSON.stringify(value));
      if (type === "Error") { column.errors += 1; return; }
      if (type === "Boolean" || typeof value === "boolean") { column.booleans += 1; return; }
      if (typeof value === "number") {
        if (isDateFormat(numberFormat[r]?.[c])) column.dates += 1;
        else column.numbers += 1;
        return;
      }
      column.texts += 1;
      // Текстовые находки — только у введённого текста: у формулы текст
      // результата исправляется в формуле, а не в ячейке.
      if (typeof formula === "string" && formula.startsWith("=")) return;
      const text = String(value);
      const where = `${cellName(r, c)}: «${text}»`;
      if (text !== text.replace(new RegExp(`^[${SPACE_CLASS}]+|[${SPACE_CLASS}]+$`, "g"), "")) note(column.edgeSpaces, where);
      if (/\S[ \u00a0\u202f]{2,}\S/.test(text)) note(column.innerSpaces, where);
      if (/[\u00a0\u202f\u2007]/.test(text)) note(column.nonBreakingSpaces, where);
      const numberKind = classifyNumberText(text, culture);
      if (numberKind?.kind === "number") note(column.numbersAsText, where);
      else if (numberKind?.kind === "foreignNumber") note(column.foreignNumbersAsText, where);
      else if (numberKind?.kind === "ambiguousNumber") note(column.ambiguousNumbersAsText, where);
      else if (numberKind?.kind === "leadingZeros") note(column.codesWithLeadingZeros, where);
      const date = parseDateText(text);
      if (date?.date) note(column.datesAsText, where);
      else if (date) note(column.ambiguousDatesAsText, where);
    });
  }
  columns.forEach((column, c) => { column.distinct = distinct[c].size; });

  return {
    address: input.address,
    hasHeaders,
    rows: values.length - bodyStart,
    columns,
    emptyRows,
    duplicateRows,
    duplicateRowsNormalized,
    culture
  };
}
