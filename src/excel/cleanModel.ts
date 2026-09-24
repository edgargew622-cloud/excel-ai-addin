/**
 * Что станет с каждой ячейкой при очистке (этап 7, 7.2.2–7.2.3).
 *
 * Модуль не обращается к Excel. Для каждой ячейки он решает: изменить —
 * и на что, или пропустить — и почему. Неоднозначное без явного ответа
 * пользователя не меняется: `01.02.2026` без порядка даты и `1,500` без
 * десятичного разделителя пропускаются с причиной, а не угадываются.
 */

import { classifyNumberText, parseDateText, parseNumberText, type DateOrder, type NumberCulture } from "./dataProfile";

/** Почему ячейка не изменится — для предпросмотра и ответа. */
export type SkipReason =
  | "formula"
  | "notText"
  | "unchanged"
  | "notNumber"
  | "ambiguousNumber"
  | "foreignNumber"
  | "leadingZeros"
  | "notDate"
  | "ambiguousDate"
  | "contradictsOrder";

export const SKIP_TEXT: Record<SkipReason, string> = {
  formula: "формула — меняется в самой формуле, а не в ячейке",
  notText: "уже не текст",
  unchanged: "менять нечего",
  notNumber: "не число",
  ambiguousNumber: "неоднозначно: дробь или тысячи (1,500) — нужен десятичный разделитель",
  foreignNumber: "разделитель не как в книге (1.5) — нужен явный десятичный разделитель",
  leadingZeros: "код с ведущими нулями — числом стал бы без нулей",
  notDate: "не дата",
  ambiguousDate: "неоднозначно: день и месяц не различить (01.02.2026) — нужен порядок даты",
  contradictsOrder: "противоречит названному порядку даты"
};

const OTHER_SPACES = /[\u00a0\u202f\u2007]/g;

/**
 * Текст без лишних пробелов: неразрывные пробелы становятся обычными,
 * края срезаются, при collapseInner повторы внутри сжимаются до одного —
 * как у функции Excel TRIM.
 */
export function cleanText(text: string, collapseInner: boolean): string {
  const unified = text.replace(OTHER_SPACES, " ");
  const trimmed = unified.replace(/^ +| +$/g, "");
  return collapseInner ? trimmed.replace(/ {2,}/g, " ") : trimmed;
}

/** Число из текста: по разделителям книги или по названному пользователем. */
export function numberFromText(
  text: string,
  culture: NumberCulture,
  decimalSeparator?: "," | "."
): { value: number } | { skip: SkipReason } {
  const kind = classifyNumberText(text, culture);
  if (kind?.kind === "leadingZeros") return { skip: "leadingZeros" };
  if (decimalSeparator) {
    const group = decimalSeparator === "," ? " " : ",";
    const value = parseNumberText(text, decimalSeparator, group);
    return value === null ? { skip: "notNumber" } : { value };
  }
  if (!kind) return { skip: "notNumber" };
  if (kind.kind === "number") return { value: kind.value! };
  return { skip: kind.kind === "ambiguousNumber" ? "ambiguousNumber" : "foreignNumber" };
}

export interface CalendarDate {
  year: number;
  month: number;
  day: number;
}

/** Дата из текста: однозначная — всегда, неоднозначная — только по названному порядку. */
export function dateFromText(text: string, order?: DateOrder): { date: CalendarDate } | { skip: SkipReason } {
  const info = parseDateText(text);
  if (!info) return { skip: "notDate" };
  const yearFirst = info.orders.includes("YMD");
  if (info.date) {
    // Порядок однозначен сам по себе; названный порядок не должен ему противоречить.
    if (order && !yearFirst && info.orders.length === 1 && info.orders[0] !== order) return { skip: "contradictsOrder" };
    return { date: info.date };
  }
  if (!order) return { skip: "ambiguousDate" };
  const match = /^(\d{1,2})[-./](\d{1,2})[-./](\d{4})$/.exec(text.trim());
  if (!match || order === "YMD") return { skip: "contradictsOrder" };
  const [a, b, year] = [Number(match[1]), Number(match[2]), Number(match[3])];
  return { date: order === "DMY" ? { year, month: b, day: a } : { year, month: a, day: b } };
}

/**
 * Порядковый номер даты в Excel (система 1900: 1 января 1900 — 1).
 * Отсчёт от 30 декабря 1899 учитывает несуществующее 29 февраля 1900,
 * которое Excel считает датой, — для дат после марта 1900 это верно.
 * Книгу в системе 1904 выдаст сверка по отображению после записи.
 */
export function excelSerial(date: CalendarDate): number {
  return Math.round((Date.UTC(date.year, date.month - 1, date.day) - Date.UTC(1899, 11, 30)) / 86_400_000);
}

/**
 * Код формата даты по шаблону книги: «ДД.ММ.ГГГГ» → «dd.mm.yyyy».
 * Коды форматов API принимает только английские — замер 24 сентября 2026 года.
 */
export function dateFormatCode(pattern: string): string {
  const code = pattern
    .replace(/Д/g, "d").replace(/д/g, "d")
    .replace(/М/g, "m").replace(/м/g, "m")
    .replace(/Г/g, "y").replace(/г/g, "y")
    .replace(/D/g, "d").replace(/M/g, "m").replace(/Y/g, "y");
  return /d/.test(code) && /m/.test(code) && /y/.test(code) ? code : "yyyy-mm-dd";
}

/** Как Excel покажет дату в этом формате — для сверки после записи. */
export function formatDate(date: CalendarDate, code: string): string {
  const pad = (value: number, width: number) => String(value).padStart(width, "0");
  return code.replace(/yyyy|yy|mm|m|dd|d/g, (token) => {
    switch (token) {
      case "yyyy": return pad(date.year, 4);
      case "yy": return pad(date.year % 100, 2);
      case "mm": return pad(date.month, 2);
      case "m": return String(date.month);
      case "dd": return pad(date.day, 2);
      default: return String(date.day);
    }
  });
}

/* --- дубликаты (7.2.4) ------------------------------------------------------------ */

/**
 * Ключ значения так, как его сравнивает Excel при удалении дубликатов.
 *
 * Замер 24 сентября 2026 года (`removeDuplicates`): регистр не различается
 * («москва» — повтор «Москва»), пробел в конце различается («Москва »
 * осталась), текст «1» и число 1 — разные значения (в сводной они
 * сливались — у сводной свои правила). Пустые ячейки равны между собой.
 */
export function duplicateKey(value: unknown): string {
  if (value === "" || value === null || value === undefined) return "e";
  if (typeof value === "number") return `n:${value}`;
  if (typeof value === "boolean") return `b:${value}`;
  return `s:${String(value).toLowerCase()}`;
}

export interface DuplicatePlan {
  /** Индексы строк данных (от 0 — первая строка после шапки), которые останутся. */
  keep: number[];
  /** Удаляемые строки и строка, повтором которой каждая является. */
  removed: { row: number; duplicateOf: number }[];
  /** Какими станут строки данных области: оставшиеся по порядку, ниже — пустые. */
  expected: unknown[][];
}

/** Что сделает удаление дубликатов: остаётся первое вхождение, остальные уходят, строки ниже поднимаются. */
export function planDuplicates(rows: readonly (readonly unknown[])[], keyColumns: readonly number[]): DuplicatePlan {
  const first = new Map<string, number>();
  const keep: number[] = [];
  const removed: { row: number; duplicateOf: number }[] = [];
  rows.forEach((row, index) => {
    const key = JSON.stringify(keyColumns.map((column) => duplicateKey(row[column])));
    const seen = first.get(key);
    if (seen === undefined) {
      first.set(key, index);
      keep.push(index);
    } else {
      removed.push({ row: index, duplicateOf: seen });
    }
  });
  const width = rows[0]?.length ?? 0;
  const expected = [
    ...keep.map((index) => [...rows[index]]),
    ...removed.map(() => Array.from({ length: width }, () => ""))
  ];
  return { keep, removed, expected };
}
