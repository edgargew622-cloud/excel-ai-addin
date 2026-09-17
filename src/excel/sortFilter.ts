/**
 * Проверяемая логика сортировки и фильтра без обращений к Excel.
 *
 * Сортировка переставляет существующие данные, поэтому «было → станет» здесь
 * не замена значений, а новый порядок тех же строк. Отсюда два независимых
 * свойства результата, которые проверяются по отдельности:
 * — строки сохранились целиком: ни одна не потеряна, не удвоена и не собрана
 *   из кусков разных строк;
 * — ключевой столбец действительно упорядочен.
 * Первое — жёсткая проверка: нарушение означает порчу данных. Второе —
 * мягкая: правила сравнения текста в Excel зависят от локали, и расхождение
 * с нашей оценкой не доказывает ошибку.
 */

type Cell = unknown;
type Row = readonly Cell[];

const isBlank = (value: Cell) => value === null || value === undefined || value === "";
const isErrorText = (value: Cell) => typeof value === "string" && /^#[A-Z0-9/!?]+[!?0-9A-Z]*$/.test(value);

/** Порядок типов при сортировке Excel по возрастанию: числа, текст, логические,
 * ошибки. Пустые ячейки всегда в конце, при любом направлении. */
function typeRank(value: Cell): number {
  if (typeof value === "number") return 0;
  if (typeof value === "string" && !isErrorText(value)) return 1;
  if (typeof value === "boolean") return 2;
  return 3;
}

const collator = new Intl.Collator("ru", { sensitivity: "base", numeric: false });

/** Сравнение значений ключа так, как сортирует Excel. Текст без учёта регистра. */
export function excelSortCompare(a: Cell, b: Cell, ascending = true): number {
  const blankA = isBlank(a);
  const blankB = isBlank(b);
  if (blankA || blankB) return blankA === blankB ? 0 : blankA ? 1 : -1;
  const rank = typeRank(a) - typeRank(b);
  let result: number;
  if (rank !== 0) result = rank;
  else if (typeof a === "number" && typeof b === "number") result = a - b;
  else if (typeof a === "boolean" && typeof b === "boolean") result = Number(a) - Number(b);
  else result = collator.compare(String(a), String(b));
  return ascending ? result : -result;
}

/** Ожидаемый порядок: устойчивая сортировка, как у Excel. Нужна предпросмотру. */
export function sortRowsLikeExcel(rows: readonly Row[], key: number, ascending = true): Row[] {
  return rows
    .map((row, index) => ({ row, index }))
    .sort((x, y) => excelSortCompare(x.row[key], y.row[key], ascending) || x.index - y.index)
    .map((item) => item.row);
}

export function isSortedLikeExcel(rows: readonly Row[], key: number, ascending = true): boolean {
  for (let i = 1; i < rows.length; i++) {
    if (excelSortCompare(rows[i - 1][key], rows[i][key], ascending) > 0) return false;
  }
  return true;
}

/** Строки сохранились как целые кортежи: тот же набор с теми же повторами.
 * Сравнение по строкам, а не по ячейкам — перемешанные столбцы дают те же
 * ячейки, но другие строки, и именно это должно ловиться. */
export function sameRowMultiset(before: readonly Row[], after: readonly Row[]): boolean {
  if (before.length !== after.length) return false;
  const counts = new Map<string, number>();
  for (const row of before) {
    const key = JSON.stringify(row);
    counts.set(key, (counts.get(key) ?? 0) + 1);
  }
  for (const row of after) {
    const key = JSON.stringify(row);
    const left = counts.get(key);
    if (!left) return false;
    counts.set(key, left - 1);
  }
  return true;
}

/**
 * Проверяет, не режет ли сортировка строки соседних данных.
 *
 * Если отсортировать только часть столбцов сплошной области, Excel переставит
 * их, а соседние столбцы останутся на месте — строки перемешаются молча,
 * без ошибки. Возвращает описание проблемы или null.
 */
export function partialRowSortProblem(
  range: { columnStart: number; columnEnd: number; rowStart: number; rowEnd: number },
  region: { columnStart: number; columnEnd: number; rowStart: number; rowEnd: number } | null
): string | null {
  if (!region) return null;
  const overlapsRows = region.rowStart <= range.rowEnd && region.rowEnd >= range.rowStart;
  const wider = region.columnStart < range.columnStart || region.columnEnd > range.columnEnd;
  return overlapsRows && wider ? "область уже сплошного блока данных" : null;
}

export type FilterOn = "values" | "custom";

export interface ParsedFilterCriteria {
  filterOn: FilterOn;
  values?: string[];
  criterion1?: string;
}

/** Разбор условия фильтра: значения через «|», сравнение или одно значение. */
export function parseFilterCriteria(raw: string): ParsedFilterCriteria {
  const text = String(raw ?? "").trim();
  if (!text) throw new Error("Условие фильтра не может быть пустым.");
  if (text.includes("|")) {
    const values = text.split("|").map((item) => item.trim()).filter(Boolean);
    if (!values.length) throw new Error("В условии через «|» нет ни одного значения.");
    return { filterOn: "values", values };
  }
  if (/^(>=|<=|<>|>|<|=)/.test(text)) return { filterOn: "custom", criterion1: text };
  return { filterOn: "values", values: [text] };
}

/** Описание уже стоящего автофильтра для предпросмотра и сравнения. */
export interface AutoFilterState {
  enabled: boolean;
  address: string | null;
  /** Число столбцов с настоящим условием. */
  activeColumns: number;
  /** Номера столбцов внутри области фильтра, где условие есть. */
  activeIndexes: number[];
  criteria: string;
}

export function sameAutoFilterState(a: AutoFilterState, b: AutoFilterState): boolean {
  return a.enabled === b.enabled &&
    (a.address ?? "").toUpperCase() === (b.address ?? "").toUpperCase() &&
    a.criteria === b.criteria;
}

/**
 * Есть ли в условии содержимое.
 *
 * Проверка в Excel 17 сентября 2026 года: для столбцов без фильтра Excel отдаёт
 * заготовку условия с первым значением типа — `BottomItems` — и больше ничем.
 * Такая заготовка ничего не скрывает, и считать её условием нельзя: иначе
 * предпросмотр сообщал о прежнем фильтре «в трёх столбцах», когда были видны
 * все строки. Настоящее условие несёт значения, порог, сравнение, цвет или значок.
 */
export function hasCondition(item: any): boolean {
  if (!item || typeof item !== "object" || !item.filterOn) return false;
  const text = (value: unknown) => typeof value === "string" && value.trim() !== "";
  // Повторная проверка в Excel: столбец без условия всё равно считался активным.
  // Вероятная причина — непустые значения по умолчанию в заготовке: динамический
  // фильтр «Unknown» и значок с набором «Invalid». Они ничего не отбирают.
  const dynamic = text(item.dynamicCriteria) && String(item.dynamicCriteria).toLowerCase() !== "unknown";
  const icon = item.icon && typeof item.icon === "object" && text(item.icon.set) &&
    String(item.icon.set).toLowerCase() !== "invalid";
  return (Array.isArray(item.values) && item.values.length > 0) ||
    text(item.criterion1) ||
    text(item.criterion2) ||
    dynamic ||
    text(item.color) ||
    Boolean(icon);
}

/** Условие Excel содержит служебные поля; для сравнения нужны только значимые. */
export function describeCriteria(criteria: readonly unknown[] | null | undefined): {
  activeColumns: number;
  activeIndexes: number[];
  text: string;
} {
  const list = Array.isArray(criteria) ? criteria : [];
  const activeIndexes: number[] = [];
  const meaningful = list.map((item: any, index) => {
    if (!hasCondition(item)) return null;
    activeIndexes.push(index);
    return {
      filterOn: item.filterOn,
      ...(Array.isArray(item.values) && item.values.length ? { values: item.values } : {}),
      ...(item.criterion1 ? { criterion1: item.criterion1 } : {}),
      ...(item.criterion2 ? { criterion2: item.criterion2 } : {}),
      ...(item.dynamicCriteria && String(item.dynamicCriteria).toLowerCase() !== "unknown"
        ? { dynamicCriteria: item.dynamicCriteria }
        : {})
    };
  });
  return { activeColumns: activeIndexes.length, activeIndexes, text: JSON.stringify(meaningful) };
}

export type FilterChange = "new" | "adds" | "replacesColumn" | "replacesFilter";

function addressWithoutSheet(address: string | null): string {
  const text = address ?? "";
  return text.slice(text.lastIndexOf("!") + 1).split("$").join("").toUpperCase();
}

/**
 * Что сделает новый фильтр с уже стоящим.
 *
 * Проверка в Excel 17 сентября 2026 года: второй фильтр на ту же область по
 * другому столбцу сложился с первым — видимых строк осталось столько же, хотя
 * по новому условию прошли бы все. Заменяет прежний фильтр целиком только
 * фильтр на другую область листа; на той же области он добавляется, а в столбце,
 * где условие уже было, заменяет лишь это условие.
 */
export function filterChangeKind(before: AutoFilterState, targetAddress: string, column: number): FilterChange {
  if (!before.enabled || before.activeColumns === 0) return "new";
  if (addressWithoutSheet(before.address) !== addressWithoutSheet(targetAddress)) return "replacesFilter";
  return before.activeIndexes.includes(column) ? "replacesColumn" : "adds";
}

/**
 * Похожа ли первая строка на заголовки.
 *
 * Проверка в Excel 17 сентября 2026 года: признак искался только в ключевом
 * столбце, и при сортировке по текстовому «Городу» предупреждение промолчало,
 * хотя «Количество», «Цена» и «Сумма» ниже числовые. Смотреть нужно на всю
 * область: первая строка целиком текстовая, а хотя бы в одном столбце ниже
 * стоят числа или логические значения.
 */
export function firstRowLooksLikeHeader(values: readonly (readonly unknown[])[]): boolean {
  if (values.length < 2) return false;
  const first = values[0] ?? [];
  if (!first.length || !first.every((cell) => typeof cell === "string" && cell.trim() !== "")) return false;
  for (let column = 0; column < first.length; column++) {
    if (values.slice(1).some((row) => typeof row[column] === "number" || typeof row[column] === "boolean")) return true;
  }
  return false;
}

/** Условие в словах, пригодных для отчёта: без служебных полей Excel. */
export function conditionText(item: any): string {
  if (Array.isArray(item?.values) && item.values.length) return `значения: ${item.values.join(" | ")}`;
  if (item?.criterion1 && item?.criterion2) return `${item.criterion1} ${item.operator ?? "и"} ${item.criterion2}`;
  if (item?.criterion1) return String(item.criterion1);
  if (item?.dynamicCriteria) return String(item.dynamicCriteria);
  if (item?.color) return `цвет ${item.color}`;
  return String(item?.filterOn ?? "условие");
}
