/**
 * Реестр свойств оформления.
 *
 * До этого `format_range` знал три свойства — числовой формат, жирность
 * и заливку, — и каждое было прописано по отдельности в пяти местах: чтение,
 * запись, сверка, отмена, предпросмотр. Добавлять так ещё десяток свойств
 * значило бы размножить эти места. Теперь свойство описывается здесь один раз:
 * как его прочитать, записать, чего ждать после записи и как сравнивать.
 *
 * Модуль не знает о плане и отмене: он только отвечает, что такое свойство.
 * Поэтому его можно проверять без Excel, а использовать и в плане, и в отмене.
 */

export type BorderMode = "all" | "outline" | "inside" | "none";
export type BorderWeight = "Thin" | "Medium" | "Thick";
export type HorizontalAlignment = "General" | "Left" | "Center" | "Right" | "Justify";
export type VerticalAlignment = "Top" | "Center" | "Bottom";
export type AutofitMode = "columns" | "rows" | "both";

export interface BorderRequest {
  mode: BorderMode;
  color?: string;
  weight?: BorderWeight;
}

/** Что именно просят изменить в оформлении. Остальные свойства не трогаются. */
export interface FormatRequest {
  numberFormat?: string;
  bold?: boolean;
  fillColor?: string;
  italic?: boolean;
  underline?: boolean;
  fontColor?: string;
  fontSize?: number;
  fontName?: string;
  horizontalAlignment?: HorizontalAlignment;
  verticalAlignment?: VerticalAlignment;
  wrapText?: boolean;
  columnWidth?: number;
  rowHeight?: number;
  borders?: BorderRequest;
}

export type FormatKey = keyof FormatRequest;

/**
 * Состояние свойств. Значение либо однородно по области, либо равно null:
 * так Office.js сообщает о неоднородности. Подавать null как «не задано»
 * нельзя — по этому снимку ловится ручная правка перед запуском.
 */
export type FormatSnapshot = Partial<Record<FormatKey, unknown>>;

/** Размеры области: от них зависит, какие границы у неё вообще есть. */
export interface RangeShape {
  rowCount: number;
  columnCount: number;
}

export interface FormatProperty {
  readonly key: FormatKey;
  /** Где живёт свойство: у ячейки, у столбца или у строки. От этого зависит отмена. */
  readonly scope: "cell" | "column" | "row";
  load(range: any, shape: RangeShape): void;
  read(range: any, shape: RangeShape): unknown;
  write(range: any, value: unknown, shape: RangeShape): void;
  /** Каким свойство прочитается после записи запрошенного значения. */
  expected(value: unknown, shape: RangeShape): unknown;
  same(a: unknown, b: unknown): boolean;
  /** Как вернуть свойство одной ячейки, столбца или строки из снимка отмены. */
  restore?(range: any, value: unknown, shape: RangeShape): void;
}

/* --- сравнение ------------------------------------------------------------ */

/**
 * Приводит строку к сравнимому виду.
 *
 * Проверка в Excel 17 сентября 2026 года: код `0.00 ₽` Excel сохраняет как
 * `0.00 \₽` — экранирует литеральный символ. Это тот же формат, но буквальное
 * сравнение объявляло его расхождением и останавливало задачу. Литерал в коде
 * формата можно записать тремя равнозначными способами: как есть, через
 * обратный слеш и в двойных кавычках, — все три сводятся к одному.
 * Цвета Excel пишет в своём регистре, поэтому регистр тоже не учитывается.
 */
export function canonicalFormatText(value: string): string {
  return value
    // Символ валюты Excel может записать с кодом языка: [$₽-419] — тот же ₽.
    .replace(/\[\$([^\]-]*)-[0-9a-f]+\]/gi, "$1")
    .replace(/"([^"]*)"/g, "$1")
    .replace(/\\(.)/g, "$1")
    .toLowerCase();
}

/** Цвета и коды формата сравниваются в приведённом виде, прочее — строго. */
export function sameFormatValue(a: unknown, b: unknown): boolean {
  return typeof a === "string" && typeof b === "string"
    ? canonicalFormatText(a) === canonicalFormatText(b)
    : a === b;
}

/**
 * Ширина и высота сравниваются с допуском.
 *
 * Excel приводит размеры к сетке экрана: высота строки кратна 0,75 пункта,
 * а ширина столбца — ширине символа шрифта по умолчанию. Запрошенные 100
 * пунктов прочитаются как 99,75 или 100,5, и это не расхождение.
 */
const SIZE_TOLERANCE = 1.5;

function sameSize(a: unknown, b: unknown): boolean {
  if (typeof a === "number" && typeof b === "number") return Math.abs(a - b) <= SIZE_TOLERANCE;
  return a === b;
}

export function normalizeColor(value: string): string {
  const trimmed = value.trim();
  if (!/^#?[0-9A-Fa-f]{6}$/.test(trimmed)) {
    throw new Error(`Цвет "${value}" не в формате HEX, ожидается #RRGGBB.`);
  }
  return (trimmed.startsWith("#") ? trimmed : `#${trimmed}`).toUpperCase();
}

/* --- границы -------------------------------------------------------------- */

const OUTER_EDGES = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"] as const;

/**
 * Какие границы есть у области такой формы.
 *
 * Внутренние горизонтальные есть только при нескольких строках, вертикальные —
 * при нескольких столбцах. Спрашивать Excel о несуществующей внутренней
 * границе одной ячейки бессмысленно, а сверять её — тем более.
 */
export function edgesOf(shape: RangeShape, mode: BorderMode | "every"): string[] {
  const inside: string[] = [];
  if (shape.rowCount > 1) inside.push("InsideHorizontal");
  if (shape.columnCount > 1) inside.push("InsideVertical");
  if (mode === "outline") return [...OUTER_EDGES];
  if (mode === "inside") return inside;
  return [...OUTER_EDGES, ...inside];
}

/** Одна граница в сравнимом виде: «None» или «стиль|толщина|цвет». */
function edgeText(style: unknown, weight: unknown, color: unknown): string | null {
  if (style === null || style === undefined) return null;
  if (style === "None") return "None";
  if (weight === null || color === null) return null;
  return `${style}|${weight}|${typeof color === "string" ? color.toUpperCase() : color}`;
}

/** Где лежат загруженные объекты границ: чтение должно брать ровно их. */
const loadedBorders = new WeakMap<object, Record<string, any>>();

function loadBorders(range: any, shape: RangeShape) {
  const items: Record<string, any> = {};
  for (const edge of edgesOf(shape, "every")) {
    const item = range.format.borders.getItem(edge);
    item.load(["style", "weight", "color"]);
    items[edge] = item;
  }
  loadedBorders.set(range, items);
}

function readBorders(range: any): Record<string, string | null> {
  const items = loadedBorders.get(range) ?? {};
  const result: Record<string, string | null> = {};
  for (const [edge, item] of Object.entries(items)) result[edge] = edgeText(item.style, item.weight, item.color);
  return result;
}

/** Каким станет каждая затронутая граница. Незатронутые в ожидание не входят. */
export function expectedBorders(request: BorderRequest, shape: RangeShape): Record<string, string> {
  const expected: Record<string, string> = {};
  const line = `Continuous|${request.weight ?? "Thin"}|${normalizeColor(request.color ?? "#000000")}`;
  if (request.mode === "none") {
    for (const edge of edgesOf(shape, "every")) expected[edge] = "None";
    return expected;
  }
  for (const edge of edgesOf(shape, request.mode)) expected[edge] = line;
  return expected;
}

/**
 * Совпадают ли границы: сравниваются только те, что есть в ожидании.
 * Внутренние границы при режиме «outline» не трогаются и не сверяются.
 */
function sameBorders(actual: unknown, expected: unknown): boolean {
  if (!actual || !expected || typeof actual !== "object" || typeof expected !== "object") return actual === expected;
  const a = actual as Record<string, unknown>;
  const e = expected as Record<string, unknown>;
  const keys = new Set([...Object.keys(a), ...Object.keys(e)]);
  for (const key of keys) {
    // Ожидание по режиму может не включать часть границ — их не сверяем.
    if (!(key in e) || !(key in a)) continue;
    if (!sameFormatValue(a[key], e[key])) return false;
  }
  return true;
}

function writeBorders(range: any, value: unknown, shape: RangeShape) {
  const request = value as BorderRequest;
  const edges = request.mode === "none" ? edgesOf(shape, "every") : edgesOf(shape, request.mode);
  for (const edge of edges) {
    const item = range.format.borders.getItem(edge);
    if (request.mode === "none") {
      item.style = "None";
      continue;
    }
    item.style = "Continuous";
    item.weight = request.weight ?? "Thin";
    item.color = normalizeColor(request.color ?? "#000000");
  }
}

/** Отмена возвращает границы ячейки такими, какими они были: стиль, толщину и цвет. */
function restoreBorders(range: any, value: unknown) {
  if (!value || typeof value !== "object") return;
  for (const [edge, text] of Object.entries(value as Record<string, string | null>)) {
    if (text === null) continue;
    const item = range.format.borders.getItem(edge);
    if (text === "None") { item.style = "None"; continue; }
    const [style, weight, color] = text.split("|");
    item.style = style;
    item.weight = weight;
    item.color = color;
  }
}

/* --- реестр --------------------------------------------------------------- */

/** Однородное значение матрицы или null, если ячейки различаются. */
function uniform(matrix: unknown): unknown {
  if (!Array.isArray(matrix)) return matrix ?? null;
  const first = (matrix as unknown[][])[0]?.[0];
  for (const row of matrix as unknown[][]) {
    for (const value of row) if (!sameFormatValue(value, first)) return null;
  }
  return first ?? null;
}

function fontProperty(key: FormatKey, name: string, options: Partial<FormatProperty> = {}): FormatProperty {
  return {
    key,
    scope: "cell",
    load: (range) => range.format.font.load(name),
    read: (range) => range.format.font[name] ?? null,
    write: (range, value) => { range.format.font[name] = value; },
    expected: (value) => value,
    same: sameFormatValue,
    ...options
  };
}

function formatProperty(key: FormatKey, name: string, options: Partial<FormatProperty> = {}): FormatProperty {
  return {
    key,
    scope: "cell",
    load: (range) => range.format.load(name),
    read: (range) => range.format[name] ?? null,
    write: (range, value) => { range.format[name] = value; },
    expected: (value) => value,
    same: sameFormatValue,
    ...options
  };
}

/**
 * Порядок важен: в этом порядке свойства перечисляются в предпросмотре
 * и применяются. Первые три — прежние, чтобы не менять сложившийся вид.
 */
export const FORMAT_PROPERTIES: readonly FormatProperty[] = [
  {
    key: "numberFormat",
    scope: "cell",
    load: (range) => range.load("numberFormat"),
    read: (range) => uniform(range.numberFormat),
    write: (range, value, shape) => {
      range.numberFormat = Array.from({ length: shape.rowCount }, () =>
        Array.from({ length: shape.columnCount }, () => value as string));
    },
    expected: (value) => value,
    same: sameFormatValue
  },
  fontProperty("bold", "bold"),
  {
    key: "fillColor",
    scope: "cell",
    load: (range) => range.format.fill.load("color"),
    read: (range) => range.format.fill.color ?? null,
    write: (range, value) => { range.format.fill.color = normalizeColor(String(value)); },
    expected: (value) => normalizeColor(String(value)),
    same: sameFormatValue,
    // Пустая заливка — это не белый цвет, а её отсутствие.
    restore: (range, value) => {
      if (value) range.format.fill.color = value;
      else range.format.fill.clear();
    }
  },
  fontProperty("italic", "italic"),
  fontProperty("underline", "underline", {
    // Excel хранит подчёркивание словом; просим мы его да или нет.
    read: (range) => {
      const value = range.format.font.underline;
      if (value === null || value === undefined) return null;
      return value === "None" ? false : value === "Single" ? true : value;
    },
    write: (range, value) => { range.format.font.underline = value ? "Single" : "None"; },
    restore: (range, value) => {
      if (value === null || value === undefined) return;
      range.format.font.underline = typeof value === "string" ? value : value ? "Single" : "None";
    }
  }),
  fontProperty("fontColor", "color", {
    write: (range, value) => { range.format.font.color = normalizeColor(String(value)); },
    expected: (value) => normalizeColor(String(value))
  }),
  fontProperty("fontSize", "size"),
  fontProperty("fontName", "name"),
  formatProperty("horizontalAlignment", "horizontalAlignment"),
  formatProperty("verticalAlignment", "verticalAlignment"),
  formatProperty("wrapText", "wrapText"),
  formatProperty("columnWidth", "columnWidth", { scope: "column", same: sameSize }),
  formatProperty("rowHeight", "rowHeight", { scope: "row", same: sameSize }),
  {
    key: "borders",
    scope: "cell",
    load: loadBorders,
    read: (range) => readBorders(range),
    write: writeBorders,
    expected: (value, shape) => expectedBorders(value as BorderRequest, shape),
    same: sameBorders,
    restore: restoreBorders
  }
];

export const FORMAT_PROPERTY = new Map<FormatKey, FormatProperty>(
  FORMAT_PROPERTIES.map((property) => [property.key, property])
);

/* --- запрос --------------------------------------------------------------- */

const LIMITS = {
  fontSize: [1, 409],
  columnWidth: [0, 1600],
  rowHeight: [0, 409]
} as const;

/**
 * Разбирает аргументы инструмента в запрос. Всё, что можно проверить без Excel,
 * проверяется здесь: до предпросмотра, а не после подтверждения.
 */
export function parseFormatRequest(args: Record<string, unknown>): { request: FormatRequest; autofit?: AutofitMode } {
  const request: FormatRequest = {};
  for (const key of ["numberFormat", "fontName"] as const) {
    if (typeof args[key] === "string") {
      if (!(args[key] as string).trim()) throw new Error(`${key} не может быть пустой строкой.`);
      request[key] = args[key] as string;
    }
  }
  for (const key of ["bold", "italic", "underline", "wrapText"] as const) {
    if (typeof args[key] === "boolean") request[key] = args[key] as boolean;
  }
  for (const key of ["fillColor", "fontColor"] as const) {
    if (typeof args[key] === "string") request[key] = normalizeColor(args[key] as string);
  }
  for (const key of ["fontSize", "columnWidth", "rowHeight"] as const) {
    if (typeof args[key] === "number") {
      const [min, max] = LIMITS[key];
      const value = args[key] as number;
      if (!Number.isFinite(value) || value < min || value > max) {
        throw new Error(`${key} должен быть от ${min} до ${max}, получено ${value}.`);
      }
      request[key] = value;
    }
  }
  if (typeof args.horizontalAlignment === "string") request.horizontalAlignment = args.horizontalAlignment as HorizontalAlignment;
  if (typeof args.verticalAlignment === "string") request.verticalAlignment = args.verticalAlignment as VerticalAlignment;
  if (typeof args.borders === "string") {
    request.borders = {
      mode: args.borders as BorderMode,
      ...(typeof args.borderColor === "string" ? { color: normalizeColor(args.borderColor) } : {}),
      ...(typeof args.borderWeight === "string" ? { weight: args.borderWeight as BorderWeight } : {})
    };
  } else if (typeof args.borderColor === "string" || typeof args.borderWeight === "string") {
    throw new Error("Цвет и толщина границы задаются вместе с borders: укажите, какие границы рисовать.");
  }

  const autofit = typeof args.autofit === "string" ? args.autofit as AutofitMode : undefined;
  // Точная ширина и автоподбор взаимно исключают друг друга: одно из двух
  // молча затёрло бы другое, и сверка упала бы на ровном месте.
  if (autofit && (autofit === "columns" || autofit === "both") && request.columnWidth !== undefined) {
    throw new Error("Нельзя одновременно задать columnWidth и автоподбор ширины столбцов: выберите одно.");
  }
  if (autofit && (autofit === "rows" || autofit === "both") && request.rowHeight !== undefined) {
    throw new Error("Нельзя одновременно задать rowHeight и автоподбор высоты строк: выберите одно.");
  }
  return { request, ...(autofit ? { autofit } : {}) };
}

export function requestedFormatKeys(request: FormatRequest): FormatKey[] {
  return FORMAT_PROPERTIES.map((property) => property.key).filter((key) => request[key] !== undefined);
}

/** Каким станет оформление, если операция пройдёт: это показывает предпросмотр
 * и с этим же сверяется результат. */
export function expectedFormatSnapshot(request: FormatRequest, shape: RangeShape = { rowCount: 1, columnCount: 1 }): FormatSnapshot {
  const expected: FormatSnapshot = {};
  for (const key of requestedFormatKeys(request)) {
    expected[key] = FORMAT_PROPERTY.get(key)!.expected(request[key], shape);
  }
  return expected;
}

export function formatSnapshotsEqual(a: FormatSnapshot, b: FormatSnapshot): boolean {
  const keys = new Set([...Object.keys(a), ...Object.keys(b)]) as Set<FormatKey>;
  for (const key of keys) {
    const property = FORMAT_PROPERTY.get(key);
    if (!(property ? property.same(a[key], b[key]) : sameFormatValue(a[key], b[key]))) return false;
  }
  return true;
}

/** Какие свойства разошлись — для сообщения об ошибке и для отчёта. */
export function formatDifferences(actual: FormatSnapshot, expected: FormatSnapshot): FormatKey[] {
  return (Object.keys(expected) as FormatKey[]).filter((key) => {
    const property = FORMAT_PROPERTY.get(key);
    return !(property ? property.same(actual[key], expected[key]) : sameFormatValue(actual[key], expected[key]));
  });
}

/** Загружает и читает набор свойств у одного диапазона — области или ячейки. */
export function loadFormat(range: any, keys: readonly FormatKey[], shape: RangeShape) {
  for (const key of keys) FORMAT_PROPERTY.get(key)!.load(range, shape);
}

export function readFormat(range: any, keys: readonly FormatKey[], shape: RangeShape): FormatSnapshot {
  const snapshot: FormatSnapshot = {};
  for (const key of keys) snapshot[key] = FORMAT_PROPERTY.get(key)!.read(range, shape);
  return snapshot;
}
