/**
 * Чистая логика второй партии оформления: закрепление областей, правила
 * условного форматирования и превращение области в таблицу Excel.
 *
 * Все три — не свойства ячеек, и проверяются иначе, чем `format_range`.
 * Закрепление — настройка вида листа: сверяется по месту закрепления.
 * Условное форматирование — правило поверх ячеек: цвет ячейки после него
 * через API не прочитать, поэтому сверяется само правило, а какие ячейки оно
 * подсветит, панель оценивает сама и называет это оценкой. Таблица — объект
 * листа со своим поведением: сверяются её границы, стиль и заголовки.
 *
 * Модуль не обращается к Excel, поэтому проверяется без него.
 */

import { parseA1Rect } from "./a1";
import { normalizeColor } from "./formatProps";

/* --- закрепление ---------------------------------------------------------- */

export interface FreezeState {
  /** Сколько верхних строк закреплено. */
  rows: number;
  /** Сколько левых столбцов закреплено. */
  columns: number;
}

export const MAX_FROZEN_ROWS = 100;
export const MAX_FROZEN_COLUMNS = 50;

/**
 * Место закрепления в понятном виде.
 *
 * Excel отдаёт закреплённую область адресом: `$1:$2` — две строки, `A:B` —
 * два столбца, `A1:B2` — и то и другое. Нет закрепления — нет адреса.
 */
export function parseFreezeLocation(address: string | null | undefined): FreezeState | null {
  if (!address) return { rows: 0, columns: 0 };
  const rect = parseA1Rect(address.slice(address.lastIndexOf("!") + 1));
  if (!rect) return null;
  if (rect.kind === "rows") return { rows: rect.rowEnd, columns: 0 };
  if (rect.kind === "columns") return { rows: 0, columns: rect.columnEnd };
  return { rows: rect.rowEnd, columns: rect.columnEnd };
}

export function sameFreeze(a: FreezeState | null, b: FreezeState | null): boolean {
  return Boolean(a && b && a.rows === b.rows && a.columns === b.columns);
}

export function describeFreeze(state: FreezeState | null): string {
  if (!state) return "не удалось определить";
  if (!state.rows && !state.columns) return "ничего не закреплено";
  const parts: string[] = [];
  if (state.rows) parts.push(`строки 1–${state.rows}`);
  if (state.columns) parts.push(`столбцы 1–${state.columns}`);
  return parts.join(" и ");
}

export function parseFreezeRequest(args: Record<string, unknown>): FreezeState {
  const rows = args.rows === undefined ? 0 : args.rows;
  const columns = args.columns === undefined ? 0 : args.columns;
  if (!Number.isInteger(rows) || (rows as number) < 0 || (rows as number) > MAX_FROZEN_ROWS) {
    throw new Error(`rows должен быть целым от 0 до ${MAX_FROZEN_ROWS}.`);
  }
  if (!Number.isInteger(columns) || (columns as number) < 0 || (columns as number) > MAX_FROZEN_COLUMNS) {
    throw new Error(`columns должен быть целым от 0 до ${MAX_FROZEN_COLUMNS}.`);
  }
  return { rows: rows as number, columns: columns as number };
}

/* --- условное форматирование ---------------------------------------------- */

export type ComparisonRule =
  | "greaterThan"
  | "lessThan"
  | "greaterOrEqual"
  | "lessOrEqual"
  | "equalTo"
  | "notEqualTo"
  | "between";

export type ConditionalRuleKind = ComparisonRule | "textContains" | "formula" | "colorScale" | "dataBar";

/** Имена операторов Office.js для сравнения значения ячейки. */
export const CELL_VALUE_OPERATOR: Record<ComparisonRule, string> = {
  greaterThan: "GreaterThan",
  lessThan: "LessThan",
  greaterOrEqual: "GreaterThanOrEqual",
  lessOrEqual: "LessThanOrEqual",
  equalTo: "EqualTo",
  notEqualTo: "NotEqualTo",
  between: "Between"
};

export interface HighlightFormat {
  fillColor?: string;
  fontColor?: string;
  bold?: boolean;
}

export interface ConditionalRequest {
  rule: ConditionalRuleKind;
  /** Правило сравнения: с чем сравнивать. Для between — нижняя граница. */
  value?: number | string;
  /** Верхняя граница для between. */
  value2?: number;
  /** Текст для textContains. */
  text?: string;
  /** Условие formula: формула для левой верхней ячейки области, со знаком «=». */
  formula?: string;
  /** Как подсветить ячейку для сравнения и текста. */
  highlight?: HighlightFormat;
  /** Цветовая шкала: цвета минимума, середины (необязательно) и максимума. */
  scale?: { minColor: string; midColor?: string; maxColor: string };
  /** Гистограмма в ячейке: цвет полосы. */
  barColor?: string;
}

const COMPARISONS = new Set(Object.keys(CELL_VALUE_OPERATOR));

/** Значение для формулы правила: число как есть, текст — в кавычках Excel. */
export function ruleFormula(value: number | string): string {
  if (typeof value === "number") return String(value);
  return `="${value.replace(/"/g, '""')}"`;
}

export function parseConditionalRequest(args: Record<string, unknown>): ConditionalRequest {
  const rule = args.rule as ConditionalRuleKind;
  const highlight: HighlightFormat = {
    ...(typeof args.fillColor === "string" ? { fillColor: normalizeColor(args.fillColor) } : {}),
    ...(typeof args.fontColor === "string" ? { fontColor: normalizeColor(args.fontColor) } : {}),
    ...(typeof args.bold === "boolean" ? { bold: args.bold } : {})
  };
  const hasHighlight = Object.keys(highlight).length > 0;

  if (COMPARISONS.has(rule)) {
    if (typeof args.value !== "number" && typeof args.value !== "string") {
      throw new Error(`Для правила ${rule} нужно значение value — с чем сравнивать.`);
    }
    if (rule === "between") {
      if (typeof args.value !== "number" || typeof args.value2 !== "number") {
        throw new Error("Для between нужны два числа: value — нижняя граница, value2 — верхняя.");
      }
      if (args.value > args.value2) throw new Error("Для between value должен быть не больше value2.");
    } else if (typeof args.value === "string" && rule !== "equalTo" && rule !== "notEqualTo") {
      // «Больше, чем текст» Excel сравнивает по алфавиту — почти всегда это не то, что имели в виду.
      throw new Error(`Правило ${rule} сравнивает числа; для текста используйте equalTo, notEqualTo или textContains.`);
    }
    if (!hasHighlight) throw new Error("Укажите, как подсветить ячейки: fillColor, fontColor или bold.");
    return {
      rule,
      value: args.value as number | string,
      ...(rule === "between" ? { value2: args.value2 as number } : {}),
      highlight
    };
  }
  if (rule === "textContains") {
    if (typeof args.text !== "string" || !args.text.trim()) throw new Error("Для textContains нужен непустой text.");
    if (!hasHighlight) throw new Error("Укажите, как подсветить ячейки: fillColor, fontColor или bold.");
    return { rule, text: args.text, highlight };
  }
  if (rule === "formula") {
    const text = typeof args.formula === "string" ? args.formula.trim() : "";
    if (!text || text === "=") throw new Error("Для правила formula нужна формула условия, например =$C2>1000.");
    if (!hasHighlight) throw new Error("Укажите, как подсветить ячейки: fillColor, fontColor или bold.");
    return { rule, formula: text.startsWith("=") ? text : `=${text}`, highlight };
  }
  if (rule === "colorScale") {
    if (typeof args.minColor !== "string" || typeof args.maxColor !== "string") {
      throw new Error("Для цветовой шкалы нужны minColor и maxColor; midColor — по желанию.");
    }
    return {
      rule,
      scale: {
        minColor: normalizeColor(args.minColor),
        ...(typeof args.midColor === "string" ? { midColor: normalizeColor(args.midColor) } : {}),
        maxColor: normalizeColor(args.maxColor)
      }
    };
  }
  if (rule === "dataBar") {
    return { rule, barColor: normalizeColor(typeof args.barColor === "string" ? args.barColor : "#638EC6") };
  }
  throw new Error(`Неизвестное правило ${String(rule)}.`);
}

/** Одна строка о правиле — для предпросмотра и отчёта. */
export function describeConditionalRule(request: ConditionalRequest): string {
  const how = request.highlight
    ? [
        request.highlight.fillColor && `заливка ${request.highlight.fillColor}`,
        request.highlight.fontColor && `текст ${request.highlight.fontColor}`,
        request.highlight.bold === true && "жирный",
        request.highlight.bold === false && "не жирный"
      ].filter(Boolean).join(", ")
    : "";
  const shown = (value: unknown) => (typeof value === "string" ? `«${value}»` : String(value));
  switch (request.rule) {
    case "greaterThan": return `значение больше ${shown(request.value)} → ${how}`;
    case "lessThan": return `значение меньше ${shown(request.value)} → ${how}`;
    case "greaterOrEqual": return `значение не меньше ${shown(request.value)} → ${how}`;
    case "lessOrEqual": return `значение не больше ${shown(request.value)} → ${how}`;
    case "equalTo": return `значение равно ${shown(request.value)} → ${how}`;
    case "notEqualTo": return `значение не равно ${shown(request.value)} → ${how}`;
    case "between": return `значение от ${request.value} до ${request.value2} → ${how}`;
    case "textContains": return `текст содержит «${request.text}» → ${how}`;
    case "formula": return `формула ${request.formula} истинна → ${how}`;
    case "colorScale": {
      const scale = request.scale!;
      return `цветовая шкала: минимум ${scale.minColor}${scale.midColor ? `, середина ${scale.midColor}` : ""}, максимум ${scale.maxColor}`;
    }
    case "dataBar": return `гистограмма в ячейках цветом ${request.barColor}`;
  }
}

/** Тип правила в Office.js — по нему сверяется, что добавилось именно оно. */
export function officeRuleType(rule: ConditionalRuleKind): string {
  if (COMPARISONS.has(rule)) return "CellValue";
  if (rule === "textContains") return "ContainsText";
  if (rule === "formula") return "Custom";
  if (rule === "colorScale") return "ColorScale";
  return "DataBar";
}

/**
 * Правило условного форматирования, как его прочитала панель из Excel.
 * Поля заполнены по типу правила; у незаданных свойств Excel отдаёт null.
 */
export interface RuleSnapshot {
  id: string;
  type: string;
  priority: number | null;
  /** Адрес области правила; null — Excel её не отдал. */
  range: string | null;
  /** Условие: cellValue.rule, textComparison.rule и подобные. */
  rule?: Record<string, unknown> | null;
  fill?: string | null;
  fontColor?: string | null;
  bold?: boolean | null;
  /** Точки цветовой шкалы, критерии значков. */
  criteria?: Record<string, any> | null;
  barColor?: string | null;
  /** Содержимое правила этого типа панель не читает: сверены только тип, приоритет и область. */
  contentUnread?: true;
}

/** Формула в правиле Excel приходит со знаком равенства: «=1000». */
function sameRuleFormula(actual: unknown, expected: string): boolean {
  const clean = (value: unknown) => String(value ?? "").trim().replace(/^=/, "").toLowerCase();
  return clean(actual) === clean(expected);
}

const sameColor = (actual: unknown, expected: string) => String(actual ?? "").toUpperCase() === expected.toUpperCase();
const shown = (value: unknown) => (value === null || value === undefined || value === "" ? "нет" : String(value));

/**
 * Чем добавленное правило расходится с запрошенным (план стабилизации, S3.2).
 *
 * Сверяются все свойства, которые можно запросить: тип, область, условие
 * и оператор, заливка, цвет и жирность текста, все точки цветовой шкалы,
 * цвет полосы. Незапрошенная середина шкалы — тоже расхождение: шкала
 * из трёх цветов выглядит иначе, чем из двух. Пустой список — правило
 * подтверждено именно такое, как просили, а не просто «что-то добавилось».
 */
export function conditionalRuleMismatches(request: ConditionalRequest, rule: RuleSnapshot, address: string): string[] {
  const problems: string[] = [];
  const type = officeRuleType(request.rule);
  if (rule.type !== type) problems.push(`тип ${rule.type} вместо ${type}`);
  const where = rule.range === null ? null : rule.range.slice(rule.range.lastIndexOf("!") + 1).replace(/\$/g, "");
  if (where !== address) problems.push(where === null ? "область правила не прочиталась" : `область ${where} вместо ${address}`);
  if (rule.contentUnread) {
    problems.push("содержимое правила не прочиталось обратно");
    return problems;
  }

  const highlight = request.highlight ?? {};
  const checkHighlight = () => {
    if (highlight.fillColor && !sameColor(rule.fill, highlight.fillColor)) problems.push(`цвет заливки ${shown(rule.fill)} вместо ${highlight.fillColor}`);
    if (highlight.fontColor && !sameColor(rule.fontColor, highlight.fontColor)) problems.push(`цвет текста ${shown(rule.fontColor)} вместо ${highlight.fontColor}`);
    if (typeof highlight.bold === "boolean" && rule.bold !== highlight.bold) {
      problems.push(`жирность ${rule.bold === true ? "включена" : rule.bold === false ? "выключена" : "не задана"} вместо ${highlight.bold ? "включена" : "выключена"}`);
    }
  };

  if (request.rule === "textContains") {
    if (rule.rule?.operator !== "Contains") problems.push(`оператор ${shown(rule.rule?.operator)} вместо Contains`);
    if (String(rule.rule?.text ?? "") !== String(request.text)) problems.push(`текст условия «${shown(rule.rule?.text)}» вместо «${request.text}»`);
    checkHighlight();
  } else if (request.rule === "formula") {
    // Замер 24 сентября 2026 года: Excel отдаёт формулу так, как её записали, —
    // для левой верхней ячейки области, со знаком «=».
    if (!sameRuleFormula(rule.rule?.formula, String(request.formula))) problems.push(`формула ${shown(rule.rule?.formula)} вместо ${request.formula}`);
    checkHighlight();
  } else if (request.rule === "colorScale") {
    const scale = request.scale!;
    const point = (name: string, text: string, color: string, kind: string) => {
      const item = rule.criteria?.[name];
      if (!item || !sameColor(item.color, color) || item.type !== kind) {
        problems.push(`${text} шкалы: ${item ? `${shown(item.color)}, ${shown(item.type)}` : "нет"} вместо ${color}, ${kind}`);
      }
    };
    point("minimum", "минимум", scale.minColor, "LowestValue");
    point("maximum", "максимум", scale.maxColor, "HighestValue");
    const mid = rule.criteria?.midpoint;
    if (scale.midColor) {
      if (!mid || !sameColor(mid.color, scale.midColor) || mid.type !== "Percentile" || String(mid.formula) !== "50") {
        problems.push(`середина шкалы: ${mid ? `${shown(mid.color)}, ${shown(mid.type)} ${shown(mid.formula)}` : "нет"} вместо ${scale.midColor}, 50-й процентиль`);
      }
    } else if (mid && mid.color) {
      problems.push(`у шкалы появилась середина ${mid.color}, которую не просили`);
    }
  } else if (request.rule === "dataBar") {
    if (!sameColor(rule.barColor, String(request.barColor))) problems.push(`цвет полосы ${shown(rule.barColor)} вместо ${request.barColor}`);
  } else {
    const operator = CELL_VALUE_OPERATOR[request.rule as ComparisonRule];
    if (rule.rule?.operator !== operator) problems.push(`оператор ${shown(rule.rule?.operator)} вместо ${operator}`);
    if (!sameRuleFormula(rule.rule?.formula1, ruleFormula(request.value as number | string))) {
      problems.push(`значение ${shown(rule.rule?.formula1)} вместо ${ruleFormula(request.value as number | string)}`);
    }
    if (request.rule === "between" && !sameRuleFormula(rule.rule?.formula2, ruleFormula(request.value2 as number))) {
      problems.push(`верхняя граница ${shown(rule.rule?.formula2)} вместо ${ruleFormula(request.value2 as number)}`);
    }
    checkHighlight();
  }
  return problems;
}

function numeric(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

/**
 * Совпадает ли значение ячейки с правилом — оценка панели.
 *
 * Excel применяет правило сам и через API не сообщает, какие ячейки
 * подсвечены. Поэтому панель считает сама — по тем же правилам: пустая ячейка
 * для сравнения с числом равна нулю, текст с числом не сравнивается,
 * «содержит» не различает регистр. Это оценка, и она так и называется.
 */
export function ruleMatches(request: ConditionalRequest, value: unknown): boolean | null {
  // Условие по формуле считает только Excel: оценки у панели нет.
  if (request.rule === "colorScale" || request.rule === "dataBar" || request.rule === "formula") return null;
  if (request.rule === "textContains") {
    if (value === null || value === undefined || value === "") return false;
    return String(value).toLowerCase().includes(String(request.text).toLowerCase());
  }
  if (typeof request.value === "string") {
    const equal = typeof value === "string" && value.toLowerCase() === request.value.toLowerCase();
    return request.rule === "equalTo" ? equal : !equal;
  }
  const cell = value === "" || value === null || value === undefined ? 0 : numeric(value);
  if (cell === null) return false;
  const limit = request.value as number;
  switch (request.rule) {
    case "greaterThan": return cell > limit;
    case "lessThan": return cell < limit;
    case "greaterOrEqual": return cell >= limit;
    case "lessOrEqual": return cell <= limit;
    case "equalTo": return cell === limit;
    case "notEqualTo": return cell !== limit;
    case "between": return cell >= limit && cell <= (request.value2 as number);
  }
  return null;
}

/* --- таблица Excel ---------------------------------------------------------- */

const TABLE_STYLE = /^TableStyle(Light([1-9]|1\d|2[01])|Medium([1-9]|1\d|2[0-8])|Dark([1-9]|1[01]))$/;
const TABLE_NAME = /^[A-Za-z_Ѐ-ӿ][A-Za-z0-9_.Ѐ-ӿ]{0,254}$/;

export function checkTableStyle(style: string): string {
  if (!TABLE_STYLE.test(style)) {
    throw new Error(`Стиль «${style}» не существует. Встроенные стили: TableStyleLight1–21, TableStyleMedium1–28, TableStyleDark1–11.`);
  }
  return style;
}

export function checkTableName(name: string): string {
  const trimmed = name.trim();
  // Имя, похожее на адрес ячейки, Excel не примет: A1 или R1C1 — это ссылки.
  if (!TABLE_NAME.test(trimmed) || /^[A-Za-z]{1,3}\d+$/.test(trimmed) || /^[Rr]\d*[Cc]\d*$/.test(trimmed)) {
    throw new Error(`Имя таблицы «${name}» не подходит: оно должно начинаться с буквы, без пробелов и не быть похожим на адрес ячейки.`);
  }
  return trimmed;
}

/**
 * Что Excel сделает с заголовками при создании таблицы.
 *
 * Заголовок таблицы — всегда текст и всегда уникален. Пустой заголовок Excel
 * заменит на «Столбец1», повтор — на «Имя2», формулу — на её текущее значение.
 * Это изменение данных, и о нём надо сказать до операции, а не после.
 */
export function headerProblems(headerValues: readonly unknown[], headerFormulas: readonly unknown[]): string[] {
  const problems: string[] = [];
  const seen = new Map<string, number>();
  headerValues.forEach((value, index) => {
    const column = index + 1;
    const text = value === null || value === undefined ? "" : String(value).trim();
    if (!text) problems.push(`столбец ${column}: пустой заголовок — Excel подставит «Столбец${column}» или похожее имя`);
    const formula = headerFormulas[index];
    if (typeof formula === "string" && formula.startsWith("=")) {
      problems.push(`столбец ${column}: в заголовке формула ${formula} — Excel заменит её значением «${text}»`);
    }
    if (typeof value === "number") problems.push(`столбец ${column}: заголовок-число ${value} станет текстом`);
    if (text) {
      const key = text.toLowerCase();
      const first = seen.get(key);
      if (first !== undefined) problems.push(`столбцы ${first} и ${column}: одинаковый заголовок «${text}» — Excel переименует второй`);
      else seen.set(key, column);
    }
  });
  return problems;
}

/* --- имя листа -------------------------------------------------------------- */

/** Excel не принимает эти знаки в имени листа. */
const FORBIDDEN_IN_SHEET_NAME = /[:\\/?*\[\]]/;
export const MAX_SHEET_NAME = 31;

/**
 * Проверяет имя нового листа по правилам Excel.
 *
 * Правила жёсткие и молчаливые: слишком длинное имя или запрещённый знак
 * Excel не исправит, а откажет на середине операции. Совпадение имени
 * он тоже не разрешает, причём без учёта регистра. Поэтому всё это
 * проверяется здесь, до обращения к книге.
 */
export function checkSheetName(name: unknown, existing: readonly string[]): string {
  if (typeof name !== "string") throw new Error("Имя листа должно быть строкой.");
  const trimmed = name.trim();
  if (!trimmed) throw new Error("Имя листа не может быть пустым.");
  if (trimmed.length > MAX_SHEET_NAME) {
    throw new Error(`Имя листа не длиннее ${MAX_SHEET_NAME} знаков, в «${trimmed}» их ${trimmed.length}.`);
  }
  if (FORBIDDEN_IN_SHEET_NAME.test(trimmed)) {
    throw new Error(`В имени листа нельзя использовать : \ / ? * [ ] — проверьте «${trimmed}».`);
  }
  if (trimmed.startsWith("'") || trimmed.endsWith("'")) {
    throw new Error("Имя листа не может начинаться или заканчиваться апострофом.");
  }
  // «История» — служебное имя листа общего доступа, Excel его занимает сам.
  if (/^(history|история)$/i.test(trimmed)) throw new Error(`«${trimmed}» — служебное имя листа Excel, выберите другое.`);
  const clash = existing.find((item) => item.trim().toLowerCase() === trimmed.toLowerCase());
  if (clash) throw new Error(`Лист «${clash}» в книге уже есть: имена листов не повторяются.`);
  return trimmed;
}

/** Свободное имя вида «Отчёт 2»: для подсказки в отказе. */
export function freeSheetName(wanted: string, existing: readonly string[]): string {
  const taken = new Set(existing.map((item) => item.trim().toLowerCase()));
  const base = wanted.trim().slice(0, MAX_SHEET_NAME - 3);
  for (let index = 2; index < 100; index++) {
    const candidate = `${base} ${index}`;
    if (!taken.has(candidate.toLowerCase())) return candidate;
  }
  return `${base} ${Date.now().toString(36)}`;
}
