/**
 * Донастройка сводной (срез 10.7, 07.10.2026): имя и формат полей значений,
 * итоги, автоширина, ещё одно поле в столбцах. Применяется после того, как
 * числа сводной сверены с расчётом панели, и проверяется обратным чтением.
 *
 * Замер в Excel 07.10.2026 (ExcelApi 1.14, русский Excel):
 * - формат поля значений (DataPivotHierarchy.numberFormat) Excel читает в
 *   записи своего языка: «# ##0 "к"», а не «#,##0,"к"», — английская запись
 *   давала «3600,0,к». Поэтому формат переводится по разделителям культуры;
 * - формат поля держится после обновления сводной, в отличие от формата ячеек;
 * - имя поля значений, совпадающее с заголовком источника («Выручка»), Excel
 *   не принимает — к нему добавляется пробел, как делают и вручную;
 * - showRowGrandTotals, showColumnGrandTotals, autoFormat, subtotals поля и
 *   несколько полей в столбцах работают и читаются обратно.
 */

export type GrandTotals = "both" | "rows" | "columns" | "none";
export const GRAND_TOTALS: readonly GrandTotals[] = ["both", "rows", "columns", "none"];

/** Что поставить на поле значений: подпись и формат (английская запись, как у ячеек). */
/**
 * Дополнительные вычисления поля значений (срез 10.9). Замер 07.10.2026:
 * доли, нарастающий итог, отличие от выбранного элемента и ранг работают;
 * «(предыдущий)» элемент, как в окне Excel, надстройкам недоступен —
 * отличие считается только от названного элемента (например, от «янв»).
 */
export type ShowAs =
  | "percentOfGrandTotal" | "percentOfColumnTotal" | "percentOfRowTotal" | "percentOfParentRowTotal"
  | "runningTotal" | "differenceFrom" | "percentDifferenceFrom" | "rankDescending" | "rankAscending" | "none";
export const SHOW_AS: readonly ShowAs[] = [
  "percentOfGrandTotal", "percentOfColumnTotal", "percentOfRowTotal", "percentOfParentRowTotal",
  "runningTotal", "differenceFrom", "percentDifferenceFrom", "rankDescending", "rankAscending", "none"
];
/** Имена Office.js: у ранга по убыванию в API опечатка — «RankDecending». */
const OFFICE_SHOW_AS: Record<ShowAs, string> = {
  percentOfGrandTotal: "PercentOfGrandTotal", percentOfColumnTotal: "PercentOfColumnTotal", percentOfRowTotal: "PercentOfRowTotal",
  percentOfParentRowTotal: "PercentOfParentRowTotal", runningTotal: "RunningTotal", differenceFrom: "DifferenceFrom",
  percentDifferenceFrom: "PercentDifferenceFrom", rankDescending: "RankDecending", rankAscending: "RankAscending", none: "None"
};
export const SHOW_AS_TEXT: Record<ShowAs, string> = {
  percentOfGrandTotal: "% от общего итога", percentOfColumnTotal: "% от итога по столбцу", percentOfRowTotal: "% от итога по строке",
  percentOfParentRowTotal: "% от итога родительской строки", runningTotal: "нарастающий итог", differenceFrom: "отличие от",
  percentDifferenceFrom: "% отличия от", rankDescending: "ранг (1 — наибольшее)", rankAscending: "ранг (1 — наименьшее)", none: "обычные значения"
};
/** Каким вычислениям нужно поле-основа и элемент. */
export const SHOW_AS_NEEDS_FIELD: ReadonlySet<ShowAs> = new Set(["runningTotal", "differenceFrom", "percentDifferenceFrom", "rankDescending", "rankAscending"]);
export const SHOW_AS_NEEDS_ITEM: ReadonlySet<ShowAs> = new Set(["differenceFrom", "percentDifferenceFrom"]);

export interface ShowAsRequest { calculation: ShowAs; baseField?: string; baseItem?: string }

export interface ValueFieldFinish {
  /** Номер поля значений по порядку добавления. */
  index: number;
  label?: string;
  numberFormat?: string;
  showAs?: ShowAsRequest;
}

/** Проверка запроса до Excel: без поля-основы Excel ответит невнятной ошибкой. */
export function checkShowAs(raw: unknown): ShowAsRequest | undefined {
  if (raw === undefined || raw === null) return undefined;
  const value = raw as Partial<ShowAsRequest>;
  if (!SHOW_AS.includes(value.calculation as ShowAs)) throw new Error(`showAs.calculation — одно из: ${SHOW_AS.join(", ")}.`);
  const calculation = value.calculation as ShowAs;
  if (SHOW_AS_NEEDS_FIELD.has(calculation) && !value.baseField?.trim()) {
    throw new Error(`Для «${SHOW_AS_TEXT[calculation]}» нужно baseField — поле строк или столбцов, по которому считать (например «Месяц»).`);
  }
  if (SHOW_AS_NEEDS_ITEM.has(calculation) && !value.baseItem?.trim()) {
    throw new Error(`Для «${SHOW_AS_TEXT[calculation]}» нужно baseItem — элемент, от которого считать (например «янв»). «Предыдущий» Excel надстройкам не даёт.`);
  }
  return { calculation, ...(value.baseField?.trim() ? { baseField: value.baseField.trim() } : {}), ...(value.baseItem?.trim() ? { baseItem: value.baseItem.trim() } : {}) };
}

/**
 * Excel 2021 (16.0.14334) портит память и падает, если менять подпись, формат
 * или вычисление поля значений, когда таких полей в сводной 4 и больше: замер
 * 07.10.2026 — десятки падений 0xc0000005/0xc0000374, при 2–3 полях ни одного.
 * Поэтому донастройку полей значений делаем только при 3 и меньше.
 */
export const MAX_TUNED_VALUE_FIELDS = 3;
export function valueFieldsTooManyText(count: number): string {
  return `В сводной будет ${count} полей значений, а Excel падает, если при ${MAX_TUNED_VALUE_FIELDS + 1} и больше менять их подписи, формат или вычисления (доля, нарастающий итог, ранг). ` +
    `Варианты: не больше ${MAX_TUNED_VALUE_FIELDS} полей значений с донастройкой; или все поля без label, numberFormat и showAs — а подписи и формат пользователь поставит сам: ` +
    "правый щелчок по заголовку поля → «Параметры полей значений». Скажи это пользователю прямо. Операция не выполнялась.";
}

/** Формат по умолчанию для долей и процентов — иначе 0,21 вместо 21%. */
export function defaultShowAsFormat(calculation: ShowAs): string | undefined {
  return /^percent/.test(calculation) ? "0.0%" : undefined;
}

export interface PivotFinish {
  values: readonly ValueFieldFinish[];
  grandTotals?: GrandTotals;
  /** false — убрать промежуточные итоги у всех полей строк, кроме последнего. */
  subtotals?: boolean;
  /** Поля, которые добавить в столбцы после сверки. */
  extraColumns?: readonly string[];
  /** Поля в область «Фильтры» сводной — после сверки: итогов не меняют, пока в них ничего не выбрано. */
  filterFields?: readonly string[];
}

/** Английская запись формата → запись культуры Excel: «,» и «.» вне кавычек и скобок. */
export function toLocalNumberFormat(format: string, decimalSeparator: string, groupSeparator: string): string {
  let out = "";
  let quoted = false;
  let bracket = false;
  for (let i = 0; i < format.length; i++) {
    const ch = format[i];
    if (ch === "\\" && !quoted && !bracket && i + 1 < format.length) { out += ch + format[++i]; continue; }
    if (ch === '"' && !bracket) quoted = !quoted;
    else if (ch === "[" && !quoted) bracket = true;
    else if (ch === "]" && !quoted) bracket = false;
    if (!quoted && !bracket && ch === ",") { out += groupSeparator; continue; }
    if (!quoted && !bracket && ch === ".") { out += decimalSeparator; continue; }
    out += ch;
  }
  return out;
}

/** Формат для диаграммы (оси, подписи): запись языка Excel, а пробел-разделитель
 * — неразрывный: с обычным деление на тысячу («90к») диаграмма не понимает
 * (замер 07.10.2026). */
export async function chartNumberFormat(ctx: Excel.RequestContext, format: string): Promise<string> {
  const separators = await culture(ctx);
  return toLocalNumberFormat(format, separators.decimal, separators.group === " " ? "\u00a0" : separators.group);
}

/** Подпись поля значений, которую Excel примет: совпадение с заголовком источника — с пробелом. */
export function acceptedValueLabel(label: string, sourceHeaders: readonly unknown[]): string {
  const clean = label.trim();
  const clash = sourceHeaders.some((header) => String(header ?? "").trim().toLowerCase() === clean.toLowerCase());
  return clash ? `${clean} ` : clean;
}

export function finishText(finish: PivotFinish, valueNames: readonly string[]): string[] {
  const lines: string[] = [];
  for (const item of finish.values) {
    const name = valueNames[item.index] ?? `поле ${item.index + 1}`;
    if (item.label) lines.push(`Поле значений «${name}» будет называться «${item.label.trim()}»`);
    if (item.showAs && item.showAs.calculation !== "none") {
      lines.push(`Поле «${item.label?.trim() || name}» показывает: ${SHOW_AS_TEXT[item.showAs.calculation]}${item.showAs.baseItem ? ` «${item.showAs.baseItem}»` : ""}${item.showAs.baseField ? ` (по полю «${item.showAs.baseField}»)` : ""}`);
    }
    if (item.numberFormat) lines.push(`Формат поля «${item.label?.trim() || name}»: ${item.numberFormat} — держится после обновления сводной`);
  }
  if (finish.grandTotals && finish.grandTotals !== "both") {
    lines.push({ none: "Без общих итогов", rows: "Общий итог только по строкам", columns: "Общий итог только по столбцам" }[finish.grandTotals as "none" | "rows" | "columns"]);
  }
  if (finish.subtotals === false) lines.push("Без промежуточных итогов");
  if (finish.extraColumns?.length) lines.push(`Ещё поле в столбцах: ${finish.extraColumns.map((item) => `«${item}»`).join(", ")}`);
  if (finish.filterFields?.length) lines.push(`Поля в «Фильтрах» сводной: ${finish.filterFields.map((item) => `«${item}»`).join(", ")} — у кнопки фильтра есть поиск`);
  return lines;
}

export async function culture(ctx: Excel.RequestContext): Promise<{ decimal: string; group: string }> {
  try {
    const format = (ctx.application as any).cultureInfo.numberFormat;
    format.load(["numberDecimalSeparator", "numberGroupSeparator"]);
    await ctx.sync();
    return { decimal: String(format.numberDecimalSeparator || ","), group: String(format.numberGroupSeparator || " ") };
  } catch {
    // Язык Excel неизвестен (ExcelApi ниже 1.11) — формат не переводится.
    return { decimal: ".", group: "," };
  }
}

/**
 * Применяет донастройку и читает её обратно. Возвращает перечень того, что
 * не совпало; пустой — всё встало. Автоширина при обновлении выключается
 * всегда: иначе каждое обновление заново раздвигает столбцы.
 */
export async function applyPivotFinish(ctx: Excel.RequestContext, pivot: Excel.PivotTable, finish: PivotFinish, sourceHeaders: readonly unknown[]): Promise<{ problems: string[]; applied: Record<string, unknown> }> {
  const problems: string[] = [];
  const layout = pivot.layout as any;
  try { layout.autoFormat = false; } catch { /* старый Excel */ }
  if (finish.grandTotals) {
    layout.showRowGrandTotals = finish.grandTotals === "both" || finish.grandTotals === "rows";
    layout.showColumnGrandTotals = finish.grandTotals === "both" || finish.grandTotals === "columns";
  }
  for (const name of finish.extraColumns ?? []) pivot.columnHierarchies.add(pivot.hierarchies.getItem(name));
  for (const name of finish.filterFields ?? []) pivot.filterHierarchies.add(pivot.hierarchies.getItem(name));
  await ctx.sync();

  if (finish.subtotals === false) {
    const rows = pivot.rowHierarchies;
    rows.load("items/name");
    await ctx.sync();
    // У последнего поля промежуточных итогов и так нет.
    for (const hierarchy of rows.items.slice(0, -1)) hierarchy.fields.getItem(hierarchy.name).subtotals = { Automatic: false } as any;
    await ctx.sync();
  }

  const data = pivot.dataHierarchies;
  data.load("items/name");
  await ctx.sync();
  if (finish.values.length && data.items.length > MAX_TUNED_VALUE_FIELDS) {
    // Страховка: проверка до Excel это уже отсекает.
    problems.push(`полей значений ${data.items.length} — подписи, формат и вычисления не ставились: Excel при этом падает`);
    finish = { ...finish, values: [] };
  }
  const separators = finish.values.some((item) => item.numberFormat) ? await culture(ctx) : null;
  const separatorsAll = finish.values.some((item) => item.numberFormat || (item.showAs && defaultShowAsFormat(item.showAs.calculation))) ? (separators ?? await culture(ctx)) : null;
  const wanted = finish.values.map((item) => {
    const format = item.numberFormat ?? (item.showAs ? defaultShowAsFormat(item.showAs.calculation) : undefined);
    return {
      index: item.index,
      label: item.label ? acceptedValueLabel(item.label, sourceHeaders) : undefined,
      numberFormat: format && separatorsAll ? toLocalNumberFormat(format, separatorsAll.decimal, separatorsAll.group) : undefined,
      showAs: item.showAs
    };
  });
  // Поле-основа вычисления — среди полей строк и столбцов.
  const axes = [pivot.rowHierarchies, pivot.columnHierarchies];
  if (wanted.some((item) => item.showAs?.baseField)) {
    for (const axis of axes) axis.load("items/name");
    await ctx.sync();
  }
  const baseFieldOf = (name: string) => {
    for (const axis of axes) {
      const found = axis.items.find((item) => item.name.trim().toLowerCase() === name.trim().toLowerCase());
      if (found) return found.fields.getItem(found.name);
    }
    return null;
  };
  for (const item of wanted) {
    if (!item.showAs) continue;
    const hierarchy = data.items[item.index];
    if (!hierarchy) continue;
    const rule: any = { calculation: OFFICE_SHOW_AS[item.showAs.calculation] };
    if (item.showAs.baseField) {
      const field = baseFieldOf(item.showAs.baseField);
      if (!field) { problems.push(`поля «${item.showAs.baseField}» нет в строках и столбцах сводной — вычисление не поставлено`); continue; }
      rule.baseField = field;
      if (item.showAs.baseItem) {
        const items = field.items;
        items.load("items/name");
        await ctx.sync();
        const base = items.items.find((element) => element.name.trim().toLowerCase() === item.showAs!.baseItem!.toLowerCase());
        if (!base) { problems.push(`в поле «${item.showAs.baseField}» нет элемента «${item.showAs.baseItem}»; есть: ${items.items.slice(0, 12).map((element) => `«${element.name}»`).join(", ")}`); continue; }
        rule.baseItem = base;
      }
    }
    hierarchy.showAs = rule;
  }
  await ctx.sync();
  for (const item of wanted) {
    const hierarchy = data.items[item.index];
    if (!hierarchy) { problems.push(`поля значений №${item.index + 1} нет`); continue; }
    if (item.numberFormat) hierarchy.numberFormat = item.numberFormat;
    if (item.label) hierarchy.name = item.label;
  }
  await ctx.sync();

  // Обратное чтение: то ли встало, что просили.
  const after = pivot.dataHierarchies;
  after.load("items/name,items/numberFormat,items/showAs");
  layout.load(["showRowGrandTotals", "showColumnGrandTotals", "autoFormat"]);
  const columns = pivot.columnHierarchies;
  columns.load("items/name");
  await ctx.sync();
  for (const item of wanted) {
    const got = after.items[item.index];
    if (!got) continue;
    if (item.label && got.name !== item.label) problems.push(`поле значений названо «${got.name}» вместо «${item.label.trim()}»`);
    if (item.numberFormat && got.numberFormat !== item.numberFormat) problems.push(`формат поля «${got.name}» — «${got.numberFormat}» вместо «${item.numberFormat}»`);
    if (item.showAs && String((got as any).showAs?.calculation ?? "") !== OFFICE_SHOW_AS[item.showAs.calculation]) {
      problems.push(`поле «${got.name}» считает «${(got as any).showAs?.calculation}» вместо «${SHOW_AS_TEXT[item.showAs.calculation]}»`);
    }
  }
  if (finish.grandTotals) {
    const rows = finish.grandTotals === "both" || finish.grandTotals === "rows";
    const cols = finish.grandTotals === "both" || finish.grandTotals === "columns";
    if (layout.showRowGrandTotals !== rows || layout.showColumnGrandTotals !== cols) problems.push("общие итоги не переключились");
  }
  for (const name of finish.extraColumns ?? []) {
    if (!columns.items.some((item) => item.name === name)) problems.push(`поле «${name}» не встало в столбцы`);
  }
  let pageFields: string[] = [];
  if (finish.filterFields?.length) {
    const filters = pivot.filterHierarchies;
    filters.load("items/name");
    await ctx.sync();
    pageFields = filters.items.map((item) => item.name);
    for (const name of finish.filterFields) {
      if (!pageFields.some((item) => item.trim().toLowerCase() === name.trim().toLowerCase())) problems.push(`поле «${name}» не встало в «Фильтры» сводной`);
    }
  }
  const spaced = wanted.filter((item) => item.label && item.label.endsWith(" ")).map((item) => item.label!.trim());
  return {
    problems,
    applied: {
      ...(spaced.length ? {
        labelNote: `Подпись ${spaced.map((name) => `«${name}»`).join(", ")} Excel хранит с пробелом в конце: имя, совпадающее с полем источника, он иначе не принимает. ` +
          "Это не ошибка, поле названо как просили — не переименовывай его."
      } : {}),
      valueFields: after.items.map((item) => ({ name: item.name, numberFormat: item.numberFormat, showAs: String((item as any).showAs?.calculation ?? "None") })),
      grandTotals: { rows: layout.showRowGrandTotals, columns: layout.showColumnGrandTotals },
      autoFitOnRefresh: layout.autoFormat,
      columns: columns.items.map((item) => item.name).filter((name) => name !== "Значения" && name !== "Values"),
      ...(finish.filterFields?.length ? { filterFields: pageFields } : {})
    }
  };
}
