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
export interface ValueFieldFinish {
  /** Номер поля значений по порядку добавления. */
  index: number;
  label?: string;
  numberFormat?: string;
}

export interface PivotFinish {
  values: readonly ValueFieldFinish[];
  grandTotals?: GrandTotals;
  /** false — убрать промежуточные итоги у всех полей строк, кроме последнего. */
  subtotals?: boolean;
  /** Поля, которые добавить в столбцы после сверки. */
  extraColumns?: readonly string[];
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
    if (item.numberFormat) lines.push(`Формат поля «${item.label?.trim() || name}»: ${item.numberFormat} — держится после обновления сводной`);
  }
  if (finish.grandTotals && finish.grandTotals !== "both") {
    lines.push({ none: "Без общих итогов", rows: "Общий итог только по строкам", columns: "Общий итог только по столбцам" }[finish.grandTotals as "none" | "rows" | "columns"]);
  }
  if (finish.subtotals === false) lines.push("Без промежуточных итогов");
  if (finish.extraColumns?.length) lines.push(`Ещё поле в столбцах: ${finish.extraColumns.map((item) => `«${item}»`).join(", ")}`);
  return lines;
}

async function culture(ctx: Excel.RequestContext): Promise<{ decimal: string; group: string }> {
  try {
    const format = (ctx.application as any).cultureInfo.numberFormat;
    format.load(["numberDecimalSeparator", "numberGroupSeparator"]);
    await ctx.sync();
    return { decimal: String(format.numberDecimalSeparator || ","), group: String(format.numberGroupSeparator || " ") };
  } catch {
    return { decimal: ",", group: " " };
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
  const separators = finish.values.some((item) => item.numberFormat) ? await culture(ctx) : null;
  const wanted = finish.values.map((item) => ({
    index: item.index,
    label: item.label ? acceptedValueLabel(item.label, sourceHeaders) : undefined,
    numberFormat: item.numberFormat && separators ? toLocalNumberFormat(item.numberFormat, separators.decimal, separators.group) : undefined
  }));
  for (const item of wanted) {
    const hierarchy = data.items[item.index];
    if (!hierarchy) { problems.push(`поля значений №${item.index + 1} нет`); continue; }
    if (item.numberFormat) hierarchy.numberFormat = item.numberFormat;
    if (item.label) hierarchy.name = item.label;
  }
  await ctx.sync();

  // Обратное чтение: то ли встало, что просили.
  const after = pivot.dataHierarchies;
  after.load("items/name,items/numberFormat");
  layout.load(["showRowGrandTotals", "showColumnGrandTotals", "autoFormat"]);
  const columns = pivot.columnHierarchies;
  columns.load("items/name");
  await ctx.sync();
  for (const item of wanted) {
    const got = after.items[item.index];
    if (!got) continue;
    if (item.label && got.name !== item.label) problems.push(`поле значений названо «${got.name}» вместо «${item.label.trim()}»`);
    if (item.numberFormat && got.numberFormat !== item.numberFormat) problems.push(`формат поля «${got.name}» — «${got.numberFormat}» вместо «${item.numberFormat}»`);
  }
  if (finish.grandTotals) {
    const rows = finish.grandTotals === "both" || finish.grandTotals === "rows";
    const cols = finish.grandTotals === "both" || finish.grandTotals === "columns";
    if (layout.showRowGrandTotals !== rows || layout.showColumnGrandTotals !== cols) problems.push("общие итоги не переключились");
  }
  for (const name of finish.extraColumns ?? []) {
    if (!columns.items.some((item) => item.name === name)) problems.push(`поле «${name}» не встало в столбцы`);
  }
  const spaced = wanted.filter((item) => item.label && item.label.endsWith(" ")).map((item) => item.label!.trim());
  return {
    problems,
    applied: {
      ...(spaced.length ? {
        labelNote: `Подпись ${spaced.map((name) => `«${name}»`).join(", ")} Excel хранит с пробелом в конце: имя, совпадающее с полем источника, он иначе не принимает. ` +
          "Это не ошибка, поле названо как просили — не переименовывай его."
      } : {}),
      valueFields: after.items.map((item) => ({ name: item.name, numberFormat: item.numberFormat })),
      grandTotals: { rows: layout.showRowGrandTotals, columns: layout.showColumnGrandTotals },
      autoFitOnRefresh: layout.autoFormat,
      columns: columns.items.map((item) => item.name).filter((name) => name !== "Значения" && name !== "Values")
    }
  };
}
