/**
 * Сводная таблица на проверяемом пути.
 *
 * Подготовка считает сводную сама (`pivotModel.ts`): группы, итоги, общий
 * итог и размер. По размеру проверяется, что место под сводной пусто, —
 * иначе Excel молча затёр бы данные или отказал на середине. Исполнение
 * строит сводную и сверяет прочитанные из неё числа с расчётом панели:
 * расхождение значит, что Excel свёл не то, что ожидалось, и об этом
 * говорится, а не молчится.
 */

import { checkSheetName, freeSheetName } from "./sheetRules";
import {
  assertPlanWorkbook,
  checkAddress,
  deepFreeze,
  MAX_IO_CELLS,
  preflightToolArgs,
  probeMergedAreas,
  rangeOf,
  ToolError,
  ToolExecutionError
} from "./excelTools";
import {
  AGGREGATIONS,
  AGGREGATION_TEXT,
  expectPivot,
  fieldIndex,
  OFFICE_AGGREGATION,
  pivotHeaderProblems,
  pivotMismatches,
  pivotKey,
  pivotLabel,
  dateGroupFormula,
  dateGroupLabel,
  isDateFormat,
  DATE_GROUPINGS,
  DATE_GROUPING_TEXT,
  type DateGrouping,
  type Aggregation,
  type PivotExpectation,
  type PivotFilter,
  type PivotOptions,
  type PivotValueField
} from "./pivotModel";
import { intersects, parseA1Rect, type A1Rect } from "./a1";
import { columnLetters } from "./formulaFill";
import { placementCell } from "./chartModel";
import { action, getStructuralRevision, isCustomUndoAvailable, push } from "./undo";
import { captureTarget, officeCapabilities, type WorkbookTarget } from "./workbookContext";
import { applyPivotFinish, checkShowAs, MAX_TUNED_VALUE_FIELDS, valueFieldsTooManyText, finishText, GRAND_TOTALS, type GrandTotals, type PivotFinish } from "./pivotFinish";

export interface CreatePivotPlan {
  readonly kind: "create_pivot_table";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly name: string;
  readonly sourceAddress: string;
  readonly sourceRows: number;
  readonly rowFields: readonly string[];
  readonly valueFields: readonly PivotValueField[];
  /** Поле в столбцах — заголовок, как в источнике (8.2). */
  readonly columnField?: string;
  /** Фильтры: имена полей как в источнике, элементы — подписями Excel (8.2). */
  readonly filters: readonly PlannedPivotFilter[];
  /** Порядок элементов поля строк по итогу поля значений `by` (8.2). */
  readonly sort?: { readonly field: string; readonly by: number; readonly order: "asc" | "desc" };
  /**
   * Группировка дат (8.2): вспомогательные столбцы правее источника с
   * формулами года, квартала или месяца. Сводная строится по источнику
   * вместе с ними; отмена убирает их вместе со сводной.
   */
  readonly dateGroups?: DateGroupPlan;
  /** Область, по которой строится сводная: источник и вспомогательные столбцы. */
  readonly pivotSourceAddress: string;
  readonly destSheet: string;
  /** Пусто, если лист создаётся самой операцией (newSheet). */
  readonly destSheetId: string;
  /** Лист создаётся при исполнении (этап 7, 7.3.4); отмена уберёт его, если он останется пустым. */
  readonly newSheet?: true;
  /** Левый верхний угол и вся область, которую займёт сводная. */
  readonly destCell: string;
  readonly destArea: string;
  readonly expectation: PivotExpectation;
  readonly preview: readonly string[];
  /** Слепок источника — формулы и значения: по ним посчитан расчёт панели. */
  readonly signature: string;
  readonly sourceSameSheet: boolean;
  /** Источник — таблица Excel целиком: сводная строится по её имени и после
   * обновления подхватывает новые строки (срез 10.7). */
  readonly sourceTable?: string;
  /** Заголовки источника — для подписи поля значений, совпадающей с ними. */
  readonly sourceHeaders: readonly string[];
  /** Донастройка после сверки (10.7): подписи, формат, итоги, ещё поля в столбцах. */
  readonly finish: PivotFinish;
  readonly undoAvailable: boolean;
  readonly undoNote?: string;
  readonly createdAt: string;
}

/** Донастройка из аргументов модели: подписи и формат полей, итоги, второе поле в столбцах. */
function parseFinish(a: { values: unknown[]; grandTotals?: unknown; subtotals?: unknown; columns?: unknown; filterFields?: unknown }, headers: readonly unknown[]): PivotFinish {
  const values = (Array.isArray(a.values) ? a.values : []).flatMap((raw: any, index) => {
    const label = typeof raw?.label === "string" && raw.label.trim() ? raw.label.trim() : undefined;
    const numberFormat = typeof raw?.numberFormat === "string" && raw.numberFormat.trim() ? raw.numberFormat.trim() : undefined;
    let showAs;
    try { showAs = checkShowAs(raw?.showAs); } catch (error: any) { throw new ToolError(error.message); }
    return label || numberFormat || showAs ? [{ index, ...(label ? { label } : {}), ...(numberFormat ? { numberFormat } : {}), ...(showAs ? { showAs } : {}) }] : [];
  });
  const count = Array.isArray(a.values) ? a.values.length : 0;
  if (values.length && count > MAX_TUNED_VALUE_FIELDS) throw new ToolError(valueFieldsTooManyText(count));
  const labels = values.map((item) => item.label?.toLowerCase()).filter(Boolean);
  if (new Set(labels).size !== labels.length) throw new ToolError("Подписи полей значений повторяются: у каждого поля своя.");
  if (a.grandTotals !== undefined && !GRAND_TOTALS.includes(a.grandTotals as GrandTotals)) {
    throw new ToolError(`grandTotals — ${GRAND_TOTALS.join(", ")}.`);
  }
  const columns = Array.isArray(a.columns) ? a.columns.map(String) : [];
  const extraColumns = columns.slice(1);
  const missing = extraColumns.filter((name) => fieldIndex(headers, name) < 0);
  if (missing.length) throw new ToolError(`Нет полей ${missing.map((name) => `«${name}»`).join(", ")} для столбцов сводной.`);
  const filterFields = Array.isArray(a.filterFields) ? a.filterFields.map((name) => String(name).trim()).filter(Boolean) : [];
  const unknownFilter = filterFields.filter((name) => fieldIndex(headers, name) < 0);
  if (unknownFilter.length) {
    throw new ToolError(`Нет полей ${unknownFilter.map((name) => `«${name}»`).join(", ")} для «Фильтров» сводной. Заголовки источника: ${headers.map((item) => `«${item}»`).join(", ")}.`);
  }
  const busy = filterFields.filter((name) => [...(Array.isArray(a.columns) ? a.columns : []), ...((a as any).rows ?? [])].some((other: unknown) => String(other).trim().toLowerCase() === name.toLowerCase()));
  if (busy.length) throw new ToolError(`Поле ${busy.map((name) => `«${name}»`).join(", ")} уже в строках или столбцах — в «Фильтры» его не поставить.`);
  return {
    ...(filterFields.length ? { filterFields: filterFields.map((name) => String(headers[fieldIndex(headers, name)])) } : {}),
    values,
    ...(a.grandTotals !== undefined ? { grandTotals: a.grandTotals as GrandTotals } : {}),
    ...(a.subtotals === false ? { subtotals: false } : {}),
    ...(extraColumns.length ? { extraColumns } : {})
  };
}

export type PlannedPivotFilter =
  | { readonly field: string; readonly axis: "rows" | "columns"; readonly include: readonly string[] }
  | { readonly field: string; readonly axis: "rows" | "columns"; readonly top: number; readonly by: number }
  | { readonly field: string; readonly axis: "rows" | "columns"; readonly bottom: number; readonly by: number };

/** Фильтры и порядок сводной (8.2) — ExcelApi 1.12 и 1.9. */
const PIVOT_FILTER_API = "1.12";
const PIVOT_SORT_API = "1.9";

function apiSupported(version: string): boolean {
  try {
    return Boolean((globalThis as any).Office?.context?.requirements?.isSetSupported?.("ExcelApi", version));
  } catch {
    return false;
  }
}

interface RawPivotArgs {
  columns?: unknown;
  filters?: unknown;
  sort?: unknown;
  groupDates?: unknown;
}

export interface DateGroupPlan {
  readonly field: string;
  readonly levels: readonly DateGrouping[];
  /** Заголовки вспомогательных столбцов, например «Дата (месяц)». */
  readonly names: readonly string[];
  /** Где они встанут — вместе со строкой заголовков. */
  readonly address: string;
  /** Что туда записать: заголовки и формулы. */
  readonly formulas: readonly (readonly string[])[];
  /** Что формулы обязаны дать — посчитано панелью по тем же датам. */
  readonly expected: readonly (readonly unknown[])[];
}

const isBlankCell = (value: unknown) => value === "" || value === null || value === undefined;
const sameName = (x: string, y: string) => x.trim().toLowerCase() === y.trim().toLowerCase();

/**
 * Группировка дат (8.2). Замер 27.09.2026: поле дат в сводной Excel сам не
 * группирует — каждая дата отдельной строкой. Поэтому правее источника
 * встают вспомогательные столбцы с формулами, и сводная строится уже по ним:
 * на место поля дат в rows или columns встают «Дата (квартал)», «Дата (месяц)».
 */
async function planDateGroups(
  ctx: Excel.RequestContext,
  sheet: Excel.Worksheet,
  range: Excel.Range,
  values: readonly (readonly unknown[])[],
  a: { rows: string[]; columns?: unknown; groupDates?: unknown }
): Promise<{ plan: DateGroupPlan; values: unknown[][]; rows: string[]; columns?: string[]; rect: A1Rect }> {
  const headers = values[0];
  const raw = a.groupDates as { field?: unknown; by?: unknown };
  const name = typeof raw?.field === "string" ? raw.field : "";
  const column = fieldIndex(headers, name);
  if (column < 0) throw new ToolError(`groupDates: поля «${name}» нет. Заголовки источника: ${headers.map((value) => `«${String(value)}»`).join(", ")}.`);
  const field = String(headers[column]);
  const asked = (Array.isArray(raw.by) ? raw.by : [raw.by]) as unknown[];
  if (!asked.length || asked.some((level) => !DATE_GROUPINGS.includes(level as DateGrouping))) {
    throw new ToolError("groupDates.by — month, quarter или year, или список из них, например [\"quarter\", \"month\"].");
  }
  const levels = DATE_GROUPINGS.filter((level) => asked.includes(level));
  const columnsAsked = Array.isArray(a.columns) ? a.columns.map(String) : [];
  const inRows = a.rows.some((row) => sameName(row, field));
  const inColumns = columnsAsked.some((item) => sameName(item, field));
  if (!inRows && !inColumns) {
    throw new ToolError(`groupDates по «${field}»: поставьте это поле в rows или columns — на его место встанут ${levels.map((level) => DATE_GROUPING_TEXT[level]).join(" и ")}.`);
  }
  if (inColumns && levels.length > 1) throw new ToolError("В столбцах сводной — одно поле: для дат в columns выберите один уровень группировки.");

  const body = values.slice(1);
  const formats = sheet.getRangeByIndexes(range.rowIndex + 1, range.columnIndex + column, range.rowCount - 1, 1);
  formats.load("numberFormat");
  await ctx.sync();
  const numberFormats = formats.numberFormat as unknown[][];
  const texts: string[] = [];
  const numbers: number[] = [];
  body.forEach((row, index) => {
    const value = row[column];
    if (isBlankCell(value)) return;
    if (typeof value !== "number") texts.push(String(value));
    // До 1 марта 1900 года Excel считает несуществующее 29 февраля — такие даты не группируем.
    else if (value < 61 || value > 2958465 || !isDateFormat(numberFormats[index]?.[0])) numbers.push(value);
  });
  if (texts.length) {
    throw new ToolError(
      `В поле «${field}» ${texts.length} значений записаны текстом, а не датой (например «${texts[0]}»): Excel не посчитает по ним месяц. ` +
      "Сначала преобразуйте их в даты (convert_values), потом стройте сводную."
    );
  }
  if (numbers.length) throw new ToolError(`В поле «${field}» числа без формата даты (например ${numbers[0]}): группировать их по месяцам нельзя.`);

  const names = levels.map((level) => `${field} (${DATE_GROUPING_TEXT[level]})`);
  const taken = names.filter((item) => fieldIndex(headers, item) >= 0);
  if (taken.length) throw new ToolError(`В источнике уже есть ${taken.map((item) => `«${item}»`).join(", ")}: стройте сводную по нему, без groupDates.`);

  const helper = sheet.getRangeByIndexes(range.rowIndex, range.columnIndex + range.columnCount, range.rowCount, levels.length);
  helper.load("address");
  await ctx.sync();
  const address = withoutSheet(String(helper.address));
  const rect = parseA1Rect(address)!;
  const check = await checkDestination(ctx, sheet, { rect, address }, null);
  if (check.problem) throw new ToolError(`Вспомогательные столбцы для дат встанут в ${address}: ${check.problem.replace(/^Сводная займёт [^ ]+ /, "")}`);
  if (check.occupied) {
    throw new ToolError(
      `Вспомогательные столбцы для дат встают сразу правее источника, в ${sheet.name}!${address}, а там ${check.occupied} непустых ячеек. ` +
      "Операция не выполнялась: освободите эти столбцы или добавьте месяц в источник сами."
    );
  }

  const dateColumn = columnLetters(range.columnIndex + column + 1);
  const formulas = [names, ...body.map((_, index) => levels.map((level) => dateGroupFormula(`${dateColumn}${range.rowIndex + 2 + index}`, level)))];
  const expected = [names, ...body.map((row) => levels.map((level) => (isBlankCell(row[column]) ? "" : dateGroupLabel(row[column] as number, level))))];
  const extended = values.map((row, index) => [...row, ...expected[index]]);
  const rows = a.rows.flatMap((row) => (sameName(row, field) ? names : [row]));
  const columns = inColumns ? [names[0]] : columnsAsked.length ? columnsAsked : undefined;
  const sourceRect = parseA1Rect(withoutSheet(String(range.address)))!;
  return {
    plan: { field, levels, names, address, formulas, expected },
    values: extended,
    rows,
    ...(columns ? { columns } : {}),
    rect: { ...sourceRect, columnEnd: rect.columnEnd }
  };
}

/** Убирает вспомогательные столбцы дат; true — если их больше нет. */
async function clearDateHelpers(ctx: Excel.RequestContext, sheet: Excel.Worksheet, plan: DateGroupPlan): Promise<boolean> {
  try {
    const helper = sheet.getRange(plan.address);
    helper.clear("Contents" as any);
    await ctx.sync();
    helper.load("formulas");
    await ctx.sync();
    return (helper.formulas as unknown[][]).every((row) => row.every(isBlankCell));
  } catch {
    return false;
  }
}

/**
 * Поле в столбцах, фильтры и порядок (8.2): проверка по шапке и данным
 * источника. Всё, что Excel сделал бы не так, как просили, или молча
 * пропустил, отклоняется здесь — до карточки.
 */
function resolvePivotOptions(
  values: readonly (readonly unknown[])[],
  rows: readonly string[],
  valueFields: readonly PivotValueField[],
  raw: RawPivotArgs
): { options: PivotOptions; columnField?: string; filters: PlannedPivotFilter[]; sort?: { field: string; by: number; order: "asc" | "desc" } } {
  const headers = values[0];
  const header = (name: string) => String(headers[fieldIndex(headers, name)]);
  const same = (x: string, y: string) => x.trim().toLowerCase() === y.trim().toLowerCase();
  const byIndex = (name: unknown, what: string) => {
    if (name === undefined) return 0;
    const index = valueFields.findIndex((item) => same(item.field, String(name)));
    if (index < 0) throw new ToolError(`${what}: поля «${String(name)}» нет в values. Поля значений: ${valueFields.map((item) => `«${item.field}»`).join(", ")}.`);
    return index;
  };

  let columnField: string | undefined;
  if (raw.columns !== undefined) {
    if (!Array.isArray(raw.columns) || raw.columns.length > 1) throw new ToolError("В columns — одно поле: вложенные поля в столбцах панель пока не строит.");
    const name = raw.columns[0];
    if (typeof name === "string" && name.trim()) {
      if (fieldIndex(headers, name) < 0) {
        throw new ToolError(`Нет поля «${name}» для columns. Заголовки источника: ${headers.map((value) => `«${String(value)}»`).join(", ")}.`);
      }
      if (rows.some((row) => same(row, name))) throw new ToolError(`Поле «${name}» уже в rows: одно поле не может быть и строками, и столбцами.`);
      columnField = header(name);
    }
  }

  const filters: PlannedPivotFilter[] = [];
  const modelFilters: PivotFilter[] = [];
  if (raw.filters !== undefined) {
    if (!Array.isArray(raw.filters)) throw new ToolError("filters — список фильтров полей.");
    if (raw.filters.length && !apiSupported(PIVOT_FILTER_API)) {
      throw new ToolError(`Фильтры сводной требуют ExcelApi ${PIVOT_FILTER_API} (Office 2021 и новее), а этот Excel его не поддерживает. Постройте сводную без filters.`);
    }
    for (const item of raw.filters as { field?: unknown; include?: unknown; top?: unknown; bottom?: unknown; by?: unknown }[]) {
      const name = typeof item?.field === "string" ? item.field : "";
      const onRows = rows.some((row) => same(row, name));
      const onColumns = columnField !== undefined && same(columnField, name);
      if (!onRows && !onColumns) {
        throw new ToolError(`Фильтр по «${name}»: такого поля нет ни в rows, ни в columns — filters отбирает значения полей сводной. ` +
          `Чтобы поставить «${name}» в область «Фильтры» сводной (кнопка с поиском над ней), передай filterFields: ["${name}"].`);
      }
      const field = header(name);
      if (filters.some((other) => same(other.field, field))) throw new ToolError(`По полю «${field}» два фильтра: у поля сводной фильтр один.`);
      const kinds = ["include", "top", "bottom"].filter((key) => (item as any)[key] !== undefined);
      if (kinds.length !== 1) throw new ToolError(`Фильтр по «${field}»: нужно ровно одно из include, top, bottom.`);
      const axis = onColumns ? "columns" as const : "rows" as const;
      if (item.include !== undefined) {
        if (!Array.isArray(item.include) || !item.include.length) throw new ToolError(`Фильтр по «${field}»: include — непустой список значений.`);
        const column = fieldIndex(headers, field);
        const labels = new Map<string, string>();
        for (const row of values.slice(1)) {
          const cell = row[column];
          if (cell === "" || cell === null || cell === undefined) continue;
          const key = pivotKey(cell);
          if (!labels.has(key)) labels.set(key, pivotLabel(cell));
        }
        const unknown = (item.include as unknown[]).filter((value) => !labels.has(pivotKey(value)));
        if (unknown.length) {
          throw new ToolError(
            `Фильтр по «${field}»: значений ${unknown.map((value) => `«${String(value)}»`).join(", ")} в данных нет. ` +
            `Есть: ${[...labels.values()].slice(0, 20).map((label) => `«${label}»`).join(", ")}${labels.size > 20 ? " и другие" : ""}.`
          );
        }
        const include = [...new Set((item.include as unknown[]).map((value) => labels.get(pivotKey(value))!))];
        filters.push({ field, axis, include });
        modelFilters.push({ field, include });
      } else {
        if (onRows && !same(rows[0], field)) {
          throw new ToolError(`Первые/последние N — только для первого поля rows («${rows[0]}») или поля columns: для внутренних полей Excel отбирает N внутри каждой группы, и такой расчёт панель пока не делает.`);
        }
        const count = Number(item.top ?? item.bottom);
        if (!Number.isInteger(count) || count < 1) throw new ToolError(`Фильтр по «${field}»: N — целое число от 1.`);
        const by = byIndex(item.by, `Фильтр по «${field}»`);
        if (item.top !== undefined) { filters.push({ field, axis, top: count, by }); modelFilters.push({ field, top: count, by }); }
        else { filters.push({ field, axis, bottom: count, by }); modelFilters.push({ field, bottom: count, by }); }
      }
    }
  }

  let sort: { field: string; by: number; order: "asc" | "desc" } | undefined;
  if (raw.sort !== undefined) {
    const item = raw.sort as { field?: unknown; by?: unknown; order?: unknown };
    const name = typeof item?.field === "string" ? item.field : "";
    if (!rows.some((row) => same(row, name))) throw new ToolError(`Порядок по «${name}»: такого поля нет в rows. Упорядочить можно только поле строк.`);
    if (item.order !== "asc" && item.order !== "desc") throw new ToolError("sort.order — desc или asc.");
    if (!apiSupported(PIVOT_SORT_API)) throw new ToolError(`Порядок по значению требует ExcelApi ${PIVOT_SORT_API}, а этот Excel его не поддерживает.`);
    sort = { field: header(name), by: byIndex(item.by, `Порядок по «${name}»`), order: item.order };
  }

  return {
    options: { ...(columnField ? { columnField } : {}), ...(modelFilters.length ? { filters: modelFilters } : {}), ...(sort ? { sort } : {}) },
    ...(columnField ? { columnField } : {}),
    filters,
    ...(sort ? { sort } : {})
  };
}

/** Строки предпросмотра про столбцы, фильтры и порядок. */
function optionsPreview(expectation: PivotExpectation, resolved: ReturnType<typeof resolvePivotOptions>, valueFields: readonly PivotValueField[]): string[] {
  const lines: string[] = [];
  if (resolved.columnField) {
    // Excel ставит элементы по алфавиту — так же их и называем.
    const labels = expectation.columns.map((item) => item.label).sort((x, y) => x.localeCompare(y, "ru"));
    lines.push(`Столбцы по «${resolved.columnField}»: ${labels.slice(0, 8).join(", ")}${labels.length > 8 ? ` и ещё ${labels.length - 8}` : ""}`);
  }
  for (const filter of resolved.filters) {
    const out = expectation.filteredOut.find((item) => item.field === filter.field)?.items ?? [];
    const what = "include" in filter
      ? `только ${filter.include.map((label) => `«${label}»`).join(", ")}`
      : `${"top" in filter ? "первые" : "последние"} ${"top" in filter ? filter.top : filter.bottom} по «${valueFields[filter.by].field}»`;
    lines.push(`Фильтр «${filter.field}»: ${what}${out.length ? `; уберутся ${out.slice(0, 6).map((label) => `«${label}»`).join(", ")}${out.length > 6 ? ` и ещё ${out.length - 6}` : ""}` : ""}`);
  }
  if (resolved.sort) {
    lines.push(`Порядок «${resolved.sort.field}»: по «${valueFields[resolved.sort.by].field}», ${resolved.sort.order === "desc" ? "от большего к меньшему" : "от меньшего к большему"}`);
  }
  return lines;
}

/** Место, которое должно быть пустым: итоговая сводная и она же до фильтров. */
function placeArea(cell: string, expectation: PivotExpectation) {
  return areaAt(cell, Math.max(expectation.height, expectation.fullHeight), Math.max(expectation.width, expectation.fullWidth));
}

function withoutSheet(address: string): string {
  return address.slice(address.lastIndexOf("!") + 1);
}

function areaAt(cell: string, height: number, width: number): { rect: A1Rect; address: string } {
  const start = parseA1Rect(cell)!;
  const rect: A1Rect = {
    kind: "cells",
    rowStart: start.rowStart,
    columnStart: start.columnStart,
    rowEnd: start.rowStart + height - 1,
    columnEnd: start.columnStart + width - 1
  };
  const address = `${columnLetters(rect.columnStart)}${rect.rowStart}:${columnLetters(rect.columnEnd)}${rect.rowEnd}`;
  return { rect, address };
}

function parseValueFields(raw: unknown): PivotValueField[] {
  if (!Array.isArray(raw) || !raw.length) throw new ToolError("Нужно хотя бы одно поле в values.");
  return raw.map((item) => {
    if (typeof item === "string") return { field: item, aggregation: "sum" as Aggregation };
    const value = item as { field?: unknown; aggregation?: unknown };
    if (typeof value?.field !== "string" || !value.field.trim()) throw new ToolError("У каждого поля значений нужен field — заголовок столбца.");
    const aggregation = (value.aggregation ?? "sum") as Aggregation;
    if (!AGGREGATIONS.includes(aggregation)) throw new ToolError(`Неизвестная агрегация ${String(value.aggregation)}: доступны ${AGGREGATIONS.join(", ")}.`);
    return { field: value.field, aggregation };
  });
}

/**
 * Сколько ячеек места заняты.
 *
 * Занятость смотрится и по формулам, и по значениям. Формула `=""` даёт
 * пустое значение, и проверка по одним значениям принимала её за свободную
 * ячейку, хотя сводная её затёрла бы (план стабилизации, S2).
 */
function occupiedCells(formulas: unknown[][], values: unknown[][]): number {
  let count = 0;
  formulas.forEach((row, r) => row.forEach((formula, c) => {
    const value = values[r]?.[c];
    if ((formula !== "" && formula !== null && formula !== undefined) || (value !== "" && value !== null && value !== undefined)) count += 1;
  }));
  return count;
}

/**
 * Таблицы листа — без проглатывания ошибки.
 *
 * Общий `readTableRanges` при сбое отдаёт пустой список: для предупреждений
 * этого хватает. Но для места сводной «не смогли прочитать» — не «таблиц нет»,
 * и строить поверх непроверенного места нельзя.
 */
async function readTablesStrict(ctx: Excel.RequestContext, sheet: Excel.Worksheet): Promise<{ name: string; address: string }[]> {
  try {
    const tables = sheet.tables;
    tables.load("items/name");
    await ctx.sync();
    const ranges = tables.items.map((table) => {
      const range = table.getRange();
      range.load("address");
      return { name: table.name, range };
    });
    if (ranges.length) await ctx.sync();
    return ranges.map((item) => ({ name: item.name, address: String(item.range.address) }));
  } catch (error: any) {
    throw new ToolError(`Не удалось прочитать таблицы листа ${sheet.name}: ${error?.message ?? error}. Место под сводной не проверено, и строить на нём нельзя.`);
  }
}

interface DestinationCheck {
  /** Причина, по которой место не годится вовсе; null — годится, если пусто. */
  problem: string | null;
  occupied: number;
}

/**
 * Одно правило места для всех случаев: подготовки, повторной проверки перед
 * созданием и поиска свободного места. Иначе подсказанное «свободное» место
 * могло бы не пройти следующую же проверку.
 */
async function checkDestination(
  ctx: Excel.RequestContext,
  sheet: Excel.Worksheet,
  area: { rect: A1Rect; address: string },
  sourceRect: A1Rect | null
): Promise<DestinationCheck> {
  sheet.load("name");
  sheet.protection.load("protected");
  const place = sheet.getRange(area.address);
  place.load(["address", "rowIndex", "columnIndex", "rowCount", "columnCount", "formulas", "values"]);
  await ctx.sync();
  const occupied = occupiedCells(place.formulas as unknown[][], place.values as unknown[][]);
  const refuse = (problem: string) => ({ problem, occupied });

  if (sheet.protection.protected) return refuse(`Лист ${sheet.name} защищён: сводную на нём не построить. Снимите защиту или выберите другой лист.`);
  if (sourceRect && intersects(area.rect, sourceRect)) return refuse(`Сводная займёт ${area.address} и наложится на источник.`);

  const tables = await readTablesStrict(ctx, sheet);
  const hitTable = tables.find((table) => {
    const rect = parseA1Rect(withoutSheet(table.address));
    return rect && intersects(rect, area.rect);
  });
  if (hitTable) return refuse(`Сводная займёт ${area.address} и заденет таблицу ${hitTable.name} (${hitTable.address}).`);

  const pivots = sheet.pivotTables;
  pivots.load("items/name");
  await ctx.sync();
  const pivotRanges = pivots.items.map((item) => {
    const layoutRange = item.layout.getRange();
    layoutRange.load("address");
    return { name: item.name, range: layoutRange };
  });
  if (pivotRanges.length) await ctx.sync();
  const hitPivot = pivotRanges.find((item) => {
    const rect = parseA1Rect(withoutSheet(String(item.range.address)));
    return rect && intersects(rect, area.rect);
  });
  if (hitPivot) return refuse(`Сводная займёт ${area.address} и заденет сводную ${hitPivot.name}.`);

  // Excel не строит сводную поверх объединённых ячеек. Угол объединения
  // с неизвестными границами внутри места — тоже отказ: доказать, что
  // объединение не заходит в место, нельзя.
  const merged = await probeMergedAreas(ctx, sheet, place);
  const hitMerged = merged.areas.find((address) => {
    const rect = parseA1Rect(withoutSheet(address));
    return rect && intersects(rect, area.rect);
  }) ?? merged.unresolvedAnchors.find((address) => {
    const rect = parseA1Rect(withoutSheet(address));
    return rect && intersects(rect, area.rect);
  });
  if (hitMerged) return refuse(`Сводная займёт ${area.address}, а там объединённые ячейки (${hitMerged}). Excel не строит сводную поверх объединений.`);

  return { problem: null, occupied };
}

/**
 * Первое свободное место под сводную на листе назначения.
 *
 * Сначала правее занятой области, потом под ней: оба места привычны человеку
 * и не режут данные. Пустоту проверяет сам Excel — чтения дешёвые, а гадать
 * по занятой области нельзя, на ней могли остаться одиночные заметки.
 */
async function findFreeCell(
  ctx: Excel.RequestContext,
  sheet: Excel.Worksheet,
  used: Excel.Range,
  expectation: PivotExpectation,
  sourceRect: A1Rect | null
): Promise<string | null> {
  const empty = Boolean((used as any).isNullObject);
  if (empty) return "A1";
  const candidates = [
    { row: used.rowIndex + 1, column: used.columnIndex + used.columnCount + 2 },
    { row: used.rowIndex + used.rowCount + 3, column: used.columnIndex + 1 }
  ];
  for (const candidate of candidates) {
    const cell = `${columnLetters(candidate.column)}${candidate.row}`;
    const check = await checkDestination(ctx, sheet, placeArea(cell, expectation), sourceRect);
    if (!check.problem && check.occupied === 0) return cell;
  }
  return null;
}

export async function prepareCreatePivotPlan(args: unknown): Promise<CreatePivotPlan> {
  preflightToolArgs("create_pivot_table", args);
  const raw = args as { sheet?: string; sourceAddress: string; destSheet?: string; destAddress?: string; newSheet?: string; name?: string; rows: string[]; values: unknown[]; grandTotals?: unknown; subtotals?: unknown } & RawPivotArgs;
  // Сверка Excel со своим расчётом панель делает по одному полю в столбцах;
  // остальные добавляются после неё (10.7).
  const a = Array.isArray(raw.columns) && raw.columns.length > 1 ? { ...raw, columns: raw.columns.slice(0, 1) } : raw;
  if (a.newSheet?.trim() && (a.destSheet?.trim() || a.destAddress?.trim())) {
    throw new ToolError("newSheet не сочетается с destSheet и destAddress: на новом листе сводная встаёт в A1.");
  }
  const source = checkAddress(a.sourceAddress);
  if (a.destAddress !== undefined) {
    const cell = parseA1Rect(a.destAddress);
    if (!cell || cell.kind !== "cells" || cell.rowStart !== cell.rowEnd || cell.columnStart !== cell.columnEnd) {
      throw new ToolError(`destAddress должен быть одной ячейкой, например H1; получено «${a.destAddress}».`);
    }
  }
  if (!Array.isArray(a.rows) || !a.rows.length) throw new ToolError("Нужно хотя бы одно поле в rows.");
  const valueFields = parseValueFields(a.values);
  const target = await captureTarget(a.sheet);

  const prepared = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    const range = await rangeOf(ctx, sheet, source);
    range.load(["address", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
    sheet.load(["id", "name"]);
    await ctx.sync();
    const cells = range.rowCount * range.columnCount;
    if (cells > MAX_IO_CELLS) {
      throw new ToolError(`Сводная строится по области до ${MAX_IO_CELLS} ячеек: панель должна посчитать её заранее. В ${range.address} ${cells}.`);
    }
    if (range.rowCount < 2) throw new ToolError("В источнике нужна строка заголовков и хотя бы одна строка данных.");
    range.load(["values", "formulas"]);
    await ctx.sync();
    const values = range.values as unknown[][];

    const headerProblems = pivotHeaderProblems(values[0]);
    if (headerProblems.length) {
      throw new ToolError(`Шапка ${range.address} не годится для сводной: ${headerProblems.join("; ")}. Операция не выполнялась.`);
    }
    const finish = parseFinish(raw, values[0]);
    const sourceHeaders = values[0].map((item) => String(item ?? ""));
    // Своё имя сводной: Excel требует уникального в книге.
    let pivotName = `Сводная_${Date.now().toString(36)}`;
    if (typeof raw.name === "string" && raw.name.trim()) {
      pivotName = raw.name.trim();
      if (pivotName.length > 100 || /[\[\]:*?/\\]/.test(pivotName)) throw new ToolError(`Имя сводной «${pivotName}» не годится: до 100 знаков, без [ ] : * ? / \\.`);
      const taken = ctx.workbook.pivotTables.getItemOrNullObject(pivotName);
      taken.load("isNullObject");
      await ctx.sync();
      if (!taken.isNullObject) throw new ToolError(`Сводная «${pivotName}» в книге уже есть: выберите другое имя.`);
    }
    // Источник ровно совпадает с таблицей Excel — строим по таблице: так
    // после «Обновить» сводная подхватит новые строки (10.7). С группировкой
    // дат — по адресу: вспомогательные столбцы стоят за пределами таблицы.
    let sourceTable: string | undefined;
    if (a.groupDates === undefined) {
      const tables = await readTablesStrict(ctx, sheet);
      const own = withoutSheet(range.address).replace(/\$/g, "").toUpperCase();
      sourceTable = tables.find((table) => withoutSheet(table.address).replace(/\$/g, "").toUpperCase() === own)?.name;
    }
    const missing = [...a.rows, ...valueFields.map((item) => item.field)].filter((name) => fieldIndex(values[0], name) < 0);
    if (missing.length) {
      throw new ToolError(
        `Нет полей ${missing.map((name) => `«${name}»`).join(", ")}. Заголовки источника: ${values[0].map((value) => `«${String(value)}»`).join(", ")}.`
      );
    }

    // Группировка дат подменяет поле дат вспомогательными столбцами — и в
    // данных для расчёта, и в rows/columns.
    const dates = a.groupDates !== undefined ? await planDateGroups(ctx, sheet, range, values, a) : null;
    const pivotValues = dates ? dates.values : values;
    const pivotRows = dates ? dates.rows : a.rows;
    const pivotArgs = dates ? { ...a, columns: dates.columns } : a;
    const pivotSourceRect = dates ? dates.rect : parseA1Rect(withoutSheet(range.address))!;
    const pivotSourceAddress = `${columnLetters(pivotSourceRect.columnStart)}${pivotSourceRect.rowStart}:${columnLetters(pivotSourceRect.columnEnd)}${pivotSourceRect.rowEnd}`;
    const resolved = resolvePivotOptions(pivotValues, pivotRows, valueFields, pivotArgs);
    const expectation = expectPivot(pivotValues, pivotRows, valueFields, resolved.options);
    if (!expectation.nodes.length) {
      throw new ToolError("После фильтров в сводной не остаётся ни одной строки: Excel построил бы пустую. Ослабьте фильтр.");
    }
    const extraPreview = [
      ...(sourceTable ? [`Источник — таблица «${sourceTable}»: после «Обновить» сводная подхватит новые строки`] : []),
      ...finishText(finish, valueFields.map((item) => item.field)),
      ...(dates ? [`Вспомогательные столбцы ${dates.plan.address}: ${dates.plan.names.map((item) => `«${item}»`).join(", ")} — формулы от «${dates.plan.field}»; отмена уберёт их вместе со сводной`] : []),
      ...optionsPreview(expectation, resolved, valueFields)
    ];
    const planOptions = {
      ...(sourceTable ? { sourceTable } : {}),
      sourceHeaders,
      finish,
      ...(dates ? { dateGroups: dates.plan } : {}),
      pivotSourceAddress,
      ...(resolved.columnField ? { columnField: resolved.columnField } : {}),
      filters: resolved.filters,
      ...(resolved.sort ? { sort: resolved.sort } : {})
    };

    // Новый лист (этап 7, 7.3.4): создаётся при исполнении, сводная — в A1.
    // Место проверять незачем — лист будет пуст; проверяется только имя.
    if (a.newSheet?.trim()) {
      const all = ctx.workbook.worksheets;
      all.load("items/name");
      await ctx.sync();
      const names = all.items.map((item) => item.name);
      let newName: string;
      try {
        newName = checkSheetName(a.newSheet, names);
      } catch (error: any) {
        const message = String(error?.message ?? error);
        throw new ToolError(/уже есть/.test(message) ? `${message} Свободно, например, «${freeSheetName(String(a.newSheet), names)}».` : message);
      }
      const undoNew = isCustomUndoAvailable();
      return {
        kind: "create_pivot_table" as const,
        id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
        target: { ...target, sheetName: sheet.name },
        name: pivotName,
        sourceAddress: withoutSheet(range.address),
        sourceRows: range.rowCount - 1,
        rowFields: [...pivotRows],
        valueFields,
        ...planOptions,
        destSheet: newName,
        destSheetId: "",
        newSheet: true as const,
        destCell: "A1",
        destArea: areaAt("A1", expectation.height, expectation.width).address,
        expectation,
        preview: [
          ...extraPreview,
          ...expectation.groups.slice(0, 8).map((group) => `${group.label}: ${group.totals.map((value) => Math.round(value * 100) / 100).join(" · ")}`),
          `Общий итог: ${expectation.grandTotals.map((value) => Math.round(value * 100) / 100).join(" · ")}`
        ],
        signature: JSON.stringify({ formulas: range.formulas, values }),
        sourceSameSheet: false,
        undoAvailable: undoNew,
        ...(undoNew ? {} : { undoNote: "Отмена недоступна: монитор изменений Excel не активен." }),
        createdAt: new Date().toISOString()
      };
    }

    // Лист назначения: указанный или тот же.
    const destSheet = a.destSheet?.trim()
      ? ctx.workbook.worksheets.getItemOrNullObject(a.destSheet.trim())
      : sheet;
    destSheet.load(["id", "name", "isNullObject"]);
    await ctx.sync();
    if ((destSheet as any).isNullObject) {
      throw new ToolError(`Листа «${a.destSheet}» нет. Создайте его через create_sheet или укажите существующий лист.`);
    }
    const sameSheet = destSheet.id === sheet.id;

    const used = officeCapabilities().usedRangeOrNull ? destSheet.getUsedRangeOrNullObject(true) : destSheet.getUsedRange(true);
    used.load(["isNullObject", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
    await ctx.sync();
    const emptySheet = Boolean((used as any).isNullObject);
    // Вспомогательные столбцы дат ещё не записаны, но место за ними занято.
    const usedEnd = emptySheet ? 0 : used.columnIndex + used.columnCount;
    const helperEnd = dates && sameSheet ? dates.rect.columnEnd : 0;
    const usedForPlace = !emptySheet && helperEnd > usedEnd
      ? { isNullObject: false, rowIndex: used.rowIndex, columnIndex: used.columnIndex, rowCount: used.rowCount, columnCount: helperEnd - used.columnIndex }
      : used;
    const destCell = (a.destAddress?.trim().toUpperCase())
      ?? (emptySheet ? "A1" : placementCell({ rowIndex: usedForPlace.rowIndex, columnIndex: usedForPlace.columnIndex, columnCount: usedForPlace.columnCount }, sameSheet ? range.rowIndex : 0));
    const area = placeArea(destCell, expectation);

    // Место под сводной обязано быть пустым: Excel не спрашивает, а данные
    // под ней пропадают или операция рвётся на середине.
    const sourceRect = pivotSourceRect;
    const check = await checkDestination(ctx, destSheet, area, sameSheet ? sourceRect : null);
    if (check.problem) throw new ToolError(`${check.problem} Выберите другое место.`);
    if (check.occupied) {
      // Проверка 20 сентября 2026 года: отказ говорил «укажите свободное
      // место», и агент на этом сдавался, хотя рядом было пусто. Свободное
      // место ищет панель — она и так знает размер будущей сводной.
      const free = await findFreeCell(ctx, destSheet, usedForPlace as Excel.Range, expectation, sameSheet ? sourceRect : null);
      throw new ToolError(
        `Сводная займёт ${destSheet.name}!${area.address}, а там ${check.occupied} непустых ячеек — они были бы затёрты. Операция не выполнялась. ` +
        (free
          ? `Свободно, например, ${destSheet.name}!${free} — повторите с destAddress: "${free}".`
          : "Свободного места такого размера на листе не нашлось: укажите другой лист в destSheet.")
      );
    }

    const preview = [
      ...extraPreview,
      ...expectation.groups.slice(0, 8).map((group) => `${group.label}: ${group.totals.map((value) => Math.round(value * 100) / 100).join(" · ")}`),
      `Общий итог: ${expectation.grandTotals.map((value) => Math.round(value * 100) / 100).join(" · ")}`
    ];
    const undo = isCustomUndoAvailable();
    return {
      kind: "create_pivot_table" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      target: { ...target, sheetName: sheet.name },
      name: pivotName,
      sourceAddress: withoutSheet(range.address),
      sourceRows: range.rowCount - 1,
      rowFields: [...pivotRows],
      valueFields,
      ...planOptions,
      destSheet: destSheet.name,
      destSheetId: destSheet.id,
      destCell,
      // Итоговое место — после фильтров; проверялось место побольше, до них.
      destArea: areaAt(destCell, expectation.height, expectation.width).address,
      expectation,
      preview,
      // Подпись по формулам и значениям: формула источника может ссылаться
      // на другой лист, и тогда её текст прежний, а расчёт панели — уже нет.
      signature: JSON.stringify({ formulas: range.formulas, values }),
      sourceSameSheet: sameSheet,
      undoAvailable: undo,
      ...(undo ? {} : { undoNote: "Отмена недоступна: монитор изменений Excel не активен." }),
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

/** Убирает сводную, которую не удалось достроить, и лист, созданный под неё.
 * true — только если повторное чтение книги подтвердило, что их больше нет. */
async function removeUnfinishedPivot(ctx: Excel.RequestContext, name: string, createdSheetId: string | null): Promise<boolean> {
  try {
    const pivot = ctx.workbook.pivotTables.getItemOrNullObject(name);
    pivot.load("isNullObject");
    await ctx.sync();
    if (!pivot.isNullObject) {
      pivot.delete();
      await ctx.sync();
    }
    if (createdSheetId) {
      const created = ctx.workbook.worksheets.getItemOrNullObject(createdSheetId);
      created.load("isNullObject");
      await ctx.sync();
      if (!created.isNullObject) {
        // На листе уже что-то есть — значит, не только наша сводная. Не трогаем.
        const used = created.getUsedRangeOrNullObject(true);
        used.load("isNullObject");
        await ctx.sync();
        if (!used.isNullObject) return false;
        created.delete();
        await ctx.sync();
      }
    }
    const pivotLeft = ctx.workbook.pivotTables.getItemOrNullObject(name);
    pivotLeft.load("isNullObject");
    const sheetLeft = createdSheetId ? ctx.workbook.worksheets.getItemOrNullObject(createdSheetId) : null;
    sheetLeft?.load("isNullObject");
    await ctx.sync();
    return pivotLeft.isNullObject && (!sheetLeft || sheetLeft.isNullObject);
  } catch {
    return false;
  }
}

export async function executeCreatePivotPlan(plan: CreatePivotPlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(plan.target.sheetId);
    const source = sheet.getRange(plan.sourceAddress);
    source.load(["formulas", "values"]);
    await ctx.sync();
    if (JSON.stringify({ formulas: source.formulas, values: source.values }) !== plan.signature) {
      throw new ToolExecutionError(
        `Данные источника ${plan.sourceAddress} изменились после предпросмотра — формулы или их значения. ` +
        "Расчёт групп и итогов устарел. Сводная не строилась — сделайте новый предпросмотр.",
        "failed_before_write"
      );
    }
    let destSheet: Excel.Worksheet;
    let createdSheetId: string | null = null;
    if (plan.newSheet) {
      // Имя могли занять за время подтверждения.
      const all = ctx.workbook.worksheets;
      all.load("items/name");
      await ctx.sync();
      if (all.items.some((item) => item.name.trim().toLowerCase() === plan.destSheet.toLowerCase())) {
        throw new ToolExecutionError(`Лист «${plan.destSheet}» появился после предпросмотра. Сводная не строилась — выберите другое имя.`, "failed_before_write");
      }
      try {
        destSheet = ctx.workbook.worksheets.add(plan.destSheet);
        destSheet.load("id");
        await ctx.sync();
        createdSheetId = destSheet.id;
      } catch (error: any) {
        throw new ToolExecutionError(`Excel отказал в создании листа «${plan.destSheet}»: ${error?.message ?? error}. Неизвестно, появился ли он — посмотрите на книгу.`, "unknown");
      }
    } else {
      destSheet = ctx.workbook.worksheets.getItem(plan.destSheetId);
      // То же правило места, что и при подготовке: за время подтверждения
      // могли появиться данные, таблица, защита или объединение.
      const sourceRect = plan.sourceSameSheet ? parseA1Rect(plan.pivotSourceAddress) : null;
      let check: DestinationCheck;
      try {
        check = await checkDestination(ctx, destSheet, placeArea(plan.destCell, plan.expectation), sourceRect);
      } catch (error: any) {
        throw new ToolExecutionError(`${error?.message ?? error} Сводная не строилась.`, "failed_before_write");
      }
      if (check.problem) {
        throw new ToolExecutionError(`${check.problem} Это появилось после предпросмотра. Сводная не строилась.`, "failed_before_write");
      }
      if (check.occupied) {
        throw new ToolExecutionError(`Место ${plan.destSheet}!${plan.destArea} перестало быть пустым после предпросмотра. Сводная не строилась.`, "failed_before_write");
      }
    }

    // Вспомогательные столбцы дат: место перепроверяется, формулы пишутся и
    // сверяются с расчётом панели до того, как по ним строится сводная.
    const dates = plan.dateGroups;
    if (dates) {
      const rect = parseA1Rect(dates.address)!;
      let check: DestinationCheck;
      try {
        check = await checkDestination(ctx, sheet, { rect, address: dates.address }, null);
      } catch (error: any) {
        throw new ToolExecutionError(`${error?.message ?? error} Сводная не строилась.`, "failed_before_write");
      }
      if (check.problem || check.occupied) {
        throw new ToolExecutionError(`Место под вспомогательные столбцы ${dates.address} перестало быть пустым после предпросмотра. Сводная не строилась.`, "failed_before_write");
      }
      const helper = sheet.getRange(dates.address);
      try {
        helper.formulas = dates.formulas.map((row) => [...row]) as any;
        await ctx.sync();
        helper.load("values");
        await ctx.sync();
      } catch (error: any) {
        const cleared = await clearDateHelpers(ctx, sheet, dates);
        throw new ToolExecutionError(
          `Excel отказал в записи вспомогательных столбцов ${dates.address}: ${error?.message ?? error}. ` +
          (cleared ? "Они убраны, сводная не строилась." : "Посмотрите на эти столбцы — убрать их не удалось."),
          cleared ? "failed_before_write" : "unknown"
        );
      }
      const got = helper.values as unknown[][];
      const wrong = dates.expected.flatMap((row, r) => row.flatMap((value, c) => (String(got[r]?.[c] ?? "") === String(value) ? [] : [{ r, c, value, got: got[r]?.[c] }])));
      if (wrong.length) {
        const cleared = await clearDateHelpers(ctx, sheet, dates);
        const first = wrong[0];
        throw new ToolExecutionError(
          `Формулы вспомогательных столбцов дали не то, что посчитала панель: строка ${first.r + 1} — «${String(first.got)}» вместо «${String(first.value)}»` +
          `${wrong.length > 1 ? ` (всего расхождений: ${wrong.length})` : ""}. ` +
          (cleared ? "Столбцы убраны, сводная не строилась." : `Убрать столбцы ${dates.address} не удалось — посмотрите на них.`),
          cleared ? "failed_before_write" : "unknown"
        );
      }
    }
    const pivotSource: Excel.Range | Excel.Table = plan.sourceTable
      ? sheet.tables.getItem(plan.sourceTable)
      : dates ? sheet.getRange(plan.pivotSourceAddress) : source;

    let pivot: Excel.PivotTable;
    try {
      pivot = destSheet.pivotTables.add(plan.name, pivotSource, destSheet.getRange(plan.destCell));
      await ctx.sync();
    } catch (error: any) {
      if (dates) await clearDateHelpers(ctx, sheet, dates);
      // Созданный под сводную лист без сводной не нужен: убрать его, если пуст.
      if (createdSheetId) {
        try {
          const orphan = ctx.workbook.worksheets.getItem(createdSheetId);
          const pivots = orphan.pivotTables;
          pivots.load("items/name");
          await ctx.sync();
          if (!pivots.items.length) { orphan.delete(); await ctx.sync(); }
        } catch { /* лист останется — об этом говорит сообщение ниже */ }
      }
      throw new ToolExecutionError(
        `Excel отказал в построении сводной: ${error?.message ?? error}. Неизвестно, успела ли она появиться — посмотрите на лист ${plan.destSheet}.`,
        "unknown"
      );
    }

    try {
      // Макет по умолчанию задаётся в настройках Excel, а размер и сверку
      // панель рассчитывает для табличного с итогами внизу групп: только
      // в нём у каждого уровня свой столбец, и вложенные итоги проверяемы
      // (план стабилизации, S3.1). Выставляется до полей, чтобы сводная
      // ни на каком шаге не была шире рассчитанного места. Оба свойства —
      // ExcelApi 1.8, как и сами сводные.
      pivot.layout.layoutType = "Tabular" as any;
      pivot.layout.subtotalLocation = "AtBottom" as any;
      for (const field of plan.rowFields) pivot.rowHierarchies.add(pivot.hierarchies.getItem(field));
      for (const item of plan.valueFields) {
        const data = pivot.dataHierarchies.add(pivot.hierarchies.getItem(item.field));
        data.summarizeBy = OFFICE_AGGREGATION[item.aggregation] as any;
      }
      await ctx.sync();
      // Порядок по значению — до поля в столбцах: при поле в столбцах Excel
      // сортирует не по общему итогу, а по первому столбцу («Книга602»,
      // 08.10.2026: область сортировки — «Год постав. = 2000», строки вразнобой).
      let sortedEarly = false;
      if (plan.sort && plan.columnField) {
        const data = pivot.dataHierarchies;
        data.load("items/name");
        await ctx.sync();
        pivot.rowHierarchies.getItem(plan.sort.field).fields.getItem(plan.sort.field)
          .sortByValues((plan.sort.order === "desc" ? "Descending" : "Ascending") as any, data.items[plan.sort.by]);
        await ctx.sync();
        sortedEarly = true;
      }
      if (plan.columnField) {
        pivot.columnHierarchies.add(pivot.hierarchies.getItem(plan.columnField));
        await ctx.sync();
      }
      if (plan.filters.length || (plan.sort && !sortedEarly)) {
        // Фильтр «первые N» и порядок ссылаются на поле значений по его имени
        // в сводной («Сумма по полю Сумма») — оно на языке интерфейса, поэтому
        // читается из Excel, а не собирается панелью.
        const data = pivot.dataHierarchies;
        data.load("items/name");
        await ctx.sync();
        const fieldOf = (name: string, axis: "rows" | "columns") =>
          (axis === "columns" ? pivot.columnHierarchies : pivot.rowHierarchies).getItem(name).fields.getItem(name);
        for (const filter of plan.filters) {
          const field = fieldOf(filter.field, filter.axis);
          if ("include" in filter) {
            field.applyFilter({ manualFilter: { selectedItems: [...filter.include] } } as any);
          } else {
            const top = "top" in filter;
            field.applyFilter({
              valueFilter: { condition: top ? "TopN" : "BottomN", threshold: top ? filter.top : filter.bottom, value: data.items[filter.by].name, selectionType: "Items" }
            } as any);
          }
        }
        if (plan.sort && !sortedEarly) {
          fieldOf(plan.sort.field, "rows").sortByValues((plan.sort.order === "desc" ? "Descending" : "Ascending") as any, data.items[plan.sort.by]);
        }
        await ctx.sync();
      }
    } catch (error: any) {
      const reason = error?.message ?? error;
      const removed = await removeUnfinishedPivot(ctx, plan.name, createdSheetId);
      const helpersCleared = !dates || (removed && await clearDateHelpers(ctx, sheet, dates));
      if (removed && helpersCleared) {
        throw new ToolExecutionError(
          `Excel отказал в добавлении полей сводной: ${reason}. Недостроенная сводная удалена` +
            `${createdSheetId ? ` вместе с созданным под неё листом «${plan.destSheet}»` : ""}` +
            `${dates ? `, вспомогательные столбцы ${dates.address} убраны` : ""}; книга в прежнем виде.`,
          "failed_before_write"
        );
      }
      throw new ToolExecutionError(
        `Excel отказал в добавлении полей сводной: ${reason}. Недостроенную сводную ${plan.name} убрать не удалось — посмотрите на лист ${plan.destSheet}.`,
        "unknown"
      );
    }

    let undoRecorded = false;
    if (plan.undoAvailable) {
      const name = plan.name;
      undoRecorded = push(action(`сводная ${name} на листе ${plan.destSheet}`, async () => {
        const revision = getStructuralRevision();
        await Excel.run(async (undoCtx) => {
          const existing = undoCtx.workbook.pivotTables.getItemOrNullObject(name);
          existing.load("isNullObject");
          await undoCtx.sync();
          if (existing.isNullObject) throw new Error("Сводной уже нет: её удалили после операции агента. Отменять нечего.");
          if (getStructuralRevision() !== revision) throw new Error("Структура книги изменилась во время отмены. Отмена остановлена.");
          // Вспомогательные столбцы уходят вместе со сводной — если их не
          // меняли. Изменённые — уже чужая работа: отмена останавливается целиком.
          const helper = dates ? undoCtx.workbook.worksheets.getItem(plan.target.sheetId).getRange(dates.address) : null;
          if (helper) {
            helper.load("formulas");
            await undoCtx.sync();
            if (JSON.stringify(helper.formulas) !== JSON.stringify(dates!.formulas)) {
              throw new Error(`Вспомогательные столбцы ${dates!.address} изменили после операции агента. Отмена остановлена: сводная и столбцы на месте.`);
            }
          }
          existing.delete();
          await undoCtx.sync();
          if (helper) {
            helper.clear("Contents" as any);
            await undoCtx.sync();
          }
          // Лист, созданный под сводную, уходит вместе с ней — если на нём
          // больше ничего нет. Иначе остаётся: там уже чужая работа.
          if (createdSheetId) {
            const created = undoCtx.workbook.worksheets.getItemOrNullObject(createdSheetId);
            created.load("isNullObject");
            await undoCtx.sync();
            if (!created.isNullObject) {
              const used = created.getUsedRangeOrNullObject(true);
              used.load("isNullObject");
              await undoCtx.sync();
              if (used.isNullObject) { created.delete(); await undoCtx.sync(); }
            }
          }
        });
      }));
    }

    const layout = pivot.layout.getRange();
    layout.load(["address", "values"]);
    await ctx.sync();
    const actualArea = withoutSheet(String(layout.address));
    const problems = pivotMismatches(plan.expectation, layout.values as unknown[][]);
    if (actualArea !== plan.destArea) problems.unshift(`заняла ${actualArea} вместо ${plan.destArea}`);
    if (problems.length) {
      throw new ToolExecutionError(
        `Сводная ${plan.name} построена, но расходится с расчётом панели: ${problems.join("; ")}. ` +
        (undoRecorded ? "Её можно убрать кнопкой «Отменить»." : "Проверьте её на листе."),
        "applied"
      );
    }

    // Донастройка — после сверки чисел: подписи, формат, итоги и второе поле
    // в столбцах меняют вид сводной, а не её расчёт.
    let finished: Awaited<ReturnType<typeof applyPivotFinish>> | null = null;
    try {
      finished = await applyPivotFinish(ctx, pivot, plan.finish, plan.sourceHeaders);
    } catch (error: any) {
      throw new ToolExecutionError(
        `Сводная ${plan.name} построена и сверена, но донастроить её не удалось: ${error?.message ?? error}. ` +
        (undoRecorded ? "Её можно убрать кнопкой «Отменить»." : "Проверьте её на листе."),
        "applied"
      );
    }
    if (finished.problems.length) {
      throw new ToolExecutionError(
        `Сводная ${plan.name} построена и сверена, но донастройка встала не вся: ${finished.problems.join("; ")}. ` +
        (undoRecorded ? "Сводную можно убрать кнопкой «Отменить»." : "Проверьте её на листе."),
        "applied"
      );
    }
    const finalArea = pivot.layout.getRange();
    finalArea.load("address");
    await ctx.sync();

    return {
      ok: true,
      executionState: "verified",
      pivot: plan.name,
      ...(plan.sourceTable ? { sourceTable: plan.sourceTable, refreshNote: "Источник — таблица Excel: после «Обновить» (refresh_pivot) сводная подхватит новые строки." } : {}),
      finish: finished.applied,
      finalAddress: withoutSheet(String(finalArea.address)),
      sheet: plan.destSheet,
      address: actualArea,
      source: plan.sourceAddress,
      rows: plan.rowFields,
      ...(plan.columnField ? { columns: [plan.columnField], columnItems: plan.expectation.columns.map((item) => item.label).sort((x, y) => x.localeCompare(y, "ru")) } : {}),
      values: plan.valueFields.map((item) => `${item.field} — ${AGGREGATION_TEXT[item.aggregation]}`),
      ...(plan.filters.length ? {
        filters: plan.filters.map((filter) => "include" in filter
          ? { field: filter.field, include: filter.include }
          : { field: filter.field, ["top" in filter ? "top" : "bottom"]: "top" in filter ? filter.top : filter.bottom, by: plan.valueFields[filter.by].field }),
        filteredOut: plan.expectation.filteredOut
      } : {}),
      ...(plan.sort ? { sort: { field: plan.sort.field, by: plan.valueFields[plan.sort.by].field, order: plan.sort.order } } : {}),
      ...(dates ? { dateGroups: { field: dates.field, columns: dates.names, address: dates.address, note: "Вспомогательные столбцы с формулами — источник сводной; удалять их нельзя, пока нужна сводная." } } : {}),
      grandTotals: plan.expectation.grandTotals,
      groups: plan.expectation.groups.length,
      ...(plan.expectation.warnings.length ? { warnings: plan.expectation.warnings } : {}),
      note: plan.rowFields.length > 1
        ? "Итоги всех групп на всех уровнях и общий итог сверены с расчётом панели по исходным данным."
        : "Итоги каждой группы и общий итог сверены с расчётом панели по исходным данным.",
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: plan.undoNote ?? "Автоматическая отмена этой операции недоступна." })
    };
  });
}
