import {
  action,
  captureContent,
  captureExactFormat,
  exactFormatUndo,
  guardedContentUndo,
  invalidateAfterStructuralChange,
  isCustomUndoAvailable,
  push
} from "./undo";
import { supported, TOOL_BY_NAME, validateToolArgs, writableAtCurrentStage, type ToolName } from "./toolSchemas";
import { assertRangeReference, cellCount, EXCEL_MAX_COLUMNS, EXCEL_MAX_ROWS, intersects, parseA1Rect } from "./a1";
import {
  assertWorkbookTarget,
  captureTarget,
  currentWorkbookIdentity,
  getActiveContext,
  getSheetOverview,
  listSheets,
  officeCapabilities,
  type WorkbookTarget
} from "./workbookContext";
import { getRevisionCoverage, getWorkbookRevision } from "./workbookRevision";
import { recallSnapshot, recordSnapshot, setSnapshotPinned } from "./snapshotStore";
import { measureWorkbookExport } from "./workbookExport";

export class ToolError extends Error {}

export type ExecutionState = "not_started" | "applied" | "verified" | "failed_before_write" | "unknown";

/** A failed write must never be presented as proof that the workbook was unchanged. */
export class ToolExecutionError extends ToolError {
  constructor(message: string, readonly executionState: ExecutionState) {
    super(message);
    this.name = "ToolExecutionError";
  }
}

/** Office.js отдаёт null, когда свойство неоднородно по диапазону. «Значение не
 * задано» и «в области разное» — разные утверждения, и подавать их одинаково
 * нельзя: первое успокаивает, второе требует посмотреть внимательнее. Заменяем
 * null явным маркером и собираем список неоднородных свойств. */
export const MIXED_FORMAT = "разное в области";

export function markMixed(
  // Office.js отдаёт из toJSON() именованные типы без индексной сигнатуры,
  // поэтому принимаем любой объект и разбираем его по парам ключ-значение.
  source: object,
  prefix: string,
  mixed: string[]
): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(source)) {
    if (value === null) {
      mixed.push(`${prefix}${key}`);
      out[key] = MIXED_FORMAT;
    } else {
      out[key] = value;
    }
  }
  return out;
}

/** Насколько расширяется зона опроса объединений вокруг запрошенной области.
 * `getMergedAreasOrNullObject` отдаёт объединения, попадающие внутрь диапазона,
 * причём только при попадании в запрос его левого верхнего угла. Поэтому
 * ячейка внутри объединения возвращала пустой список и выглядела обычной.
 * Опрашиваем окрестность, чтобы угол нашёлся. Объединение, чей угол дальше
 * этого запаса, найдено не будет — такие в книгах редки. */
export const MERGED_PROBE_MARGIN = 20;

/** Адрес приходит как `Лист!K1:L1`; для сравнения прямоугольников имя не нужно. */
function withoutSheet(address: string): string {
  const cut = address.lastIndexOf("!");
  return cut >= 0 ? address.slice(cut + 1) : address;
}

/** Собирает адреса объединений из обоих источников Office.js. Поле address
 * перечисляет области через запятую, коллекция areas — по отдельности; какой
 * из них заполнен, зависит от сборки, поэтому берём всё, что пришло. */
export function mergedAddressCandidates(merged: {
  address?: string | null;
  areas?: { items?: { address?: string }[] } | null;
}): string[] {
  const fromAddress = String(merged.address ?? "").split(",");
  const fromAreas = (merged.areas?.items ?? []).map((area) => String(area?.address ?? ""));
  return [...fromAreas, ...fromAddress].map((item) => item.trim()).filter(Boolean);
}

export interface MergedAreasReport {
  /** Объединения с достоверными границами, задевающие запрошенную область. */
  areas: string[];
  /** Якоря, чья протяжённость неизвестна, но которые могли бы накрыть цель. */
  unresolvedAnchors: string[];
}

/** Разбирает ответ Office.js об объединениях.
 *
 * Замер на Office 16.0.14334 показал две особенности, из-за которых прежний
 * разбор уверенно сообщал «объединений нет» там, где они есть. Объединение
 * возвращается, только если в запрос попал его левый верхний угол: чтение
 * `O1:Q1` целиком внутри `N1:P1` не находило ничего. И адрес всегда усечён до
 * этого угла: `N1:P1` приходит как `N1`.
 *
 * Отсюда правило: область размером в одну ячейку объединением быть не может —
 * Excel не объединяет одну ячейку, — значит это усечённый якорь неизвестной
 * протяжённости. Объединение растёт вправо и вниз, поэтому накрыть цель может
 * только якорь, стоящий не правее и не ниже её левого верхнего угла. Такие
 * якоря возвращаются отдельно как предупреждение, а не как факт. */
export function mergedAreasTouching(addresses: string[], target: string): MergedAreasReport {
  const rect = parseA1Rect(withoutSheet(target));
  if (!rect) return { areas: [], unresolvedAnchors: [] };
  const seen = new Set<string>();
  const areas: string[] = [];
  const unresolvedAnchors: string[] = [];
  for (const raw of addresses) {
    const address = raw.trim();
    if (!address || seen.has(address)) continue;
    const area = parseA1Rect(withoutSheet(address));
    if (!area) continue;
    seen.add(address);
    if (cellCount(area) > 1) {
      if (intersects(area, rect)) areas.push(address);
      continue;
    }
    if (area.rowStart <= rect.rowStart && area.columnStart <= rect.columnStart) unresolvedAnchors.push(address);
  }
  return { areas, unresolvedAnchors };
}

export interface MergedProbeResult extends MergedAreasReport {
  diagnostics: {
    address: string | null;
    areaCount: number;
    areaItems: number;
    probed: string | null;
    /** Сборка Office.js не поддержала опрос: объединения не проверялись вовсе. */
    unavailable?: true;
  };
}

/** Опрашивает объединения вокруг диапазона и синхронизирует контекст.
 * Требует уже загруженных rowIndex, columnIndex, rowCount, columnCount
 * и address у range. Используется и подробностями, и подготовкой записи:
 * предупреждение о возможном объединении нужно прежде всего перед правкой. */
async function probeMergedAreas(
  ctx: Excel.RequestContext,
  sheet: Excel.Worksheet,
  range: Excel.Range
): Promise<MergedProbeResult> {
  // Опрос вспомогательный: он уточняет предупреждение, но не решает, можно ли
  // писать. Если сборка Office.js этих вызовов не знает, разумнее отдать
  // «не проверяли», чем уронить подготовку плана записи целиком.
  if (typeof sheet.getRangeByIndexes !== "function" || typeof range.getMergedAreasOrNullObject !== "function") {
    return {
      areas: [],
      unresolvedAnchors: [],
      diagnostics: { address: null, areaCount: 0, areaItems: 0, probed: null, unavailable: true }
    };
  }
  const startRow = Math.max(0, range.rowIndex - MERGED_PROBE_MARGIN);
  const startColumn = Math.max(0, range.columnIndex - MERGED_PROBE_MARGIN);
  const endRow = Math.min(EXCEL_MAX_ROWS - 1, range.rowIndex + range.rowCount - 1 + MERGED_PROBE_MARGIN);
  const endColumn = Math.min(EXCEL_MAX_COLUMNS - 1, range.columnIndex + range.columnCount - 1 + MERGED_PROBE_MARGIN);
  const probe = sheet.getRangeByIndexes(startRow, startColumn, endRow - startRow + 1, endColumn - startColumn + 1);
  probe.load("address");
  const merged = probe.getMergedAreasOrNullObject();
  merged.load(["isNullObject", "address", "areaCount"]);
  merged.areas.load("items/address");
  await ctx.sync();
  const report = merged.isNullObject
    ? { areas: [], unresolvedAnchors: [] }
    : mergedAreasTouching(mergedAddressCandidates(merged), range.address);
  return {
    ...report,
    diagnostics: {
      address: merged.isNullObject ? null : String(merged.address ?? ""),
      areaCount: merged.isNullObject ? 0 : merged.areaCount,
      areaItems: merged.isNullObject ? 0 : (merged.areas?.items?.length ?? 0),
      probed: probe.address ?? null
    }
  };
}

/** Единая формулировка про неизвестные границы объединений. */
export function mergeAnchorNote(anchors: readonly string[], address: string, beforeWrite: boolean): string {
  const tail = beforeWrite
    ? "Запись в объединённую ячейку ведёт себя не так, как в обычную. Отсутствие объединения здесь не доказано: проверьте цель в Excel перед подтверждением."
    : "Отсутствие объединения здесь не доказано: перед записью проверьте цель в Excel.";
  return `Рядом найдены углы объединений ${anchors.join(", ")}. Excel на этой сборке отдаёт только угол и не сообщает границы, поэтому ${address} может оказаться внутри одного из них. ${tail}`;
}

const MAX_IO_CELLS = 20_000;
const MAX_ROWS_PER_STRUCTURAL_OP = 1000;
const MAX_EXACT_FORMAT_UNDO_CELLS = 500;
const MAX_DETAILS_CELLS = 500;
const SEARCH_CHUNK_CELLS = 5_000;
const MAX_SEARCH_SHEETS = 100;

function checkAddress(address: unknown): string {
  try { return assertRangeReference(address); }
  catch (error: any) { throw new ToolError(error?.message ?? String(error)); }
}

function sheetOf(ctx: Excel.RequestContext, name?: string) {
  const n = (name ?? "").trim();
  return n ? ctx.workbook.worksheets.getItem(n) : ctx.workbook.worksheets.getActiveWorksheet();
}

async function rangeOf(ctx: Excel.RequestContext, sheet: Excel.Worksheet, reference: string): Promise<Excel.Range> {
  if (parseA1Rect(reference)) return sheet.getRange(reference);
  const localName = sheet.names.getItemOrNullObject(reference);
  const workbookName = ctx.workbook.names.getItemOrNullObject(reference);
  localName.load("isNullObject");
  workbookName.load("isNullObject");
  await ctx.sync();
  const item = !localName.isNullObject ? localName : !workbookName.isNullObject ? workbookName : null;
  if (!item) throw new ToolError(`Именованный диапазон "${reference}" не найден на листе или в книге.`);
  const range = item.getRangeOrNullObject();
  range.load("isNullObject");
  await ctx.sync();
  if (range.isNullObject) throw new ToolError(`Имя "${reference}" не ссылается на диапазон ячеек.`);
  return range;
}

function rowsAddress(startRow: number, count: number) {
  if (!Number.isInteger(startRow) || startRow < 1) throw new ToolError("startRow должен быть целым числом от 1.");
  if (!Number.isInteger(count) || count < 1 || count > MAX_ROWS_PER_STRUCTURAL_OP) {
    throw new ToolError(`count должен быть целым числом от 1 до ${MAX_ROWS_PER_STRUCTURAL_OP}.`);
  }
  return `${startRow}:${startRow + count - 1}`;
}



export async function getActiveSheetName(): Promise<string> {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    await ctx.sync();
    return sheet.name;
  });
}

/**
 * Разрешает неявный sheet относительно листа, зафиксированного в начале
 * пользовательской задачи. Это важно: активная вкладка может измениться, пока
 * модель думает или пользователь подтверждает действие.
 */
export async function resolveToolArgs(
  name: string,
  args: unknown,
  taskSheet?: string
): Promise<unknown> {
  if (!args || typeof args !== "object" || Array.isArray(args)) return args;
  const current = { ...(args as Record<string, unknown>) };
  const acceptsSheet = new Set([
    "get_sheet_overview",
    "get_range_values",
    "get_range_details",
    "set_range_values",
    "insert_rows",
    "delete_rows",
    "sort_range",
    "apply_filter",
    "create_pivot_table",
    "create_chart",
    "format_range"
  ]).has(name);
  if (!acceptsSheet) return current;
  if (typeof current.sheet === "string" && current.sheet.trim()) {
    current.sheet = current.sheet.trim();
  } else {
    current.sheet = taskSheet || (await getActiveSheetName());
  }

  if (
    name === "create_pivot_table" &&
    !(typeof current.destSheet === "string" && current.destSheet.trim())
  ) {
    current.destSheet = current.sheet;
  }
  return current;
}

function hexToColor(hex: string) {
  const v = hex.trim();
  if (!/^#?[0-9A-Fa-f]{6}$/.test(v)) {
    throw new ToolError(`Цвет "${hex}" не в формате HEX, ожидается #RRGGBB.`);
  }
  return v.startsWith("#") ? v : `#${v}`;
}

/** Checks that do not require Excel, before asking the user to approve a write. */
export function preflightToolArgs(name: string, args: unknown): void {
  const validation = validateToolArgs(name, args);
  if (!validation.ok) throw new ToolError(validation.error);
  const a = args as Record<string, any>;
  if (typeof a.address === "string") checkAddress(a.address);
  if (typeof a.sourceAddress === "string") checkAddress(a.sourceAddress);
  if (typeof a.destAddress === "string") checkAddress(a.destAddress);
  if (name === "insert_rows" || name === "delete_rows") rowsAddress(a.startRow, a.count);
  if (name === "format_range" && typeof a.fillColor === "string") hexToColor(a.fillColor);
  if (name === "set_range_values") {
    if (!Array.isArray(a.values) || !a.values.length || !Array.isArray(a.values[0]) || !a.values[0].length) {
      throw new ToolError("values должен быть непустым двумерным массивом.");
    }
    const width = a.values[0].length;
    if (a.values.some((row: unknown) => !Array.isArray(row) || row.length !== width)) {
      throw new ToolError("Строки в values разной длины.");
    }
    if (a.values.length * width > MAX_IO_CELLS) throw new ToolError(`За одну запись разрешено не более ${MAX_IO_CELLS} ячеек.`);
  }
}

async function get_active_context() {
  return getActiveContext();
}

async function list_sheets() {
  return listSheets();
}

async function get_sheet_overview(a: { sheet?: string }) {
  return getSheetOverview(a.sheet);
}

async function get_range_values(a: {
  sheet?: string;
  address: string;
  properties?: Array<"values" | "formulas" | "text" | "valueTypes" | "numberFormat">;
}) {
  const address = checkAddress(a.address);
  const revisionAtStart = getWorkbookRevision();
  const result = await Excel.run(async (ctx) => {
    const sheet = sheetOf(ctx, a.sheet);
    const range = await rangeOf(ctx, sheet, address);

    // Сначала загружаем только размеры: нельзя сначала выкачать миллион ячеек,
    // а потом обнаружить, что диапазон слишком большой.
    range.load(["address", "rowCount", "columnCount"]);
    sheet.load(["id", "name"]);
    await ctx.sync();

    const cells = range.rowCount * range.columnCount;
    if (cells > MAX_IO_CELLS) {
      throw new ToolError(
        `Диапазон ${address} содержит ${cells} ячеек — слишком много для одного запроса. Читай частями до ${MAX_IO_CELLS} ячеек.`
      );
    }

    const requested = a.properties?.length ? a.properties : ["values", "formulas"];
    range.load(requested);
    await ctx.sync();

    return {
      sheetId: sheet.id,
      sheet: sheet.name,
      address: range.address,
      rowCount: range.rowCount,
      columnCount: range.columnCount,
      properties: requested,
      ...(requested.includes("values") ? { values: range.values } : {}),
      ...(requested.includes("formulas") ? { formulas: range.formulas } : {}),
      ...(requested.includes("text") ? { text: range.text } : {}),
      ...(requested.includes("valueTypes") ? { valueTypes: range.valueTypes } : {}),
      ...(requested.includes("numberFormat") ? { numberFormat: range.numberFormat } : {})
    };
  });
  const snapshot = recordSnapshot({
    kind: "content",
    source: "get_range_values",
    workbook: currentWorkbookIdentity(),
    sheetId: result.sheetId,
    sheetName: result.sheet,
    address: result.address,
    revision: revisionAtStart,
    coverage: getRevisionCoverage(),
    payload: result
  });
  const state = snapshot ? recallSnapshot(snapshot.id, currentWorkbookIdentity(), getWorkbookRevision(), getRevisionCoverage()).state : "not_stored";
  return { ...result, snapshot: snapshot ? { id: snapshot.id, state, capturedAt: snapshot.capturedAt, revision: snapshot.revision } : { state } };
}

export interface SearchCursor {
  sheetIndex: number;
  cellOffset: number;
}

function searchFingerprint(a: Record<string, unknown>): string {
  const source = JSON.stringify({
    query: a.query,
    sheets: a.sheets ?? null,
    searchIn: a.searchIn ?? "both",
    matchCase: a.matchCase === true,
    wholeCell: a.wholeCell === true
  });
  let hash = 0x811c9dc5;
  for (let i = 0; i < source.length; i++) {
    hash ^= source.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(16).padStart(8, "0");
}

export function encodeSearchCursor(fingerprint: string, cursor: SearchCursor): string {
  return `v1:${fingerprint}:${cursor.sheetIndex}:${cursor.cellOffset}`;
}

export function decodeSearchCursor(value: string | undefined, fingerprint: string): SearchCursor {
  if (!value) return { sheetIndex: 0, cellOffset: 0 };
  const match = /^v1:([0-9a-f]{8}):(\d+):(\d+)$/.exec(value);
  if (!match || match[1] !== fingerprint) {
    throw new ToolError("continuation недействителен или относится к другому поиску.");
  }
  const sheetIndex = Number(match[2]);
  const cellOffset = Number(match[3]);
  if (!Number.isSafeInteger(sheetIndex) || !Number.isSafeInteger(cellOffset)) {
    throw new ToolError("continuation содержит недопустимую позицию.");
  }
  return { sheetIndex, cellOffset };
}

export function searchTextMatches(value: unknown, query: string, matchCase = false, wholeCell = false): boolean {
  if (value === null || value === undefined) return false;
  const candidate = String(value);
  const left = matchCase ? candidate : candidate.toLocaleLowerCase();
  const right = matchCase ? query : query.toLocaleLowerCase();
  return wholeCell ? left === right : left.includes(right);
}

function columnLabel(column: number): string {
  let value = column;
  let label = "";
  while (value > 0) {
    value -= 1;
    label = String.fromCharCode(65 + value % 26) + label;
    value = Math.floor(value / 26);
  }
  return label;
}

function cellAddress(row: number, column: number): string {
  return `${columnLabel(column)}${row}`;
}

function throwIfSearchStopped(signal?: AbortSignal, deadlineAt?: number) {
  if (signal?.aborted) throw new DOMException("Остановлено пользователем", "AbortError");
  if (deadlineAt !== undefined && Date.now() >= deadlineAt) {
    throw new ToolError("Поиск остановлен по лимиту времени. Продолжите меньшими областями или отдельным запросом.");
  }
}

async function search_workbook(a: {
  query: string;
  sheets?: string[];
  searchIn?: "values" | "formulas" | "both";
  matchCase?: boolean;
  wholeCell?: boolean;
  limit?: number;
  continuation?: string;
}, options?: { signal?: AbortSignal; deadlineAt?: number }) {
  const query = a.query.trim();
  if (!query) throw new ToolError("query не может быть пустым.");
  const limit = a.limit ?? 50;
  const fingerprint = searchFingerprint(a as Record<string, unknown>);
  const start = decodeSearchCursor(a.continuation, fingerprint);
  const requestedNames = a.sheets?.map((name) => name.trim());

  const revisionAtStart = getWorkbookRevision();
  const result = await Excel.run(async (ctx) => {
    const collection = ctx.workbook.worksheets;
    collection.load("items/id,name,position");
    await ctx.sync();
    const all = [...collection.items].sort((left, right) => left.position - right.position);
    const sheets = requestedNames
      ? requestedNames.map((name) => {
          const found = all.find((sheet) => sheet.name === name);
          if (!found) throw new ToolError(`Лист "${name}" не найден.`);
          return found;
        })
      : all;
    if (sheets.length > MAX_SEARCH_SHEETS) {
      throw new ToolError(`Поиск за один проход ограничен ${MAX_SEARCH_SHEETS} листами. Укажите нужные листы явно.`);
    }
    if (start.sheetIndex > sheets.length) throw new ToolError("continuation указывает за пределы списка листов.");

    const capabilities = officeCapabilities();
    const usedRanges = sheets.map((sheet) => capabilities.usedRangeOrNull
      ? sheet.getUsedRangeOrNullObject(true)
      : sheet.getUsedRange(true));
    for (const used of usedRanges) used.load(["address", "rowIndex", "columnIndex", "rowCount", "columnCount", "isNullObject"]);
    await ctx.sync();

    const matches: Array<{
      sheetId: string;
      sheet: string;
      address: string;
      matchedIn: Array<"value" | "formula">;
      value: unknown;
      formula?: string;
    }> = [];
    const searchedAreas: Array<{ sheetId: string; sheet: string; address: string }> = [];
    let continuation: string | undefined;

    for (let sheetIndex = start.sheetIndex; sheetIndex < sheets.length && !continuation; sheetIndex++) {
      const sheet = sheets[sheetIndex];
      const used = usedRanges[sheetIndex];
      if (used.isNullObject) continue;
      const totalCells = used.rowCount * used.columnCount;
      let offset = sheetIndex === start.sheetIndex ? start.cellOffset : 0;
      if (offset > totalCells) throw new ToolError("continuation указывает за пределы используемой области.");

      while (offset < totalCells && !continuation) {
        throwIfSearchStopped(options?.signal, options?.deadlineAt);
        const rowOffset = Math.floor(offset / used.columnCount);
        const columnOffset = offset % used.columnCount;
        const availableInRow = used.columnCount - columnOffset;
        const rowCount = columnOffset === 0
          ? Math.max(1, Math.min(used.rowCount - rowOffset, Math.floor(SEARCH_CHUNK_CELLS / used.columnCount) || 1))
          : 1;
        const columnCount = columnOffset === 0 ? used.columnCount : Math.min(availableInRow, SEARCH_CHUNK_CELLS);
        const range = sheet.getRangeByIndexes(
          used.rowIndex + rowOffset,
          used.columnIndex + columnOffset,
          rowCount,
          columnCount
        );
        range.load(["values", "formulas"]);
        await ctx.sync();
        throwIfSearchStopped(options?.signal, options?.deadlineAt);

        let processed = 0;
        outer: for (let row = 0; row < rowCount; row++) {
          for (let column = 0; column < columnCount; column++) {
            const value = (range.values as unknown[][])[row][column];
            const formula = (range.formulas as unknown[][])[row][column];
            const absoluteRow = used.rowIndex + rowOffset + row + 1;
            const absoluteColumn = used.columnIndex + columnOffset + column + 1;
            const address = cellAddress(absoluteRow, absoluteColumn);
            const formulaText = typeof formula === "string" && formula.startsWith("=") ? formula : null;
            const matchedIn: Array<"value" | "formula"> = [];
            if (a.searchIn !== "formulas" && searchTextMatches(value, query, a.matchCase, a.wholeCell)) matchedIn.push("value");
            if (a.searchIn !== "values" && formulaText !== null && searchTextMatches(formulaText, query, a.matchCase, a.wholeCell)) matchedIn.push("formula");
            if (matchedIn.length) {
              matches.push({
                sheetId: sheet.id,
                sheet: sheet.name,
                address,
                matchedIn,
                value,
                ...(formulaText === null ? {} : { formula: formulaText })
              });
              if (matches.length >= limit) {
                processed = row * columnCount + column + 1;
                const nextOffset = offset + processed;
                if (nextOffset < totalCells) {
                  continuation = encodeSearchCursor(fingerprint, { sheetIndex, cellOffset: nextOffset });
                } else {
                  const nextSheet = usedRanges.findIndex((candidate, index) =>
                    index > sheetIndex && !candidate.isNullObject && candidate.rowCount * candidate.columnCount > 0
                  );
                  if (nextSheet >= 0) continuation = encodeSearchCursor(fingerprint, { sheetIndex: nextSheet, cellOffset: 0 });
                }
                break outer;
              }
            }
            processed = row * columnCount + column + 1;
          }
        }
        if (processed > 0) {
          const firstRow = used.rowIndex + rowOffset + 1;
          const firstColumn = used.columnIndex + columnOffset + 1;
          const completedRows = Math.floor(processed / columnCount);
          const partialColumns = processed % columnCount;
          if (completedRows > 0) {
            const first = cellAddress(firstRow, firstColumn);
            const last = cellAddress(firstRow + completedRows - 1, firstColumn + columnCount - 1);
            searchedAreas.push({ sheetId: sheet.id, sheet: sheet.name, address: first === last ? first : `${first}:${last}` });
          }
          if (partialColumns > 0) {
            const partialRow = firstRow + completedRows;
            const first = cellAddress(partialRow, firstColumn);
            const last = cellAddress(partialRow, firstColumn + partialColumns - 1);
            searchedAreas.push({ sheetId: sheet.id, sheet: sheet.name, address: first === last ? first : `${first}:${last}` });
          }
        }
        offset += processed;
      }
    }

    return {
      query,
      searchIn: a.searchIn ?? "both",
      matchCase: a.matchCase === true,
      wholeCell: a.wholeCell === true,
      matches,
      searchedAreas,
      incomplete: Boolean(continuation),
      ...(continuation ? { continuation } : {}),
      readAt: new Date().toISOString()
    };
  });
  const snapshot = recordSnapshot({
    kind: "search",
    source: "search_workbook",
    workbook: currentWorkbookIdentity(),
    revision: revisionAtStart,
    coverage: getRevisionCoverage(),
    payload: result
  });
  const state = snapshot ? recallSnapshot(snapshot.id, currentWorkbookIdentity(), getWorkbookRevision(), getRevisionCoverage()).state : "not_stored";
  return { ...result, snapshot: snapshot ? { id: snapshot.id, state, capturedAt: snapshot.capturedAt, revision: snapshot.revision } : { state } };
}

async function get_range_details(a: { sheet?: string; address: string }) {
  const address = checkAddress(a.address);
  const revisionAtStart = getWorkbookRevision();
  const result = await Excel.run(async (ctx) => {
    const sheet = sheetOf(ctx, a.sheet);
    const range = await rangeOf(ctx, sheet, address);
    range.load(["address", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
    sheet.load(["id", "name"]);
    sheet.protection.load("protected");
    await ctx.sync();
    const cells = range.rowCount * range.columnCount;
    // Объявлено до второго sync, заполняется после него — см. mergedReport ниже.
    if (cells > MAX_DETAILS_CELLS) {
      throw new ToolError(`Подробности ограничены ${MAX_DETAILS_CELLS} ячейками; ${range.address} содержит ${cells}.`);
    }
    range.load("numberFormat");
    range.format.load(["horizontalAlignment", "verticalAlignment", "wrapText", "rowHeight", "columnWidth"]);
    range.format.font.load(["name", "size", "bold", "italic", "color"]);
    range.format.fill.load(["color"]);
    range.format.protection.load(["locked", "formulaHidden"]);
    const mixedFormat: string[] = [];
    const mergedReport = await probeMergedAreas(ctx, sheet, range);
    const hasValidation = officeCapabilities().pivotTables; // ExcelApi 1.8, same minimum as pivot tables.
    if (hasValidation) range.dataValidation.load(["type", "ignoreBlanks", "valid", "rule", "prompt", "errorAlert"]);
    await ctx.sync();
    return {
      sheetId: sheet.id,
      sheet: sheet.name,
      address: range.address,
      cellCount: cells,
      sheetProtected: sheet.protection.protected,
      format: {
        numberFormat: range.numberFormat,
        ...markMixed({
          horizontalAlignment: range.format.horizontalAlignment,
          verticalAlignment: range.format.verticalAlignment,
          wrapText: range.format.wrapText,
          rowHeight: range.format.rowHeight,
          columnWidth: range.format.columnWidth
        }, "", mixedFormat),
        font: markMixed(range.format.font.toJSON(), "font.", mixedFormat),
        fill: markMixed(range.format.fill.toJSON(), "fill.", mixedFormat),
        protection: markMixed(range.format.protection.toJSON(), "protection.", mixedFormat)
      },
      mixedFormat,
      ...(mixedFormat.length > 0
        ? { formatNote: `Свойства ${mixedFormat.join(", ")} неоднородны внутри ${range.address}: единого значения нет. Это не означает отсутствия оформления — чтобы узнать значение, читайте подробности по однородной части.` }
        : {}),
      mergedAreas: mergedReport.areas,
      ...(mergedReport.unresolvedAnchors.length > 0
        ? {
            mergedAnchorsUnresolved: mergedReport.unresolvedAnchors,
            mergedNote: mergeAnchorNote(mergedReport.unresolvedAnchors, range.address, false)
          }
        : {}),
      /** Сырой ответ Office.js об объединениях: без него отличить «объединения
       * нет» от «мы его не увидели» можно только в отладчике. */
      mergedProbe: mergedReport.diagnostics,
      dataValidation: hasValidation ? range.dataValidation.toJSON() : { state: "unavailable", requiredExcelApi: "1.8" },
      ...(hasValidation && range.dataValidation.type === Excel.DataValidationType.inconsistent
        ? { dataValidationNote: `Внутри ${range.address} правила ввода разные: у части ячеек правило есть, у части нет. Пустые поля правила в этом случае ничего не доказывают — проверяйте нужные ячейки по отдельности.` }
        : {}),
      readAt: new Date().toISOString()
    };
  });
  const snapshot = recordSnapshot({
    kind: "details",
    source: "get_range_details",
    workbook: currentWorkbookIdentity(),
    sheetId: result.sheetId,
    sheetName: result.sheet,
    address: result.address,
    revision: revisionAtStart,
    coverage: getRevisionCoverage(),
    payload: result
  });
  const state = snapshot ? recallSnapshot(snapshot.id, currentWorkbookIdentity(), getWorkbookRevision(), getRevisionCoverage()).state : "not_stored";
  return { ...result, snapshot: snapshot ? { id: snapshot.id, state, capturedAt: snapshot.capturedAt, revision: snapshot.revision } : { state } };
}

async function measure_workbook_export(a: { sliceSizeBytes?: number }) {
  return measureWorkbookExport({ sliceSizeBytes: a?.sliceSizeBytes });
}

async function recall_snapshot(a: { snapshotId: string }) {
  return recallSnapshot(a.snapshotId, currentWorkbookIdentity(), getWorkbookRevision(), getRevisionCoverage());
}

export interface SetRangePlan {
  readonly kind: "set_range_values";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly address: string;
  readonly resolvedAddress: string;
  readonly rows: number;
  readonly columns: number;
  readonly cellCount: number;
  readonly before: readonly (readonly unknown[])[];
  readonly after: readonly (readonly unknown[])[];
  readonly isFormula: boolean;
  readonly replacedFormulaCount: number;
  readonly undoAvailable: boolean;
  readonly errorScanAddress: string;
  readonly calculationMode: string;
  readonly createdAt: string;
  readonly beforeSnapshotId?: string;
  /** Объединения с достоверными границами, задевающие цель. */
  readonly mergedAreas?: readonly string[];
  /** Углы объединений, чья протяжённость неизвестна и может накрывать цель. */
  readonly mergedAnchorsUnresolved?: readonly string[];
  /** Готовая формулировка для предпросмотра: показывается до подтверждения. */
  readonly mergeWarning?: string;
}

function cloneMatrix(matrix: readonly (readonly unknown[])[]): unknown[][] {
  return matrix.map((row) => Array.from(row));
}

function deepFreeze<T>(value: T): T {
  if (value && typeof value === "object") {
    Object.freeze(value);
    for (const child of Object.values(value as Record<string, unknown>)) deepFreeze(child);
  }
  return value;
}

function formulaCount(matrix: readonly (readonly unknown[])[]): number {
  return matrix.flat().filter((value) => typeof value === "string" && value.startsWith("=")).length;
}

function formulaErrors(values: readonly (readonly unknown[])[]): string[] {
  const errors: string[] = [];
  values.forEach((row, rowIndex) => row.forEach((value, columnIndex) => {
    if (typeof value === "string" && /^#(?:DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!|SPILL!|CALC!)/i.test(value)) {
      errors.push(`R${rowIndex + 1}C${columnIndex + 1}:${value}`);
    }
  }));
  return errors;
}

/** Excel interprets a value beginning with '=' as a formula unless escaped. */
export function valuesForLiteralWrite(values: readonly (readonly unknown[])[]): unknown[][] {
  return cloneMatrix(values).map((row) => row.map((value) =>
    typeof value === "string" && value.startsWith("=") ? `'${value}` : value
  ));
}

/** Creates the immutable approval artifact before the user sees the write. */
export async function prepareSetRangePlan(args: unknown): Promise<SetRangePlan> {
  preflightToolArgs("set_range_values", args);
  const a = args as { sheet?: string; address: string; values: unknown[][]; isFormula?: boolean };
  const address = checkAddress(a.address);
  const target = await captureTarget(a.sheet);
  const prepared = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    const range = await rangeOf(ctx, sheet, address);
    const application = ctx.workbook.application;
    sheet.load(["id", "name"]);
    range.load(["address", "rowCount", "columnCount", "rowIndex", "columnIndex", "formulas", "values"]);
    application.load("calculationMode");
    await ctx.sync();
    if (sheet.id !== target.sheetId) throw new ToolError("Целевой лист изменился во время подготовки плана.");
    // Объединение под целью меняет поведение записи, а границы Excel не отдаёт.
    // Предупредить нужно здесь: на предпросмотре у пользователя ещё есть выбор.
    const merged = await probeMergedAreas(ctx, sheet, range);
    const rows = a.values.length;
    const columns = a.values[0].length;
    if (rows !== range.rowCount || columns !== range.columnCount) {
      throw new ToolError(`Размер не совпадает: диапазон ${range.address} — ${range.rowCount}×${range.columnCount}, values — ${rows}×${columns}.`);
    }
    return {
      kind: "set_range_values" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      target: { ...target, sheetName: sheet.name },
      address,
      resolvedAddress: range.address.slice(range.address.lastIndexOf("!") + 1),
      rows,
      columns,
      cellCount: rows * columns,
      before: cloneMatrix(range.formulas as unknown[][]),
      after: cloneMatrix(a.values),
      isFormula: a.isFormula === true,
      replacedFormulaCount: formulaCount(range.formulas as unknown[][]),
      undoAvailable: isCustomUndoAvailable(),
      errorScanAddress: range.address,
      calculationMode: String(application.calculationMode),
      ...(merged.areas.length > 0 ? { mergedAreas: merged.areas } : {}),
      ...(merged.unresolvedAnchors.length > 0 ? { mergedAnchorsUnresolved: merged.unresolvedAnchors } : {}),
      ...(merged.areas.length > 0 || merged.unresolvedAnchors.length > 0
        ? {
            mergeWarning: merged.areas.length > 0
              ? `Цель ${range.address} пересекается с объединёнными областями ${merged.areas.join(", ")}. Запись в объединённую ячейку ведёт себя не так, как в обычную.`
              : mergeAnchorNote(merged.unresolvedAnchors, range.address, true)
          }
        : {}),
      createdAt: new Date().toISOString()
    };
  });
  const snapshot = recordSnapshot({
    kind: "plan",
    source: "set_range_values:before",
    workbook: currentWorkbookIdentity(),
    sheetId: prepared.target.sheetId,
    sheetName: prepared.target.sheetName,
    address: prepared.resolvedAddress,
    revision: getWorkbookRevision(),
    coverage: getRevisionCoverage(),
    payload: prepared.before,
    pin: true
  });
  return deepFreeze({ ...prepared, ...(snapshot ? { beforeSnapshotId: snapshot.id } : {}) });
}

export function releaseSetRangePlanSnapshot(plan: SetRangePlan): void {
  if (plan.beforeSnapshotId) setSnapshotPinned(plan.beforeSnapshotId, false);
}

export interface SetRangesPlan {
  readonly kind: "set_ranges_values";
  readonly id: string;
  readonly items: readonly SetRangePlan[];
  readonly cellCount: number;
  readonly createdAt: string;
}

/** Пересечения ищем по уже разрешённым адресам: аргумент мог быть именем
 * диапазона, и два разных имени способны указывать на одни ячейки. */
export function findOverlappingWrites(items: readonly SetRangePlan[]): [number, number] | null {
  for (let i = 0; i < items.length; i++) {
    const a = parseA1Rect(items[i].resolvedAddress);
    if (!a) continue;
    for (let j = i + 1; j < items.length; j++) {
      if (items[i].target.sheetId !== items[j].target.sheetId) continue;
      const b = parseA1Rect(items[j].resolvedAddress);
      if (b && intersects(a, b)) return [i, j];
    }
  }
  return null;
}

export async function prepareSetRangesPlan(args: unknown): Promise<SetRangesPlan> {
  preflightToolArgs("set_ranges_values", args);
  const a = args as { writes: unknown[] };
  const items: SetRangePlan[] = [];
  try {
    for (let index = 0; index < a.writes.length; index++) {
      try {
        items.push(await prepareSetRangePlan(a.writes[index]));
      } catch (error: any) {
        throw new ToolError(`Операция ${index + 1}: ${error?.message ?? error}. Ни одна запись группы не выполнялась.`);
      }
    }
    const overlap = findOverlappingWrites(items);
    if (overlap) {
      const [i, j] = overlap;
      throw new ToolError(
        `Операции ${i + 1} и ${j + 1} пересекаются: ${items[i].target.sheetName}!${items[i].resolvedAddress} и ` +
        `${items[j].target.sheetName}!${items[j].resolvedAddress}. Порядок записи в общие ячейки неоднозначен, ` +
        `поэтому группа отклонена до записи. Разделите её на отдельные шаги.`
      );
    }
  } catch (error) {
    // Ни одна запись не выполнялась: закреплённые снимки больше не нужны.
    for (const item of items) releaseSetRangePlanSnapshot(item);
    throw error;
  }
  return deepFreeze({
    kind: "set_ranges_values" as const,
    id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
    items,
    cellCount: items.reduce((sum, item) => sum + item.cellCount, 0),
    createdAt: new Date().toISOString()
  });
}

export function releaseSetRangesPlanSnapshots(plan: SetRangesPlan): void {
  for (const item of plan.items) releaseSetRangePlanSnapshot(item);
}

/** Исполняет группу по порядку. Группа не транзакция: после первого сбоя
 * оставшиеся операции не начинаются, а уже выполненные не откатываются —
 * поэтому отчёт разделяет выполненное, невыполненное и неопределённое. */
export async function executeSetRangesPlan(plan: SetRangesPlan) {
  const operations: Record<string, unknown>[] = [];
  let stoppedAt: number | null = null;

  for (let index = 0; index < plan.items.length; index++) {
    const item = plan.items[index];
    const where = `${item.target.sheetName}!${item.resolvedAddress}`;
    if (stoppedAt !== null) {
      releaseSetRangePlanSnapshot(item);
      operations.push({ index: index + 1, address: where, executionState: "not_started", note: "Не начата: группа остановлена раньше." });
      continue;
    }
    try {
      const result = await executeSetRangePlan(item) as Record<string, unknown>;
      operations.push({ index: index + 1, address: where, executionState: result.executionState ?? "verified", cellCount: item.cellCount });
    } catch (error: any) {
      stoppedAt = index;
      const state = error instanceof ToolExecutionError ? error.executionState : "failed_before_write";
      operations.push({ index: index + 1, address: where, executionState: state, error: error?.message ?? String(error) });
    }
  }

  const unknown = operations.some((op) => op.executionState === "unknown");
  const applied = operations.filter((op) => op.executionState === "verified" || op.executionState === "applied").length;
  const executionState = unknown ? "unknown" : stoppedAt === null ? "verified" : applied > 0 ? "applied" : "failed_before_write";
  return {
    ok: stoppedAt === null,
    executionState,
    operations,
    appliedCount: applied,
    note: stoppedAt === null
      ? "Все операции группы выполнены и проверены."
      : `Группа остановлена на операции ${stoppedAt + 1}. Выполненные операции не откатываются автоматически. ` +
        (unknown
          ? "Итог одной из операций неизвестен: перечитайте её диапазон, прежде чем что-либо повторять."
          : "Прежде чем повторять, перечитайте затронутые диапазоны.")
  };
}

export async function executeSetRangePlan(plan: SetRangePlan) {
  try {
    assertWorkbookTarget(plan.target);
    return await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(plan.target.sheetId);
    const range = sheet.getRange(plan.resolvedAddress);
    try {
      sheet.load(["id", "name"]);
      range.load(["address", "rowCount", "columnCount", "formulas", "values"]);
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(
        `Не удалось перечитать цель перед записью: ${error?.message ?? error}. Операция не выполнялась.`,
        "failed_before_write"
      );
    }
    if (sheet.id !== plan.target.sheetId) throw new ToolError("Целевой лист удалён или заменён после предпросмотра. Создайте новый план.");
    if (JSON.stringify(range.formulas) !== JSON.stringify(plan.before)) {
      throw new ToolExecutionError(
        `Диапазон ${sheet.name}!${plan.address} изменился после предпросмотра. Запись остановлена; создайте новый предпросмотр.`,
        "failed_before_write"
      );
    }
    const beforeErrors = formulaErrors(range.values as unknown[][]);
    const undoEnabled = plan.undoAvailable && isCustomUndoAvailable();
    let before = null;
    try {
      before = undoEnabled ? await captureContent(ctx, sheet.name, plan.resolvedAddress) : null;
    } catch (error: any) {
      throw new ToolExecutionError(
        `Не удалось подготовить отмену до записи: ${error?.message ?? error}. Операция не выполнялась.`,
        "failed_before_write"
      );
    }
    const requested = cloneMatrix(plan.after);
    const assigned = plan.isFormula
      ? requested
      : valuesForLiteralWrite(requested);
    try {
      if (plan.isFormula) range.formulas = assigned as any[][];
      else range.values = assigned as any[][];
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(
        `Не удалось определить итог записи ${sheet.name}!${plan.address}: ${error?.message ?? error}. Перечитайте диапазон.`,
        "unknown"
      );
    }

    try {
      range.load(["formulas", "values"]);
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(
        `Запись ${sheet.name}!${plan.address} выполнена, но проверка не удалась: ${error?.message ?? error}. Не повторяйте запись без чтения.`,
        "applied"
      );
    }
    const actual = plan.isFormula ? range.formulas : range.values;
    if (JSON.stringify(actual) !== JSON.stringify(requested)) {
      throw new ToolExecutionError(
        `Запись выполнена, но обратное чтение ${sheet.name}!${plan.address} отличается от плана.`,
        "applied"
      );
    }

    let undoRecorded = false;
    let undoNote: string | undefined;
    if (before) {
      try {
        const after = await captureContent(ctx, sheet.name, plan.resolvedAddress);
        undoRecorded = push(guardedContentUndo(plan.isFormula ? "запись формул" : "запись значений", before, after));
      } catch (error: any) {
        undoNote = `Запись проверена, но точка custom undo не создана: ${error?.message ?? error}`;
      }
    }
    const afterErrors = formulaErrors(range.values as unknown[][]);
    const newErrors = afterErrors.filter((error) => !beforeErrors.includes(error));
    return {
      ok: true,
      planId: plan.id,
      sheetId: sheet.id,
      sheet: sheet.name,
      address: range.address,
      written: `${plan.rows}×${plan.columns}`,
      executionState: "verified" as const,
      verification: {
        address: plan.errorScanAddress,
        matched: true,
        calculationMode: plan.calculationMode,
        calculationSettingsChanged: false,
        existingErrors: beforeErrors,
        newErrors
      },
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: undoNote ?? "Custom undo недоступен или изменился после предпросмотра." })
    };
    });
  } finally {
    releaseSetRangePlanSnapshot(plan);
  }
}

async function set_range_values(a: {
  sheet?: string;
  address: string;
  values: unknown[][];
  isFormula?: boolean;
}) {
  return executeSetRangePlan(await prepareSetRangePlan(a));
}

/** Прямой путь на случай вызова без предпросмотра: подтверждение и разбор
 * плана обеспечивает цикл агента, здесь только связка подготовки и исполнения. */
async function set_ranges_values(a: { writes: unknown[] }) {
  return executeSetRangesPlan(await prepareSetRangesPlan(a));
}

async function insert_rows(a: { sheet?: string; startRow: number; count: number }) {
  const addr = rowsAddress(a.startRow, a.count);
  return Excel.run(async (ctx) => {
    const sheet = sheetOf(ctx, a.sheet);
    sheet.load("name");
    await ctx.sync();

    sheet.getRange(addr).insert(Excel.InsertShiftDirection.down);
    await ctx.sync();

    const invalidatedUndo = invalidateAfterStructuralChange();
    return {
      ok: true,
      sheet: sheet.name,
      inserted: a.count,
      at: a.startRow,
      undoable: false,
      undoNote: "Структурная вставка строк не имеет безопасного custom undo.",
      invalidatedUndo
    };
  });
}

async function delete_rows(a: { sheet?: string; startRow: number; count: number }) {
  const addr = rowsAddress(a.startRow, a.count);
  return Excel.run(async (ctx) => {
    const sheet = sheetOf(ctx, a.sheet);
    sheet.load("name");
    await ctx.sync();

    // Для удаления строк нам не нужен used range вообще. Это устраняет edge case
    // пустого листа и не заставляет Excel сканировать лишний диапазон.
    // Структурное удаление может менять ссылки, таблицы и зависимости по всей
    // книге, поэтому намеренно не регистрируем ложный custom undo.
    sheet.getRange(addr).delete(Excel.DeleteShiftDirection.up);
    await ctx.sync();

    const invalidatedUndo = invalidateAfterStructuralChange();
    return {
      ok: true,
      sheet: sheet.name,
      deleted: a.count,
      from: a.startRow,
      undoable: false,
      undoNote: "Структурное удаление строк не имеет безопасного custom undo.",
      invalidatedUndo
    };
  });
}

async function sort_range(a: {
  sheet?: string;
  address: string;
  column: number;
  ascending?: boolean;
  hasHeaders?: boolean;
}) {
  const address = checkAddress(a.address);
  return Excel.run(async (ctx) => {
    const sheet = sheetOf(ctx, a.sheet);
    const range = sheet.getRange(address);
    range.load(["columnCount", "rowCount", "address"]);
    sheet.load("name");
    await ctx.sync();

    if (range.rowCount * range.columnCount > MAX_IO_CELLS) {
      throw new ToolError(`Для безопасной отмены сортировка ограничена ${MAX_IO_CELLS} ячеек за операцию.`);
    }
    if (a.column < 0 || a.column >= range.columnCount) {
      throw new ToolError(`column=${a.column} вне диапазона: в ${range.address} ${range.columnCount} столбцов.`);
    }

    const undoEnabled = isCustomUndoAvailable();
    const before = undoEnabled ? await captureContent(ctx, sheet.name, address) : null;
    range.sort.apply(
      [{ key: a.column, ascending: a.ascending !== false, sortOn: Excel.SortOn.value }],
      false,
      a.hasHeaders === true
    );
    await ctx.sync();
    let undoRecorded = false;
    if (before) {
      const after = await captureContent(ctx, sheet.name, address);
      undoRecorded = push(guardedContentUndo("сортировка", before, after));
    }

    return {
      ok: true,
      sheet: sheet.name,
      address: range.address,
      byColumn: a.column,
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: "Custom undo недоступен: монитор структурных изменений Excel не активен." })
    };
  });
}

async function apply_filter(a: { sheet?: string; address: string; column: number; criteria: string }) {
  const address = checkAddress(a.address);
  return Excel.run(async (ctx) => {
    const sheet = sheetOf(ctx, a.sheet);
    const range = sheet.getRange(address);
    range.load("columnCount");
    sheet.load("name");
    await ctx.sync();

    if (a.column < 0 || a.column >= range.columnCount) {
      throw new ToolError(`column=${a.column} вне диапазона: в ${address} ${range.columnCount} столбцов.`);
    }

    const raw = String(a.criteria).trim();
    if (!raw) throw new ToolError("criteria не может быть пустым.");
    let criteria: Excel.FilterCriteria;

    if (raw.includes("|")) {
      criteria = { filterOn: Excel.FilterOn.values, values: raw.split("|").map((s) => s.trim()) };
    } else if (/^(>=|<=|<>|>|<|=)/.test(raw)) {
      criteria = { filterOn: Excel.FilterOn.custom, criterion1: raw };
    } else {
      criteria = { filterOn: Excel.FilterOn.values, values: [raw] };
    }

    sheet.autoFilter.apply(range, a.column, criteria);
    await ctx.sync();

    // Фильтрация не меняет данные. Не кладём её в стек «Отменить последнюю
    // правку»: clearCriteria() снял бы чужие фильтры и не восстановил бы
    // предыдущее состояние. Для фильтров нужен отдельный stateful snapshot API.
    return {
      ok: true,
      sheet: sheet.name,
      address,
      column: a.column,
      criteria: raw,
      undoable: false,
      undoNote: "Фильтр не добавлен в custom undo: восстановление прежней комбинации фильтров не гарантируется."
    };
  });
}

async function create_pivot_table(a: {
  sheet?: string;
  destSheet?: string;
  sourceAddress: string;
  destAddress: string;
  rows: string[];
  values: string[];
}) {
  const source = checkAddress(a.sourceAddress);
  const dest = checkAddress(a.destAddress);
  if (!a.rows?.length || !a.values?.length) throw new ToolError("Нужен хотя бы один столбец в rows и values.");

  return Excel.run(async (ctx) => {
    const srcSheet = sheetOf(ctx, a.sheet);
    const dstSheet = a.destSheet ? ctx.workbook.worksheets.getItem(a.destSheet) : srcSheet;
    srcSheet.load("name");
    dstSheet.load("name");
    await ctx.sync();

    const name = `Pivot_${Date.now().toString(36)}`;
    const pivot = ctx.workbook.pivotTables.add(name, srcSheet.getRange(source), dstSheet.getRange(dest));
    for (const r of a.rows) pivot.rowHierarchies.add(pivot.hierarchies.getItem(r));
    for (const v of a.values) pivot.dataHierarchies.add(pivot.hierarchies.getItem(v));
    await ctx.sync();

    const undoRecorded = push(
      action(`создание сводной ${name}`, async () => {
        await Excel.run(async (undoCtx) => {
          undoCtx.workbook.pivotTables.getItem(name).delete();
          await undoCtx.sync();
        });
      })
    );

    return {
      ok: true,
      name,
      sheet: dstSheet.name,
      at: dest,
      rows: a.rows,
      values: a.values,
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: "Custom undo недоступен: монитор структурных изменений Excel не активен." })
    };
  }).catch((e: any) => {
    throw new ToolError(
      `Не удалось создать сводную: ${e?.message ?? e}. Проверь заголовки sourceAddress и имена rows/values.`
    );
  });
}

const CHART_TYPE_KEYS = {
  ColumnClustered: "columnClustered",
  Line: "line",
  Pie: "pie",
  BarClustered: "barClustered",
  XYScatter: "xyscatter",
  Area: "area",
  Doughnut: "doughnut"
} as const;

function resolveChartType(name: string): Excel.ChartType | undefined {
  const key = CHART_TYPE_KEYS[name as keyof typeof CHART_TYPE_KEYS];
  return key ? (Excel.ChartType as any)[key] as Excel.ChartType : undefined;
}

async function create_chart(a: { sheet?: string; address: string; chartType: string; title?: string }) {
  const address = checkAddress(a.address);
  const chartType = resolveChartType(a.chartType);
  if (!chartType) throw new ToolError(`Неподдерживаемый chartType: ${a.chartType}.`);

  return Excel.run(async (ctx) => {
    const sheet = sheetOf(ctx, a.sheet);
    sheet.load("name");
    await ctx.sync();

    const chart = sheet.charts.add(chartType, sheet.getRange(address), Excel.ChartSeriesBy.auto);
    if (a.title) chart.title.text = a.title;
    chart.load("name");
    await ctx.sync();

    const sheetName = sheet.name;
    const chartName = chart.name;
    const undoRecorded = push(
      action(`создание диаграммы ${chartName}`, async () => {
        await Excel.run(async (undoCtx) => {
          undoCtx.workbook.worksheets.getItem(sheetName).charts.getItem(chartName).delete();
          await undoCtx.sync();
        });
      })
    );

    return {
      ok: true,
      sheet: sheet.name,
      chart: chart.name,
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: "Custom undo недоступен: монитор структурных изменений Excel не активен." })
    };
  });
}

async function format_range(a: {
  sheet?: string;
  address: string;
  numberFormat?: string;
  bold?: boolean;
  fillColor?: string;
}) {
  const address = checkAddress(a.address);
  return Excel.run(async (ctx) => {
    const sheet = sheetOf(ctx, a.sheet);
    const range = sheet.getRange(address);
    range.load(["rowCount", "columnCount"]);
    sheet.load("name");
    await ctx.sync();

    if (range.rowCount * range.columnCount > MAX_IO_CELLS) {
      throw new ToolError(`Форматирование ограничено ${MAX_IO_CELLS} ячеек за операцию.`);
    }

    const cellCount = range.rowCount * range.columnCount;
    const canUndoExactly = isCustomUndoAvailable() && cellCount <= MAX_EXACT_FORMAT_UNDO_CELLS;
    const formatSnapshot = canUndoExactly
      ? await captureExactFormat(ctx, sheet.name, address, {
          numberFormat: Boolean(a.numberFormat),
          bold: typeof a.bold === "boolean",
          fillColor: Boolean(a.fillColor)
        })
      : null;

    if (a.numberFormat) {
      range.numberFormat = Array.from({ length: range.rowCount }, () =>
        Array.from({ length: range.columnCount }, () => a.numberFormat as string)
      );
    }
    if (typeof a.bold === "boolean") range.format.font.bold = a.bold;
    if (a.fillColor) range.format.fill.color = hexToColor(a.fillColor);

    await ctx.sync();
    let undoRecorded = false;
    if (formatSnapshot) {
      const afterFormat = await captureExactFormat(ctx, sheet.name, address, {
        numberFormat: Boolean(a.numberFormat),
        bold: typeof a.bold === "boolean",
        fillColor: Boolean(a.fillColor)
      });
      undoRecorded = push(exactFormatUndo("форматирование", formatSnapshot, afterFormat));
    }
    const undoNote = !isCustomUndoAvailable()
      ? "Custom undo недоступен: монитор структурных изменений Excel не активен."
      : cellCount > MAX_EXACT_FORMAT_UNDO_CELLS
        ? `Точный undo форматирования ограничен ${MAX_EXACT_FORMAT_UNDO_CELLS} ячейками.`
        : undefined;
    return {
      ok: true,
      sheet: sheet.name,
      address,
      undoable: undoRecorded,
      ...(undoRecorded || !undoNote ? {} : { undoNote })
    };
  });
}

type Handler = (args: any, options?: { signal?: AbortSignal; deadlineAt?: number }) => Promise<unknown>;

const HANDLERS: Record<ToolName, Handler> = {
  get_active_context,
  list_sheets,
  get_sheet_overview,
  get_range_values,
  search_workbook,
  get_range_details,
  recall_snapshot,
  measure_workbook_export,
  set_range_values,
  set_ranges_values,
  insert_rows,
  delete_rows,
  sort_range,
  apply_filter,
  create_pivot_table,
  create_chart,
  format_range
};

export async function runTool(name: string, args: unknown, options?: { analysisOnly?: boolean; signal?: AbortSignal; deadlineAt?: number }): Promise<unknown> {
  const handler = HANDLERS[name as ToolName];
  if (!handler) {
    throw new ToolError(`Инструмента "${name}" не существует. Доступны: ${Object.keys(HANDLERS).join(", ")}.`);
  }
  const spec = TOOL_BY_NAME.get(name);
  if (options?.analysisOnly && spec?.mutating) {
    throw new ToolError(`Режим «Только анализ» запрещает инструмент ${name}. Операция не выполнялась.`);
  }
  if (spec && !supported(spec)) {
    throw new ToolError(`Инструмент ${name} недоступен в установленной версии Office.js.`);
  }
  if (spec?.mutating && !writableAtCurrentStage(spec)) {
    throw new ToolError(`Инструмент ${name} ещё не подключён к проверяемому контуру этапа 3.`);
  }
  preflightToolArgs(name, args);
  return handler(args, options);
}
