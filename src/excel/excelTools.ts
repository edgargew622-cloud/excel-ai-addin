import {
  action,
  captureContent,
  captureExactFormat,
  exactFormatUndo,
  guardedContentUndo,
  invalidateAfterStructuralChange,
  isCustomUndoAvailable,
  push,
  repairMismatchedCells
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
import { columnLetters, fillFormulaMatrix } from "./formulaFill";
import {
  countFilled,
  countRefErrors,
  deleteImpact,
  insertBlindSpots,
  rowBand,
  usesTableReference,
  type RowBand
} from "./rowOps";
import { createWorkbookBackup, lastWorkbookBackup } from "./workbookBackup";
import {
  conditionText,
  describeCriteria,
  filterChangeKind,
  firstRowLooksLikeHeader,
  isSortedLikeExcel,
  parseFilterCriteria,
  partialRowSortProblem,
  sameAutoFilterState,
  sameRowMultiset,
  sortRowsLikeExcel,
  type AutoFilterState,
  type FilterChange,
  type ParsedFilterCriteria
} from "./sortFilter";

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

/**
 * Отказывает до начала операции, если цель защищена.
 *
 * Защиту видно заранее, и заранее же отказаться честнее: попытка изменения
 * сорвалась бы в Excel, а доказать, что она не началась, было бы нельзя —
 * операция получила бы неопределённый статус на ровном месте. Здесь ничего
 * не выполнялось, и это доказуемо.
 *
 * Требует уже загруженных sheet.protection.protected, range.format.protection.locked
 * и range.address. Свойства читаются мягко: среда без них теряет предпроверку,
 * но не падает.
 */
export function assertTargetWritable(
  sheet: Excel.Worksheet,
  range: Excel.Range,
  change: "values" | "format" = "values"
): void {
  if (!sheet.protection?.protected) return;
  // Защита листа может разрешать оформление заблокированных ячеек. Проверка
  // в Excel 17 сентября 2026 года: агент верно предложил включить такое
  // разрешение, а инструмент отказал бы и после этого — ложный отказ.
  if (change === "format" && (sheet.protection as any)?.options?.allowFormatCells === true) return;
  const locked = range.format?.protection?.locked;
  if (locked === false) return;
  throw new ToolError(
    `Лист ${sheet.name} защищён, а ячейки ${range.address} ` +
    (locked === true ? "заблокированы" : "заблокированы не все одинаково") +
    ". Изменение невозможно, и оно не выполнялось. " +
    (change === "format"
      ? "Снимите защиту листа, разрешите в ней форматирование ячеек или выберите другую цель."
      : "Снимите защиту листа или выберите другую цель.")
  );
}

/** Единая формулировка про неизвестные границы объединений. */
export function mergeAnchorNote(anchors: readonly string[], address: string, beforeWrite: boolean): string {
  const tail = beforeWrite
    ? "Запись в объединённую ячейку ведёт себя не так, как в обычную. Отсутствие объединения здесь не доказано: проверьте цель в Excel перед подтверждением."
    : "Отсутствие объединения здесь не доказано: перед записью проверьте цель в Excel.";
  return `Рядом найдены углы объединений ${anchors.join(", ")}. Excel на этой сборке отдаёт только угол и не сообщает границы, поэтому ${address} может оказаться внутри одного из них. ${tail}`;
}

/** Ошибка ссылки в двух языках Excel — она же считается в rowOps. */
const REF_ERROR = /^#(REF|ССЫЛКА)!$/i;

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

async function create_workbook_backup(a: { sliceSizeBytes?: number }) {
  return createWorkbookBackup({ sliceSizeBytes: a?.sliceSizeBytes });
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
  /** Границы таблиц до записи: по ним видно, расширил ли их Excel. */
  readonly tablesBefore?: readonly TableRange[];
  readonly tableWarning?: string;
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
/**
 * Готовит значения к записи как литералы.
 *
 * Прежде апостроф ставился только перед текстом на «=». Но запись через
 * range.values — это ввод, и Excel распознаёт в тексте числа и даты: «00123»
 * теряет нули, «04.09.2026» в русской локали становится датой. Проверка после
 * записи этого могла не заметить, если Excel возвращал уже преобразованное
 * значение. Апостроф перед любым непустым текстом заставляет хранить его как
 * текст; в ячейке он не отображается. Пустая строка по-прежнему очищает ячейку.
 */
export function valuesForLiteralWrite(values: readonly (readonly unknown[])[]): unknown[][] {
  return cloneMatrix(values).map((row) => row.map((value) =>
    typeof value === "string" && value !== "" ? `'${value}` : value
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
    // Защиту запрашиваем мягко: там, где этих свойств нет, отказ от
    // предпроверки безопаснее падения подготовки плана.
    try {
      range.format?.protection?.load("locked");
      sheet.protection?.load("protected");
    } catch { /* среда без сведений о защите */ }
    application.load("calculationMode");
    await ctx.sync();
    if (sheet.id !== target.sheetId) throw new ToolError("Целевой лист изменился во время подготовки плана.");

    assertTargetWritable(sheet, range);
    // Объединение под целью меняет поведение записи, а границы Excel не отдаёт.
    // Предупредить нужно здесь: на предпросмотре у пользователя ещё есть выбор.
    const merged = await probeMergedAreas(ctx, sheet, range);
    const tables = await readTableRanges(ctx, sheet);
    const tableWarning = tableExpansionWarning(range.address, tables);
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
      tablesBefore: tables,
      ...(tableWarning ? { tableWarning } : {}),
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
      // Формула, записанная в таблицу Excel, может протянуться по столбцу сама.
      // Расхождения дописываются по одной ячейке, затем идёт обычная проверка.
      await repairMismatchedCells(ctx, range, {
        property: plan.isFormula ? "formulas" : "values",
        expected: requested,
        toWrite: assigned
      });
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
      // Замер 16 сентября 2026 года: запись в неугловую ячейку объединённой
      // области Excel принимает молча, но значение никуда не попадает. Такой
      // случай отличим — диапазон совпадает с состоянием до записи — и требует
      // не повтора, а другой цели. Состояние всё равно остаётся applied:
      // доказано лишь то, что цель не изменилась, а не вся книга.
      const unchanged = JSON.stringify(actual) === JSON.stringify(plan.before);
      throw new ToolExecutionError(
        unchanged
          ? `Запись в ${sheet.name}!${plan.address} не дала эффекта: диапазон остался прежним. ` +
            `Частая причина — цель внутри объединённой области: значение принимает только её левая верхняя ячейка. ` +
            `Повтор ничего не изменит; проверьте объединения и выберите другую цель.`
          : `Запись выполнена, но обратное чтение ${sheet.name}!${plan.address} отличается от плана.`,
        "applied"
      );
    }

    const tableChanges = describeTableChanges(plan.tablesBefore ?? [], await readTableRanges(ctx, sheet));
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
      ...(tableChanges.length ? { tableChanges, tableNote: "Excel изменил границы таблицы из-за этой записи; в отчёте это нужно назвать." } : {}),
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

/* ---------------------------------------------------------------------------
 * Вставка и удаление строк
 *
 * Самые опасные операции этапа: они меняют адреса по всему листу, у них нет
 * отката, и Excel выполняет их молча. Проверка целевых ячеек здесь бесполезна:
 * ломается не цель, а формулы в других местах книги. Поэтому план сначала
 * обходит формулы книги и называет, что именно сломается, а исполнение
 * сверяет число ошибок ссылок до и после.
 * ------------------------------------------------------------------------- */

const MAX_ROW_PREVIEW = 8;
const MAX_RISKS_REPORTED = 20;

export interface RowFormulaRisk {
  readonly sheet: string;
  readonly address: string;
  readonly formula: string;
  /** broken — станет #ССЫЛКА!, shrunk — диапазон уменьшится, missed — не охватит новые строки. */
  readonly kind: "broken" | "shrunk" | "missed";
  readonly reference: string;
}

export interface RowOpPlan {
  readonly kind: "insert_rows" | "delete_rows";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly startRow: number;
  readonly count: number;
  /** Адрес полосы строк, например 5:6. */
  readonly rowsAddress: string;
  readonly usedRangeAddress?: string;
  readonly usedRangeRows: number;
  readonly usedRangeColumns: number;
  /** Содержимое полосы для предпросмотра; пусто, если полоса слишком велика. */
  readonly preview: readonly (readonly unknown[])[];
  readonly previewTruncated: boolean;
  /** Непустые ячейки полосы — то, что удаление уничтожит без возврата. */
  readonly filledCells: number;
  /** Слепок полосы для сверки перед исполнением; null — полоса не снималась. */
  readonly bandSignature: string | null;
  readonly formulaRisks: readonly RowFormulaRisk[];
  readonly riskOverflow: number;
  /** Листы, формулы которых обойти не удалось: об этом нужно сказать прямо. */
  readonly unscannedSheets: readonly string[];
  readonly tableFormulaSheets: readonly string[];
  readonly refErrorsBefore: number;
  readonly tablesBefore: readonly TableRange[];
  readonly tableWarning?: string;
  readonly mergeWarning?: string;
  readonly backup: { name: string; at: string } | null;
  readonly undoAvailable: false;
  readonly undoNote: string;
  readonly createdAt: string;
}

interface ScannedSheet {
  name: string;
  address: string;
  rowIndex: number;
  columnIndex: number;
  formulas: unknown[][];
  values: unknown[][];
}

/** Формулы всех листов книги: по ним видно, что сломает операция. */
async function scanWorkbookFormulas(
  ctx: Excel.RequestContext
): Promise<{ sheets: ScannedSheet[]; unscanned: string[] }> {
  const collection = ctx.workbook.worksheets;
  collection.load("items/name");
  await ctx.sync();
  const items = [...collection.items];
  const unscanned: string[] = [];
  if (items.length > MAX_SEARCH_SHEETS) {
    return { sheets: [], unscanned: items.map((sheet) => sheet.name) };
  }

  const capabilities = officeCapabilities();
  const used = items.map((sheet) => capabilities.usedRangeOrNull
    ? sheet.getUsedRangeOrNullObject(true)
    : sheet.getUsedRange(true));
  for (const range of used) range.load(["address", "rowIndex", "columnIndex", "rowCount", "columnCount", "isNullObject"]);
  await ctx.sync();

  const wanted: { name: string; range: Excel.Range }[] = [];
  items.forEach((sheet, index) => {
    const range = used[index];
    if ((range as any).isNullObject) return;
    if (range.rowCount * range.columnCount > MAX_IO_CELLS) { unscanned.push(sheet.name); return; }
    range.load(["formulas", "values"]);
    wanted.push({ name: sheet.name, range });
  });
  if (wanted.length) await ctx.sync();

  return {
    sheets: wanted.map((item) => ({
      name: item.name,
      address: String(item.range.address),
      rowIndex: item.range.rowIndex,
      columnIndex: item.range.columnIndex,
      formulas: item.range.formulas as unknown[][],
      values: item.range.values as unknown[][]
    })),
    unscanned
  };
}

function scannedRefErrors(sheets: readonly ScannedSheet[]): number {
  return sheets.reduce((total, sheet) => total + countRefErrors(sheet.values), 0);
}

/** Где именно стоит формула — адрес нужен, чтобы пользователь её нашёл. */
function cellAddressOf(sheet: ScannedSheet, row: number, column: number): string {
  return `${columnLetters(sheet.columnIndex + column + 1)}${sheet.rowIndex + row + 1}`;
}

export function collectRowRisks(
  mode: "insert_rows" | "delete_rows",
  sheets: readonly ScannedSheet[],
  targetSheet: string,
  band: RowBand
): { risks: RowFormulaRisk[]; overflow: number; tableFormulaSheets: string[] } {
  const risks: RowFormulaRisk[] = [];
  const tableFormulaSheets = new Set<string>();
  let overflow = 0;

  for (const sheet of sheets) {
    sheet.formulas.forEach((row, rowIndex) => {
      row.forEach((formula, columnIndex) => {
        if (typeof formula !== "string" || !formula.startsWith("=")) return;
        // Формула внутри удаляемой полосы исчезнет вместе с ней: называть её
        // пострадавшей — значит пугать пользователя тем, чего не будет.
        const ownRow = sheet.rowIndex + rowIndex + 1;
        const insideBand = sheet.name === targetSheet && ownRow >= band.startRow && ownRow <= band.endRow;
        if (mode === "delete_rows" && insideBand) return;
        if (usesTableReference(formula)) tableFormulaSheets.add(sheet.name);
        const address = cellAddressOf(sheet, rowIndex, columnIndex);
        const add = (kind: RowFormulaRisk["kind"], reference: string) => {
          if (risks.length >= MAX_RISKS_REPORTED) { overflow += 1; return; }
          risks.push({ sheet: sheet.name, address, formula, kind, reference });
        };
        if (mode === "delete_rows") {
          const impact = deleteImpact(formula, sheet.name, targetSheet, band);
          for (const reference of impact.broken) add("broken", reference.text);
          for (const reference of impact.shrunk) add("shrunk", reference.text);
        } else {
          for (const reference of insertBlindSpots(formula, sheet.name, targetSheet, band)) {
            add("missed", reference.text);
          }
        }
      });
    });
  }
  return { risks, overflow, tableFormulaSheets: [...tableFormulaSheets] };
}

async function prepareRowOpPlan(mode: "insert_rows" | "delete_rows", args: unknown): Promise<RowOpPlan> {
  preflightToolArgs(mode, args);
  const a = args as { sheet?: string; startRow: number; count: number };
  const address = rowsAddress(a.startRow, a.count);
  const band = rowBand(a.startRow, a.count);
  const target = await captureTarget(a.sheet);

  const prepared = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    sheet.load(["id", "name"]);
    try { sheet.protection?.load("protected"); } catch { /* среда без сведений о защите */ }
    const used = officeCapabilities().usedRangeOrNull
      ? sheet.getUsedRangeOrNullObject(true)
      : sheet.getUsedRange(true);
    used.load(["address", "rowIndex", "columnIndex", "rowCount", "columnCount", "isNullObject"]);
    await ctx.sync();
    if (sheet.id !== target.sheetId) throw new ToolError("Целевой лист изменился во время подготовки плана.");
    if (sheet.protection?.protected) {
      throw new ToolError(
        `Лист ${sheet.name} защищён: строки вставить или удалить нельзя, и операция не выполнялась. Снимите защиту листа.`
      );
    }

    const empty = Boolean((used as any).isNullObject);
    const usedRows = empty ? 0 : used.rowCount;
    const usedColumns = empty ? 0 : used.columnCount;
    const lastUsedRow = empty ? 0 : used.rowIndex + used.rowCount;

    // Содержимое полосы берём только в пределах занятой области: целые строки
    // листа — это 16 384 столбца, и читать их незачем.
    let preview: unknown[][] = [];
    let previewTruncated = false;
    let filledCells = 0;
    let bandSignature: string | null = null;
    const overlapStart = Math.max(band.startRow, empty ? 1 : used.rowIndex + 1);
    const overlapEnd = Math.min(band.endRow, lastUsedRow);
    if (!empty && overlapEnd >= overlapStart && usedColumns > 0) {
      const rows = overlapEnd - overlapStart + 1;
      if (rows * usedColumns <= MAX_IO_CELLS) {
        const bandRange = sheet.getRangeByIndexes(overlapStart - 1, used.columnIndex, rows, usedColumns);
        bandRange.load(["values", "formulas"]);
        await ctx.sync();
        const values = bandRange.values as unknown[][];
        filledCells = countFilled(values);
        bandSignature = JSON.stringify(bandRange.formulas);
        preview = values.slice(0, MAX_ROW_PREVIEW).map((row) => row.slice(0, MAX_ROW_PREVIEW));
        previewTruncated = values.length > MAX_ROW_PREVIEW || usedColumns > MAX_ROW_PREVIEW;
      } else {
        previewTruncated = true;
      }
    }

    // Опрос объединений читает у диапазона его положение, поэтому диапазон
    // нужно не только создать, но и загрузить: проверка 18 сентября 2026 года
    // сорвалась на этом до записи.
    const probeColumns = Math.max(1, Math.min(usedColumns || 1, 100));
    const bandForProbe = sheet.getRangeByIndexes(band.startRow - 1, 0, a.count, probeColumns);
    bandForProbe.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
    await ctx.sync();
    const merged = await probeMergedAreas(ctx, sheet, bandForProbe);
    const tables = await readTableRanges(ctx, sheet);
    const tableWarning = tableExpansionWarning(`${sheet.name}!${address}`, tables);
    const scan = await scanWorkbookFormulas(ctx);
    const { risks, overflow, tableFormulaSheets } = collectRowRisks(mode, scan.sheets, sheet.name, band);
    const backup = lastWorkbookBackup();

    return {
      kind: mode,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      target: { ...target, sheetName: sheet.name },
      startRow: a.startRow,
      count: a.count,
      rowsAddress: address,
      ...(empty ? {} : { usedRangeAddress: String(used.address) }),
      usedRangeRows: usedRows,
      usedRangeColumns: usedColumns,
      preview,
      previewTruncated,
      filledCells,
      bandSignature,
      formulaRisks: risks,
      riskOverflow: overflow,
      unscannedSheets: scan.unscanned,
      tableFormulaSheets,
      refErrorsBefore: scannedRefErrors(scan.sheets),
      tablesBefore: tables,
      ...(tableWarning ? { tableWarning } : {}),
      ...(merged.areas.length > 0 || merged.unresolvedAnchors.length > 0
        ? { mergeWarning: "Полосу задевают объединённые ячейки. Excel может отказать в операции или разорвать объединение." }
        : {}),
      backup: backup ? { name: backup.name, at: backup.at } : null,
      undoAvailable: false as const,
      undoNote: mode === "delete_rows"
        ? "Отмены нет: удалённые строки не восстанавливаются ни кнопкой отмены панели, ни повтором операции. " +
          (backup
            ? `Вернуться можно только к резервной копии ${backup.name}.`
            : "Резервной копии в этом сеансе не создавалось — возвращаться будет не к чему.")
        : "Отмены нет: вставленные строки панель удалить обратно не может, а вся предыдущая история отмены будет очищена.",
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared) as RowOpPlan;
}

export const prepareInsertRowsPlan = (args: unknown) => prepareRowOpPlan("insert_rows", args);
export const prepareDeleteRowsPlan = (args: unknown) => prepareRowOpPlan("delete_rows", args);

export async function executeRowOpPlan(plan: RowOpPlan) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(plan.target.sheetId);
    sheet.load(["id", "name"]);
    const used = officeCapabilities().usedRangeOrNull
      ? sheet.getUsedRangeOrNullObject(true)
      : sheet.getUsedRange(true);
    used.load(["rowIndex", "columnIndex", "rowCount", "columnCount", "isNullObject"]);
    await ctx.sync();
    if (sheet.id !== plan.target.sheetId) {
      throw new ToolExecutionError("Целевой лист изменился после предпросмотра. Строки не трогали.", "failed_before_write");
    }

    // Сверка полосы: у операции нет отката, поэтому ручная правка между
    // предпросмотром и подтверждением обязана отменить операцию, а не пройти.
    if (plan.bandSignature !== null) {
      const empty = Boolean((used as any).isNullObject);
      const overlapStart = Math.max(plan.startRow, empty ? 1 : used.rowIndex + 1);
      const overlapEnd = Math.min(plan.startRow + plan.count - 1, empty ? 0 : used.rowIndex + used.rowCount);
      const rows = overlapEnd - overlapStart + 1;
      if (empty || rows < 1 || used.columnCount !== plan.usedRangeColumns) {
        throw new ToolExecutionError(
          `Занятая область листа ${sheet.name} изменилась после предпросмотра. Строки не трогали — сделайте новый предпросмотр.`,
          "failed_before_write"
        );
      }
      const bandRange = sheet.getRangeByIndexes(overlapStart - 1, used.columnIndex, rows, used.columnCount);
      bandRange.load("formulas");
      await ctx.sync();
      if (JSON.stringify(bandRange.formulas) !== plan.bandSignature) {
        throw new ToolExecutionError(
          `Строки ${sheet.name}!${plan.rowsAddress} изменились после предпросмотра. Операция не выполнялась — сделайте новый предпросмотр.`,
          "failed_before_write"
        );
      }
    }

    const range = sheet.getRange(plan.rowsAddress);
    try {
      if (plan.kind === "insert_rows") range.insert(Excel.InsertShiftDirection.down);
      else range.delete(Excel.DeleteShiftDirection.up);
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(
        `Excel отказал в операции со строками ${sheet.name}!${plan.rowsAddress}: ${error?.message ?? error}. ` +
        "Неизвестно, успела ли она примениться — перечитайте лист, прежде чем что-либо менять.",
        "unknown"
      );
    }

    // Целевые строки после такой операции всегда выглядят правильно. Смотреть
    // надо на остальную книгу: туда уходят сломанные ссылки.
    const scan = await scanWorkbookFormulas(ctx);
    const refErrorsAfter = scannedRefErrors(scan.sheets);
    const newRefErrors = refErrorsAfter - plan.refErrorsBefore;
    const brokenCells: string[] = [];
    for (const item of scan.sheets) {
      item.values.forEach((row, rowIndex) => {
        row.forEach((value, columnIndex) => {
          if (typeof value === "string" && REF_ERROR.test(value.trim()) && brokenCells.length < MAX_RISKS_REPORTED) {
            brokenCells.push(`${item.name}!${cellAddressOf(item, rowIndex, columnIndex)}`);
          }
        });
      });
    }
    const tableChanges = describeTableChanges(plan.tablesBefore, await readTableRanges(ctx, sheet));
    const invalidatedUndo = invalidateAfterStructuralChange();

    return {
      ok: true,
      executionState: "verified",
      sheet: sheet.name,
      rows: plan.rowsAddress,
      ...(plan.kind === "insert_rows"
        ? { inserted: plan.count, at: plan.startRow }
        : { deleted: plan.count, from: plan.startRow, lostFilledCells: plan.filledCells }),
      refErrorsBefore: plan.refErrorsBefore,
      refErrorsAfter,
      // Разбор последствий живёт в плане, а план модель не видит: ей показывают
      // только результат. Проверка 18 сентября 2026 года: после вставки строки
      // под данными агент отчитался «формулы не изменились» и умолчал, что две
      // суммы на другом листе перестали охватывать данные. Поэтому найденное
      // на предпросмотре уезжает в результат целиком.
      ...(plan.formulaRisks.length
        ? {
            affectedFormulas: plan.formulaRisks,
            ...(plan.riskOverflow > 0 ? { affectedFormulasOmitted: plan.riskOverflow } : {}),
            affectedFormulasNote: plan.kind === "insert_rows"
              ? "Эти формулы не охватывают вставленные строки: ошибки не будет, итог просто посчитан без них. " +
                "Ошибок ссылок такие случаи не дают, поэтому назови их пользователю поимённо — сам он их не увидит."
              : "Эти формулы ссылались на удалённые строки. Перечисли их пользователю: часть станет #ССЫЛКА!, " +
                "а укоротившиеся диапазоны молча считают по меньшему числу строк и никакой ошибки не показывают."
          }
        : {}),
      ...(plan.tableFormulaSheets.length
        ? {
            tableFormulaSheets: plan.tableFormulaSheets,
            tableFormulaNote: "На этих листах есть формулы со ссылками на таблицы; они не разбирались. Скажи, что про них ничего не проверено."
          }
        : {}),
      ...(newRefErrors > 0
        ? {
            newRefErrors,
            brokenCells,
            refNote: `В книге появилось ${newRefErrors} ошибок ссылок — формулы указывали на изменённые строки. Это нужно назвать пользователю.`
          }
        : {}),
      ...(scan.unscanned.length
        ? { unscannedSheets: scan.unscanned, scanNote: "Эти листы слишком велики для обхода формул: про них ничего не проверено." }
        : {}),
      ...(tableChanges.length
        ? { tableChanges, tableNote: "Excel изменил границы таблицы из-за этой операции; в отчёте это нужно назвать." }
        : {}),
      undoable: false,
      undoNote: plan.undoNote,
      invalidatedUndo
    };
  });
}

async function insert_rows(a: { sheet?: string; startRow: number; count: number }) {
  return executeRowOpPlan(await prepareInsertRowsPlan(a));
}

async function delete_rows(a: { sheet?: string; startRow: number; count: number }) {
  return executeRowOpPlan(await prepareDeleteRowsPlan(a));
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

/** Что именно просят изменить в оформлении. Остальные свойства не трогаются. */
export interface FormatRequest {
  numberFormat?: string;
  bold?: boolean;
  fillColor?: string;
}

/** Состояние тех же свойств до операции. Значение либо однородно по области,
 * либо равно null: Office.js так сообщает о неоднородности. Подавать null как
 * «не задано» нельзя — по этому снимку ловится ручная правка перед запуском. */
export interface FormatSnapshot {
  numberFormat?: unknown;
  bold?: unknown;
  fillColor?: unknown;
}

export interface FormatRangePlan {
  readonly kind: "format_range";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly address: string;
  readonly resolvedAddress: string;
  readonly cellCount: number;
  readonly request: FormatRequest;
  readonly before: FormatSnapshot;
  readonly expected: FormatSnapshot;
  readonly undoAvailable: boolean;
  readonly undoNote?: string;
  readonly createdAt: string;
  readonly mergedAreas?: readonly string[];
  readonly mergedAnchorsUnresolved?: readonly string[];
  readonly mergeWarning?: string;
}

export function requestedFormatKeys(request: FormatRequest): (keyof FormatRequest)[] {
  const keys: (keyof FormatRequest)[] = [];
  if (typeof request.numberFormat === "string") keys.push("numberFormat");
  if (typeof request.bold === "boolean") keys.push("bold");
  if (typeof request.fillColor === "string") keys.push("fillColor");
  return keys;
}

/** Читает только запрошенные свойства: сравнивать остальные незачем,
 * а лишние загрузки удлиняют операцию на больших областях. */
async function readFormatSnapshot(
  ctx: Excel.RequestContext,
  range: Excel.Range,
  keys: (keyof FormatRequest)[]
): Promise<FormatSnapshot> {
  if (keys.includes("numberFormat")) range.load("numberFormat");
  if (keys.includes("bold")) range.format.font.load("bold");
  if (keys.includes("fillColor")) range.format.fill.load("color");
  await ctx.sync();
  const snapshot: FormatSnapshot = {};
  if (keys.includes("numberFormat")) {
    const value = range.numberFormat as unknown;
    snapshot.numberFormat = Array.isArray(value) ? (value as any[][])[0]?.[0] ?? null : value ?? null;
  }
  if (keys.includes("bold")) snapshot.bold = range.format.font.bold ?? null;
  if (keys.includes("fillColor")) snapshot.fillColor = range.format.fill.color ?? null;
  return snapshot;
}

/** Цвета Excel возвращает в своём написании регистра, поэтому строки
 * сравниваются без учёта регистра, а прочее — строго. */
export function sameFormatValue(a: unknown, b: unknown): boolean {
  return typeof a === "string" && typeof b === "string"
    ? canonicalFormatText(a) === canonicalFormatText(b)
    : a === b;
}

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

export function formatSnapshotsEqual(a: FormatSnapshot, b: FormatSnapshot): boolean {
  const keys = new Set([...Object.keys(a), ...Object.keys(b)]) as Set<keyof FormatSnapshot>;
  for (const key of keys) if (!sameFormatValue(a[key], b[key])) return false;
  return true;
}

/** Каким станет оформление, если операция пройдёт: это показывает предпросмотр
 * и с этим же сверяется результат. */
export function expectedFormatSnapshot(request: FormatRequest): FormatSnapshot {
  const expected: FormatSnapshot = {};
  if (typeof request.numberFormat === "string") expected.numberFormat = request.numberFormat;
  if (typeof request.bold === "boolean") expected.bold = request.bold;
  if (typeof request.fillColor === "string") expected.fillColor = hexToColor(request.fillColor);
  return expected;
}

export async function prepareFormatRangePlan(args: unknown): Promise<FormatRangePlan> {
  preflightToolArgs("format_range", args);
  const a = args as { sheet?: string; address: string } & FormatRequest;
  const request: FormatRequest = {
    ...(typeof a.numberFormat === "string" ? { numberFormat: a.numberFormat } : {}),
    ...(typeof a.bold === "boolean" ? { bold: a.bold } : {}),
    ...(typeof a.fillColor === "string" ? { fillColor: a.fillColor } : {})
  };
  const keys = requestedFormatKeys(request);
  if (keys.length === 0) throw new ToolError("Не указано ни одного свойства оформления: менять нечего.");

  const address = checkAddress(a.address);
  const target = await captureTarget(a.sheet);

  const prepared = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    const range = await rangeOf(ctx, sheet, address);
    range.load(["address", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
    sheet.load(["id", "name"]);
    try {
      range.format?.protection?.load("locked");
      sheet.protection?.load(["protected", "options"]);
    } catch { /* среда без сведений о защите */ }
    await ctx.sync();
    if (sheet.id !== target.sheetId) throw new ToolError("Целевой лист изменился во время подготовки плана.");

    const cells = range.rowCount * range.columnCount;
    if (cells > MAX_IO_CELLS) {
      throw new ToolError(`Форматирование ограничено ${MAX_IO_CELLS} ячеек за операцию; ${range.address} содержит ${cells}.`);
    }
    assertTargetWritable(sheet, range, "format");

    const merged = await probeMergedAreas(ctx, sheet, range);
    const before = await readFormatSnapshot(ctx, range, keys);

    const exactUndo = isCustomUndoAvailable() && cells <= MAX_EXACT_FORMAT_UNDO_CELLS;
    const undoNote = !isCustomUndoAvailable()
      ? "Отмена недоступна: монитор изменений Excel не активен."
      : !exactUndo
        ? `Точная отмена оформления ограничена ${MAX_EXACT_FORMAT_UNDO_CELLS} ячейками, а здесь ${cells}. Операция выполнится, но откатить её автоматически будет нечем.`
        : undefined;

    return {
      kind: "format_range" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      target: { ...target, sheetName: sheet.name },
      address,
      resolvedAddress: range.address.slice(range.address.lastIndexOf("!") + 1),
      cellCount: cells,
      request,
      before,
      expected: expectedFormatSnapshot(request),
      undoAvailable: exactUndo,
      ...(undoNote ? { undoNote } : {}),
      ...(merged.areas.length > 0 ? { mergedAreas: merged.areas } : {}),
      ...(merged.unresolvedAnchors.length > 0
        ? {
            mergedAnchorsUnresolved: merged.unresolvedAnchors,
            mergeWarning: mergeAnchorNote(merged.unresolvedAnchors, range.address, true)
          }
        : {}),
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

/** Сколько ячеек показывать модели для опоры в отчёте. */
const GROUNDING_SAMPLE = 5;

/**
 * Возвращает то, по чему модель пишет отчёт: заголовок над областью и первые
 * значения самой области.
 *
 * Проверка в Excel 17 сентября 2026 года: после оформления `D2:D6` агент
 * описал в отчёте значения соседнего столбца — он пересказывал по памяти из
 * чтения трёх столбцов, потому что в ответе операции самих значений не было.
 * Отчёт должен опираться на результат операции, а не на память.
 *
 * Выборка вспомогательная: если её не удалось прочитать, операция не падает.
 */
async function groundingSample(
  ctx: Excel.RequestContext,
  sheet: Excel.Worksheet,
  range: Excel.Range
): Promise<Record<string, unknown>> {
  try {
    if (typeof sheet.getRangeByIndexes !== "function") return {};
    const rows = Math.min(range.rowCount, GROUNDING_SAMPLE);
    const columns = Math.min(range.columnCount, GROUNDING_SAMPLE);
    const sample = sheet.getRangeByIndexes(range.rowIndex, range.columnIndex, rows, columns);
    sample.load(["values", "text"]);
    const header = range.rowIndex > 0
      ? sheet.getRangeByIndexes(range.rowIndex - 1, range.columnIndex, 1, columns)
      : null;
    header?.load("values");
    await ctx.sync();
    const truncated = range.rowCount > rows || range.columnCount > columns;
    return {
      ...(header && Array.isArray(header.values) ? { headerAbove: header.values[0] } : {}),
      ...(Array.isArray(sample.values) ? { sampleValues: sample.values } : {}),
      // Отображаемый текст берётся из Excel, а не выводится из кода формата:
      // в русской локали 900 с форматом 0.00 выглядит как «900,00».
      ...(Array.isArray(sample.text) ? { sampleText: sample.text } : {}),
      ...(truncated ? { sampleNote: `Показаны первые ${rows}×${columns} ячеек из ${range.rowCount}×${range.columnCount}.` } : {})
    };
  } catch {
    return {};
  }
}

export async function executeFormatRangePlan(plan: FormatRangePlan) {
  const keys = requestedFormatKeys(plan.request);

  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(plan.target.sheetId);
    const range = sheet.getRange(plan.resolvedAddress);
    range.load(["address", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
    sheet.load(["id", "name"]);
    await ctx.sync();
    if (sheet.id !== plan.target.sheetId) {
      throw new ToolExecutionError("Целевой лист изменился после предпросмотра. Оформление не менялось.", "failed_before_write");
    }

    // Та же защита от гонки, что и у записи значений: между предпросмотром
    // и подтверждением оформление могли поменять руками.
    const current = await readFormatSnapshot(ctx, range, keys);
    if (!formatSnapshotsEqual(current, plan.before)) {
      throw new ToolExecutionError(
        `Оформление ${sheet.name}!${plan.resolvedAddress} изменилось после предпросмотра. Операция не выполнялась — сделайте новый предпросмотр.`,
        "failed_before_write"
      );
    }

    const snapshot = plan.undoAvailable
      ? await captureExactFormat(ctx, sheet.name, plan.resolvedAddress, {
          numberFormat: keys.includes("numberFormat"),
          bold: keys.includes("bold"),
          fillColor: keys.includes("fillColor")
        })
      : null;

    try {
      if (typeof plan.request.numberFormat === "string") {
        range.numberFormat = Array.from({ length: range.rowCount }, () =>
          Array.from({ length: range.columnCount }, () => plan.request.numberFormat as string)
        );
      }
      if (typeof plan.request.bold === "boolean") range.format.font.bold = plan.request.bold;
      if (typeof plan.request.fillColor === "string") range.format.fill.color = hexToColor(plan.request.fillColor);
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(
        `Не удалось определить итог форматирования ${sheet.name}!${plan.resolvedAddress}: ${error?.message ?? error}. Перечитайте оформление диапазона.`,
        "unknown"
      );
    }

    const after = await readFormatSnapshot(ctx, range, keys);
    if (!formatSnapshotsEqual(after, plan.expected)) {
      // Тот же разбор, что и у записи значений: «ничего не изменилось»
      // и «изменилось не так» — разные случаи, и повтор помогает только во втором.
      const unchanged = formatSnapshotsEqual(after, plan.before);
      throw new ToolExecutionError(
        unchanged
          ? `Форматирование ${sheet.name}!${plan.resolvedAddress} не дало эффекта: оформление осталось прежним. ` +
            `Повтор ничего не изменит; проверьте защиту листа и объединения.`
          : `Оформление применено, но обратное чтение ${sheet.name}!${plan.resolvedAddress} отличается от плана: ${JSON.stringify(after)}. ` +
            `Excel мог привести значение к своему виду.`,
        "applied"
      );
    }

    let undoRecorded = false;
    if (snapshot) {
      try {
        const afterFormat = await captureExactFormat(ctx, sheet.name, plan.resolvedAddress, {
          numberFormat: keys.includes("numberFormat"),
          bold: keys.includes("bold"),
          fillColor: keys.includes("fillColor")
        });
        undoRecorded = push(exactFormatUndo("форматирование", snapshot, afterFormat));
      } catch { undoRecorded = false; }
    }

    const grounding = await groundingSample(ctx, sheet, range);
    return {
      ok: true,
      executionState: "verified",
      sheet: sheet.name,
      address: plan.resolvedAddress,
      cellCount: plan.cellCount,
      applied: plan.request,
      before: plan.before,
      // Прочитано из Excel после операции, а не повторено из запроса: Excel
      // переписывает код формата по-своему, и отчёт должен опираться на факт.
      actual: after,
      ...grounding,
      ...(plan.mergedAreas?.length || plan.mergedAnchorsUnresolved?.length
        ? { cellCountNote: `Область задевает объединённые ячейки: видимых ячеек может быть меньше ${plan.cellCount}. Оформление объединения Excel хранит в его левой верхней ячейке.` }
        : {}),
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: plan.undoNote ?? "Автоматическая отмена этой операции недоступна." })
    };
  });
}

/** Сколько первых строк показывать в предпросмотре сортировки. */
const SORT_PREVIEW_ROWS = 5;

export interface SortRangePlan {
  readonly kind: "sort_range";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly address: string;
  readonly resolvedAddress: string;
  readonly rows: number;
  readonly columns: number;
  readonly cellCount: number;
  readonly column: number;
  readonly ascending: boolean;
  readonly hasHeaders: boolean;
  readonly keyHeader?: unknown;
  /** Состояние области до сортировки: по нему ловится ручная правка. */
  readonly beforeFormulas: readonly (readonly unknown[])[];
  readonly beforeValues: readonly (readonly unknown[])[];
  /** Первые строки данных сейчас и после сортировки, как её ожидаем мы. */
  readonly previewBefore: readonly (readonly unknown[])[];
  readonly previewAfter: readonly (readonly unknown[])[];
  readonly formulaCount: number;
  readonly formulaWarning?: string;
  readonly headerWarning?: string;
  readonly undoAvailable: boolean;
  readonly undoNote?: string;
  readonly mergeWarning?: string;
  readonly createdAt: string;
}

function rectOfAddress(address: string) {
  const rect = parseA1Rect(address.slice(address.lastIndexOf("!") + 1));
  return rect && rect.kind === "cells" ? rect : null;
}

export async function prepareSortRangePlan(args: unknown): Promise<SortRangePlan> {
  preflightToolArgs("sort_range", args);
  const a = args as { sheet?: string; address: string; column: number; ascending?: boolean; hasHeaders?: boolean; allowPartialRows?: boolean };
  const address = checkAddress(a.address);
  const target = await captureTarget(a.sheet);
  const ascending = a.ascending !== false;
  const hasHeaders = a.hasHeaders === true;

  const prepared = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    const range = await rangeOf(ctx, sheet, address);
    range.load(["address", "rowCount", "columnCount", "rowIndex", "columnIndex", "formulas", "values"]);
    sheet.load(["id", "name"]);
    try {
      range.format?.protection?.load("locked");
      sheet.protection?.load("protected");
    } catch { /* среда без сведений о защите */ }
    // Сплошной блок вокруг области: по нему видно, не режет ли сортировка строки.
    let region: Excel.Range | null = null;
    try {
      if (typeof range.getSurroundingRegion === "function") {
        region = range.getSurroundingRegion();
        region.load("address");
      }
    } catch { region = null; }
    await ctx.sync();
    if (sheet.id !== target.sheetId) throw new ToolError("Целевой лист изменился во время подготовки плана.");

    const cells = range.rowCount * range.columnCount;
    if (cells > MAX_IO_CELLS) {
      throw new ToolError(`Сортировка ограничена ${MAX_IO_CELLS} ячеек за операцию; ${range.address} содержит ${cells}.`);
    }
    if (!Number.isInteger(a.column) || a.column < 0 || a.column >= range.columnCount) {
      throw new ToolError(`column=${a.column} вне области: в ${range.address} ${range.columnCount} столбцов, отсчёт с 0.`);
    }
    const dataRows = range.rowCount - (hasHeaders ? 1 : 0);
    if (dataRows < 2) throw new ToolError("Сортировать нечего: в области меньше двух строк данных.");
    assertTargetWritable(sheet, range);

    // Главная опасность сортировки: переставить часть столбцов, оставив соседние
    // на месте. Excel не ругается, а строки перемешиваются молча.
    const problem = partialRowSortProblem(
      rectOfAddress(range.address) ?? { rowStart: 0, rowEnd: 0, columnStart: 0, columnEnd: 0 },
      region?.address ? rectOfAddress(region.address) : null
    );
    if (problem && a.allowPartialRows !== true) {
      throw new ToolError(
        `${range.address} — ${problem}${region?.address ? ` ${region.address}` : ""}. Сортировка переставит только эти столбцы, ` +
        `а соседние данные в тех же строках останутся на месте, и строки перемешаются. Операция не выполнялась. ` +
        `Укажите всю область${region?.address ? ` ${region.address.slice(region.address.lastIndexOf("!") + 1)}` : ""}; ` +
        `если нужно сортировать именно часть, спросите пользователя и повторите с allowPartialRows.`
      );
    }

    const merged = await probeMergedAreas(ctx, sheet, range);
    const values = range.values as unknown[][];
    const formulas = range.formulas as unknown[][];
    const data = hasHeaders ? values.slice(1) : values;
    const expected = sortRowsLikeExcel(data, a.column, ascending);
    const formulasInside = formulaCount(formulas);
    const keyHeader = hasHeaders ? values[0]?.[a.column] : undefined;

    // Если первая строка похожа на заголовки, а флаг не выставлен, заголовок
    // уедет в середину данных — это частая и неприятная ошибка.
    const looksLikeHeader = !hasHeaders && firstRowLooksLikeHeader(values);

    const undo = isCustomUndoAvailable();
    return {
      kind: "sort_range" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      target: { ...target, sheetName: sheet.name },
      address,
      resolvedAddress: range.address.slice(range.address.lastIndexOf("!") + 1),
      rows: range.rowCount,
      columns: range.columnCount,
      cellCount: cells,
      column: a.column,
      ascending,
      hasHeaders,
      ...(keyHeader !== undefined ? { keyHeader } : {}),
      beforeFormulas: cloneMatrix(formulas),
      beforeValues: cloneMatrix(values),
      previewBefore: data.slice(0, SORT_PREVIEW_ROWS).map((row) => [...row]),
      previewAfter: expected.slice(0, SORT_PREVIEW_ROWS).map((row) => [...row]),
      formulaCount: formulasInside,
      ...(formulasInside > 0
        ? { formulaWarning: `В области ${formulasInside} формул. Ссылки внутри строки Excel сдвинет вместе с ней, но ссылки на другие строки после сортировки могут указывать не туда.` }
        : {}),
      ...(looksLikeHeader
        ? { headerWarning: "Первая строка похожа на заголовки, но hasHeaders не выставлен: она будет отсортирована вместе с данными." }
        : {}),
      undoAvailable: undo,
      ...(undo ? {} : { undoNote: "Отмена недоступна: монитор изменений Excel не активен." }),
      ...(merged.areas.length > 0 || merged.unresolvedAnchors.length > 0
        ? { mergeWarning: "В области или рядом есть объединённые ячейки. Excel отказывается сортировать объединения разного размера." }
        : {}),
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executeSortRangePlan(plan: SortRangePlan) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(plan.target.sheetId);
    const range = sheet.getRange(plan.resolvedAddress);
    range.load(["address", "rowCount", "columnCount", "rowIndex", "columnIndex", "formulas", "values"]);
    sheet.load(["id", "name"]);
    await ctx.sync();
    if (sheet.id !== plan.target.sheetId) {
      throw new ToolExecutionError("Целевой лист изменился после предпросмотра. Сортировка не выполнялась.", "failed_before_write");
    }
    if (JSON.stringify(range.formulas) !== JSON.stringify(plan.beforeFormulas)) {
      throw new ToolExecutionError(
        `Данные ${sheet.name}!${plan.resolvedAddress} изменились после предпросмотра. Сортировка не выполнялась — сделайте новый предпросмотр.`,
        "failed_before_write"
      );
    }

    const before = plan.undoAvailable ? await captureContent(ctx, sheet.name, plan.resolvedAddress) : null;

    try {
      range.sort.apply(
        [{ key: plan.column, ascending: plan.ascending, sortOn: Excel.SortOn.value }],
        false,
        plan.hasHeaders
      );
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(
        `Не удалось определить итог сортировки ${sheet.name}!${plan.resolvedAddress}: ${error?.message ?? error}. Перечитайте область.`,
        "unknown"
      );
    }

    range.load(["values", "formulas"]);
    await ctx.sync();
    const valuesAfter = range.values as unknown[][];
    const dataBefore = plan.hasHeaders ? plan.beforeValues.slice(1) : plan.beforeValues;
    const dataAfter = plan.hasHeaders ? valuesAfter.slice(1) : valuesAfter;

    if (plan.hasHeaders && JSON.stringify(valuesAfter[0]) !== JSON.stringify(plan.beforeValues[0])) {
      throw new ToolExecutionError(
        `Сортировка ${sheet.name}!${plan.resolvedAddress} выполнена, но строка заголовков сдвинулась. Перечитайте область.`,
        "applied"
      );
    }
    // Жёсткая проверка: строки сохранились целиком. Нарушение — порча данных.
    if (!sameRowMultiset(dataBefore, dataAfter)) {
      const unchanged = JSON.stringify(valuesAfter) === JSON.stringify(plan.beforeValues);
      throw new ToolExecutionError(
        unchanged
          ? `Сортировка ${sheet.name}!${plan.resolvedAddress} не дала эффекта: порядок остался прежним. Повтор ничего не изменит.`
          : `Сортировка ${sheet.name}!${plan.resolvedAddress} выполнена, но набор строк после неё не совпадает с исходным. ` +
            `Это может означать перемешанные строки или пересчёт формул со ссылками на другие строки. Перечитайте область, прежде чем что-либо менять.`,
        "applied"
      );
    }

    let undoRecorded = false;
    if (before) {
      try {
        const after = await captureContent(ctx, sheet.name, plan.resolvedAddress);
        undoRecorded = push(guardedContentUndo("сортировка", before, after));
      } catch { undoRecorded = false; }
    }

    // Мягкая проверка: порядок по нашей оценке. Текст Excel сравнивает по правилам
    // локали, поэтому расхождение не объявляется ошибкой, а называется.
    const ordered = isSortedLikeExcel(dataAfter, plan.column, plan.ascending);
    const grounding = await groundingSample(ctx, sheet, range as any);
    return {
      ok: true,
      executionState: "verified",
      sheet: sheet.name,
      address: plan.resolvedAddress,
      rows: plan.rows,
      column: plan.column,
      ...(plan.keyHeader !== undefined ? { keyHeader: plan.keyHeader } : {}),
      ascending: plan.ascending,
      rowsPreserved: true,
      firstRowsAfter: dataAfter.slice(0, SORT_PREVIEW_ROWS),
      ...(ordered ? {} : { orderNote: "Строки сохранены, но порядок ключевого столбца отличается от ожидаемого нами: Excel сравнивает текст по правилам своей локали. Проверьте порядок глазами." }),
      ...grounding,
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: plan.undoNote ?? "Автоматическая отмена этой сортировки недоступна." })
    };
  });
}

export interface ApplyFilterPlan {
  readonly kind: "apply_filter";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly address: string;
  readonly resolvedAddress: string;
  readonly column: number;
  readonly columnHeader?: unknown;
  readonly criteria: ParsedFilterCriteria;
  readonly criteriaText: string;
  readonly before: AutoFilterState;
  readonly rows: number;
  readonly visibleRowsBefore: number | null;
  readonly replacesExisting: boolean;
  /** new — фильтра не было; adds — условие добавится к существующим;
   * replacesColumn — заменится условие этого столбца; replacesFilter —
   * фильтр другой области будет заменён целиком. */
  readonly change: FilterChange;
  readonly createdAt: string;
}

async function readAutoFilterState(ctx: Excel.RequestContext, sheet: Excel.Worksheet): Promise<AutoFilterState> {
  return (await readAutoFilterDetails(ctx, sheet)).state;
}

/** Состояние фильтра вместе с тем, что нужно отчёту: действующие условия
 * с заголовками столбцов и сырой ответ Excel. Сырой ответ нужен, чтобы по
 * одной проверке было видно, как Excel заполняет столбцы без условия. */
async function readAutoFilterDetails(ctx: Excel.RequestContext, sheet: Excel.Worksheet): Promise<{
  state: AutoFilterState;
  conditions: { column: number; header?: unknown; condition: string }[];
  raw: unknown[];
}> {
  const filter = sheet.autoFilter;
  filter.load(["enabled", "criteria"]);
  const filterRange = filter.getRangeOrNullObject();
  filterRange.load(["isNullObject", "address"]);
  await ctx.sync();
  const raw = Array.isArray(filter.criteria) ? (filter.criteria as unknown[]) : [];
  const described = describeCriteria(raw);
  let headers: unknown[] = [];
  if (!filterRange.isNullObject && typeof (filterRange as any).getRow === "function") {
    try {
      const headerRow = filterRange.getRow(0);
      headerRow.load("values");
      await ctx.sync();
      headers = (headerRow.values as unknown[][])[0] ?? [];
    } catch { headers = []; }
  }
  return {
    state: {
      enabled: Boolean(filter.enabled),
      address: filterRange.isNullObject ? null : String(filterRange.address),
      activeColumns: described.activeColumns,
      activeIndexes: described.activeIndexes,
      criteria: described.text
    },
    conditions: described.activeIndexes.map((index) => ({
      column: index,
      ...(headers[index] !== undefined ? { header: headers[index] } : {}),
      condition: conditionText(raw[index])
    })),
    raw
  };
}

async function visibleRowCount(ctx: Excel.RequestContext, range: Excel.Range): Promise<number | null> {
  try {
    const view = range.getVisibleView();
    view.load("rowCount");
    await ctx.sync();
    return typeof view.rowCount === "number" ? view.rowCount : null;
  } catch {
    return null;
  }
}

export async function prepareApplyFilterPlan(args: unknown): Promise<ApplyFilterPlan> {
  preflightToolArgs("apply_filter", args);
  const a = args as { sheet?: string; address: string; column: number; criteria: string };
  const address = checkAddress(a.address);
  let criteria: ParsedFilterCriteria;
  try { criteria = parseFilterCriteria(a.criteria); }
  catch (error: any) { throw new ToolError(error?.message ?? String(error)); }
  const target = await captureTarget(a.sheet);

  const prepared = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    const range = await rangeOf(ctx, sheet, address);
    range.load(["address", "rowCount", "columnCount", "values"]);
    sheet.load(["id", "name"]);
    const tables = sheet.tables;
    tables.load("items/name");
    await ctx.sync();
    if (sheet.id !== target.sheetId) throw new ToolError("Целевой лист изменился во время подготовки плана.");
    if (!Number.isInteger(a.column) || a.column < 0 || a.column >= range.columnCount) {
      throw new ToolError(`column=${a.column} вне области: в ${range.address} ${range.columnCount} столбцов, отсчёт с 0.`);
    }

    // У таблиц Excel свой фильтр в каждом столбце; автофильтр листа поверх
    // таблицы Excel не ставит, а попытка выглядела бы как сбой.
    const rect = rectOfAddress(range.address);
    const tableRanges = tables.items.map((table) => {
      const tableRange = table.getRange();
      tableRange.load("address");
      return { name: table.name, range: tableRange };
    });
    await ctx.sync();
    for (const table of tableRanges) {
      const tableRect = rectOfAddress(table.range.address);
      if (rect && tableRect && intersects(rect, tableRect)) {
        throw new ToolError(
          `${range.address} пересекается с таблицей Excel «${table.name}» (${table.range.address}). ` +
          `У таблицы свой фильтр, и автофильтр листа поверх неё не ставится. Операция не выполнялась.`
        );
      }
    }

    const before = await readAutoFilterState(ctx, sheet);
    const visibleRowsBefore = await visibleRowCount(ctx, range);
    const resolvedAddress = range.address.slice(range.address.lastIndexOf("!") + 1);
    const values = range.values as unknown[][];
    return {
      kind: "apply_filter" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      target: { ...target, sheetName: sheet.name },
      address,
      resolvedAddress,
      column: a.column,
      ...(values[0]?.[a.column] !== undefined ? { columnHeader: values[0][a.column] } : {}),
      criteria,
      criteriaText: String(a.criteria).trim(),
      before,
      rows: range.rowCount,
      visibleRowsBefore,
      // На листе один автофильтр: новый на другой области заменит прежний
      // вместе со всеми его условиями, и это нужно показать до подтверждения.
      replacesExisting: filterChangeKind(before, resolvedAddress, a.column) === "replacesFilter",
      change: filterChangeKind(before, resolvedAddress, a.column),
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executeApplyFilterPlan(plan: ApplyFilterPlan) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(plan.target.sheetId);
    const range = sheet.getRange(plan.resolvedAddress);
    sheet.load(["id", "name"]);
    range.load("address");
    await ctx.sync();
    if (sheet.id !== plan.target.sheetId) {
      throw new ToolExecutionError("Целевой лист изменился после предпросмотра. Фильтр не менялся.", "failed_before_write");
    }
    const current = await readAutoFilterState(ctx, sheet);
    if (!sameAutoFilterState(current, plan.before)) {
      throw new ToolExecutionError(
        `Фильтр на листе ${sheet.name} изменился после предпросмотра. Операция не выполнялась — сделайте новый предпросмотр.`,
        "failed_before_write"
      );
    }

    const criteria: Excel.FilterCriteria = plan.criteria.filterOn === "custom"
      ? { filterOn: Excel.FilterOn.custom, criterion1: plan.criteria.criterion1 }
      : { filterOn: Excel.FilterOn.values, values: [...(plan.criteria.values ?? [])] };
    try {
      sheet.autoFilter.apply(range, plan.column, criteria);
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(
        `Не удалось определить итог фильтра на ${sheet.name}!${plan.resolvedAddress}: ${error?.message ?? error}. Перечитайте состояние фильтра.`,
        "unknown"
      );
    }

    const details = await readAutoFilterDetails(ctx, sheet);
    const after = details.state;
    const expectedRect = rectOfAddress(plan.resolvedAddress);
    const actualRect = after.address ? rectOfAddress(after.address) : null;
    const coversTarget = Boolean(expectedRect && actualRect && intersects(expectedRect, actualRect));
    if (!after.enabled || !coversTarget || after.activeColumns === 0) {
      throw new ToolExecutionError(
        `Фильтр на ${sheet.name}!${plan.resolvedAddress} применён, но обратное чтение его не подтверждает: ${JSON.stringify(after)}.`,
        "applied"
      );
    }
    const visibleRowsAfter = await visibleRowCount(ctx, range);
    return {
      ok: true,
      executionState: "verified",
      sheet: sheet.name,
      address: plan.resolvedAddress,
      filterAddress: after.address,
      column: plan.column,
      ...(plan.columnHeader !== undefined ? { columnHeader: plan.columnHeader } : {}),
      criteria: plan.criteriaText,
      // «rows» путался с видимыми строками: агент отчитался «строк до 4, после 4»
      // и решил, что ничего не скрыто, хотя видно было две строки из четырёх.
      areaRows: plan.rows,
      visibleRowsBefore: plan.visibleRowsBefore,
      visibleRowsAfter,
      ...(typeof visibleRowsAfter === "number" ? { hiddenRowsAfter: plan.rows - visibleRowsAfter } : {}),
      conditionsAfter: details.conditions,
      criteriaRaw: details.raw,
      filterChange: plan.change,
      ...(plan.change === "replacesFilter" ? { replacedFilter: plan.before } : {}),
      ...(plan.change === "adds" ? { note: "Условие добавлено к уже стоящим условиям фильтра; прежние условия сохранены." } : {}),
      undoable: false,
      undoNote: "Фильтр данных не меняет, но прежнюю комбинацию условий автоматически не вернуть. Снять фильтр можно в Excel: Данные → Очистить."
    };
  });
}

async function sort_range(a: unknown) {
  return executeSortRangePlan(await prepareSortRangePlan(a));
}

async function apply_filter(a: unknown) {
  return executeApplyFilterPlan(await prepareApplyFilterPlan(a));
}

export interface TableRange {
  name: string;
  address: string;
}

/**
 * Предупреждение о том, что запись расширит таблицу Excel.
 *
 * Проверка в Excel 18 сентября 2026 года: запись формул в `H2:H6` рядом
 * с таблицей `SalesTable` (A1:G6) молча расширила её до `A1:H6` и добавила
 * столбец с автоматическим заголовком «Столбец1». Проверка результата смотрела
 * только на целевые ячейки и структурного изменения не заметила.
 *
 * Таблица растёт вправо и вниз, поэтому опасны соседство справа по тем же
 * строкам и снизу по тем же столбцам. Пересечение с таблицей — отдельный
 * случай: запись попадает внутрь неё.
 */
export function tableExpansionWarning(targetAddress: string, tables: readonly TableRange[]): string | null {
  const target = parseA1Rect(withoutSheet(targetAddress));
  if (!target) return null;
  const touching: string[] = [];
  const inside: string[] = [];
  for (const table of tables) {
    const rect = parseA1Rect(withoutSheet(table.address));
    if (!rect) continue;
    if (intersects(rect, target)) { inside.push(`${table.name} (${table.address})`); continue; }
    const sameRows = rect.rowStart <= target.rowEnd && rect.rowEnd >= target.rowStart;
    const sameColumns = rect.columnStart <= target.columnEnd && rect.columnEnd >= target.columnStart;
    const rightOf = sameRows && target.columnStart === rect.columnEnd + 1;
    const below = sameColumns && target.rowStart === rect.rowEnd + 1;
    if (rightOf || below) touching.push(`${table.name} (${table.address})`);
  }
  if (inside.length) {
    return `Цель находится внутри таблицы Excel ${inside.join(", ")}. Запись пойдёт в ячейки таблицы; формула в столбце таблицы может протянуться на весь столбец.`;
  }
  if (touching.length) {
    return `Цель вплотную примыкает к таблице Excel ${touching.join(", ")}. Excel расширит таблицу на эту область и добавит столбец или строку с автоматическим заголовком.`;
  }
  return null;
}

/** Границы таблиц листа: нужны и для предупреждения, и для сверки после операции. */
async function readTableRanges(ctx: Excel.RequestContext, sheet: Excel.Worksheet): Promise<TableRange[]> {
  try {
    const tables = sheet.tables;
    tables.load("items/name");
    await ctx.sync();
    const ranges = tables.items.map((table) => {
      const range = table.getRange();
      range.load("address");
      return { name: table.name, range };
    });
    if (!ranges.length) return [];
    await ctx.sync();
    return ranges.map((item) => ({ name: item.name, address: String(item.range.address) }));
  } catch {
    return [];
  }
}

/** Что стало с таблицами после операции: расширение видно только сравнением. */
function describeTableChanges(before: readonly TableRange[], after: readonly TableRange[]) {
  const changes: { name: string; before: string; after: string }[] = [];
  for (const item of after) {
    const previous = before.find((table) => table.name === item.name);
    if (previous && previous.address !== item.address) {
      changes.push({ name: item.name, before: previous.address, after: item.address });
    }
  }
  for (const item of after) {
    if (!before.some((table) => table.name === item.name)) {
      changes.push({ name: item.name, before: "не было", after: item.address });
    }
  }
  return changes;
}

export interface FillRangePlan {
  readonly kind: "fill_range";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly address: string;
  readonly resolvedAddress: string;
  readonly anchorAddress: string;
  readonly rows: number;
  readonly columns: number;
  readonly cellCount: number;
  readonly value: string | number | boolean;
  readonly isFormula: boolean;
  readonly beforeFormulas: readonly (readonly unknown[])[];
  /** Сколько непустых ячеек будет затёрто: это главное последствие операции. */
  readonly occupiedCells: number;
  readonly sampleBefore: readonly (readonly unknown[])[];
  /** Границы таблиц до операции: по ним видно, расширил ли их Excel. */
  readonly tablesBefore: readonly TableRange[];
  readonly tableWarning?: string;
  readonly undoAvailable: boolean;
  readonly undoNote?: string;
  readonly mergeWarning?: string;
  readonly createdAt: string;
}

/** Левая верхняя ячейка области: с неё Excel начинает заполнение. */
export function anchorOf(address: string): string {
  const rect = parseA1Rect(address);
  if (!rect || rect.kind !== "cells") return address;
  const column = (index: number): string => {
    let value = "";
    let left = index;
    while (left > 0) {
      const remainder = (left - 1) % 26;
      value = String.fromCharCode(65 + remainder) + value;
      left = Math.floor((left - 1) / 26);
    }
    return value;
  };
  return `${column(rect.columnStart)}${rect.rowStart}`;
}

export async function prepareFillRangePlan(args: unknown): Promise<FillRangePlan> {
  preflightToolArgs("fill_range", args);
  const a = args as { sheet?: string; address: string; value: string | number | boolean; isFormula?: boolean };
  const address = checkAddress(a.address);
  const isFormula = a.isFormula === true;
  if (isFormula && !(typeof a.value === "string" && a.value.startsWith("="))) {
    throw new ToolError("При isFormula=true значение должно быть формулой, начинающейся со знака равенства.");
  }
  const target = await captureTarget(a.sheet);

  const prepared = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    const range = await rangeOf(ctx, sheet, address);
    range.load(["address", "rowCount", "columnCount", "rowIndex", "columnIndex", "formulas", "values"]);
    sheet.load(["id", "name"]);
    try {
      range.format?.protection?.load("locked");
      sheet.protection?.load("protected");
    } catch { /* среда без сведений о защите */ }
    await ctx.sync();
    if (sheet.id !== target.sheetId) throw new ToolError("Целевой лист изменился во время подготовки плана.");

    const cells = range.rowCount * range.columnCount;
    if (cells > MAX_IO_CELLS) {
      throw new ToolError(`Заполнение ограничено ${MAX_IO_CELLS} ячеек за операцию; ${range.address} содержит ${cells}.`);
    }
    if (typeof (range as any).autoFill !== "function" && cells > 1) {
      throw new ToolError("Эта сборка Excel не поддерживает заполнение диапазона (Range.autoFill). Операция не выполнялась.");
    }
    assertTargetWritable(sheet, range);

    const merged = await probeMergedAreas(ctx, sheet, range);
    const tables = await readTableRanges(ctx, sheet);
    const tableWarning = tableExpansionWarning(range.address, tables);
    const formulas = cloneMatrix(range.formulas as unknown[][]);
    const occupied = formulas.flat().filter((cell) => cell !== "" && cell !== null && cell !== undefined).length;
    const resolvedAddress = range.address.slice(range.address.lastIndexOf("!") + 1);

    return {
      kind: "fill_range" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      target: { ...target, sheetName: sheet.name },
      address,
      resolvedAddress,
      anchorAddress: anchorOf(resolvedAddress),
      rows: range.rowCount,
      columns: range.columnCount,
      cellCount: cells,
      value: a.value,
      isFormula,
      beforeFormulas: formulas,
      occupiedCells: occupied,
      sampleBefore: formulas.slice(0, 5).map((row) => row.slice(0, 5)),
      tablesBefore: tables,
      ...(tableWarning ? { tableWarning } : {}),
      undoAvailable: isCustomUndoAvailable(),
      ...(isCustomUndoAvailable() ? {} : { undoNote: "Отмена недоступна: монитор изменений Excel не активен." }),
      ...(merged.areas.length > 0 || merged.unresolvedAnchors.length > 0
        ? { mergeWarning: mergeAnchorNote(merged.unresolvedAnchors.length ? merged.unresolvedAnchors : merged.areas, range.address, true) }
        : {}),
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

/**
 * Заполняет область одной формулой или значением.
 *
 * Формулы строит Excel своим заполнением: записывается только первая ячейка,
 * остальное протягивается через `autoFill`, и относительные ссылки Excel
 * подстраивает сам. Модель при этом передаёт одну формулу вместо массива
 * на тысячи ячеек, который упирался в предел длины её ответа.
 */
export async function executeFillRangePlan(plan: FillRangePlan) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(plan.target.sheetId);
    const range = sheet.getRange(plan.resolvedAddress);
    const anchor = sheet.getRange(plan.anchorAddress);
    range.load(["address", "rowCount", "columnCount", "rowIndex", "columnIndex", "formulas", "values"]);
    sheet.load(["id", "name"]);
    await ctx.sync();
    if (sheet.id !== plan.target.sheetId) {
      throw new ToolExecutionError("Целевой лист изменился после предпросмотра. Заполнение не выполнялось.", "failed_before_write");
    }
    if (JSON.stringify(range.formulas) !== JSON.stringify(plan.beforeFormulas)) {
      throw new ToolExecutionError(
        `Данные ${sheet.name}!${plan.resolvedAddress} изменились после предпросмотра. Заполнение не выполнялось — сделайте новый предпросмотр.`,
        "failed_before_write"
      );
    }

    const before = plan.undoAvailable ? await captureContent(ctx, sheet.name, plan.resolvedAddress) : null;
    const assigned = plan.isFormula ? plan.value : valuesForLiteralWrite([[plan.value]])[0][0];

    let filledBy: "autoFill" | "formulas" = "autoFill";
    try {
      if (plan.isFormula) anchor.formulas = [[assigned]] as any[][];
      else anchor.values = [[assigned]] as any[][];
      // Проверка в Excel 18 сентября 2026 года: запись первой ячейки и протяжка
      // в одном пакете дают внутреннюю ошибку Excel — он тянет то, чего ещё
      // не видит. Между ними нужна синхронизация.
      await ctx.sync();
      if (plan.cellCount > 1) {
        try {
          anchor.autoFill(range, Excel.AutoFillType.fillDefault);
          await ctx.sync();
        } catch (fillError: any) {
          // Та же проверка: рядом с таблицей Excel отвечает на протяжку
          // внутренней ошибкой. Тогда строим те же формулы сами и пишем их
          // обычной записью — результат совпадает с протяжкой за угол.
          console.warn(`autoFill не сработал (${fillError?.message ?? fillError}); заполняю формулами построчно`);
          filledBy = "formulas";
          if (plan.isFormula) {
            range.formulas = fillFormulaMatrix(String(plan.value), plan.rows, plan.columns) as any[][];
          } else {
            range.values = Array.from({ length: plan.rows }, () =>
              Array.from({ length: plan.columns }, () => assigned)) as any[][];
          }
          await ctx.sync();
        }
      }
    } catch (error: any) {
      throw new ToolExecutionError(
        `Не удалось определить итог заполнения ${sheet.name}!${plan.resolvedAddress}: ${error?.message ?? error}. Перечитайте область.`,
        "unknown"
      );
    }

    range.load(["formulas", "values"]);
    await ctx.sync();
    const formulasAfter = range.formulas as unknown[][];
    const empty = formulasAfter.flat().filter((cell) => cell === "" || cell === null || cell === undefined).length;
    const anchorAfter = formulasAfter[0]?.[0];

    if (JSON.stringify(formulasAfter) === JSON.stringify(plan.beforeFormulas)) {
      throw new ToolExecutionError(
        `Заполнение ${sheet.name}!${plan.resolvedAddress} не дало эффекта: область осталась прежней. ` +
        `Повтор ничего не изменит; проверьте защиту листа и объединения.`,
        "applied"
      );
    }
    if (empty > 0) {
      throw new ToolExecutionError(
        `Заполнение ${sheet.name}!${plan.resolvedAddress} прошло частично: ${empty} ячеек остались пустыми. Перечитайте область.`,
        "applied"
      );
    }
    if (plan.isFormula && String(anchorAfter) !== String(plan.value)) {
      throw new ToolExecutionError(
        `Первая ячейка ${sheet.name}!${plan.anchorAddress} содержит ${JSON.stringify(anchorAfter)} вместо запрошенной формулы. Перечитайте область.`,
        "applied"
      );
    }

    let undoRecorded = false;
    if (before) {
      try {
        const after = await captureContent(ctx, sheet.name, plan.resolvedAddress);
        undoRecorded = push(guardedContentUndo(plan.isFormula ? "заполнение формулой" : "заполнение значением", before, after));
      } catch { undoRecorded = false; }
    }

    const grounding = await groundingSample(ctx, sheet, range as any);
    const tableChanges = describeTableChanges(plan.tablesBefore, await readTableRanges(ctx, sheet));
    return {
      ok: true,
      executionState: "verified",
      sheet: sheet.name,
      address: plan.resolvedAddress,
      cellCount: plan.cellCount,
      filledWith: plan.value,
      filledBy,
      ...(filledBy === "formulas"
        ? { fillNote: "Протяжка Excel не сработала, поэтому формулы построены и записаны панелью; ссылки подставлены так же, как при протяжке." }
        : {}),
      ...(tableChanges.length ? { tableChanges, tableNote: "Excel изменил границы таблицы из-за этой записи; в отчёте это нужно назвать." } : {}),
      isFormula: plan.isFormula,
      // Формулы Excel подстроил под каждую строку сам: видно по краям области.
      firstFormula: formulasAfter[0]?.[0],
      lastFormula: formulasAfter[formulasAfter.length - 1]?.[(formulasAfter[0]?.length ?? 1) - 1],
      overwrittenCells: plan.occupiedCells,
      ...grounding,
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: plan.undoNote ?? "Автоматическая отмена этого заполнения недоступна." })
    };
  });
}

async function fill_range(a: unknown) {
  return executeFillRangePlan(await prepareFillRangePlan(a));
}

/** Прямой путь на случай вызова без предпросмотра: подтверждение и разбор
 * плана обеспечивает цикл агента, здесь только связка подготовки и исполнения. */
async function format_range(a: unknown) {
  return executeFormatRangePlan(await prepareFormatRangePlan(a));
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
  create_workbook_backup,
  set_range_values,
  set_ranges_values,
  fill_range,
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
