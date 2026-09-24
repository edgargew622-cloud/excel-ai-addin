import { getRevisionCoverage, getWorkbookRevision } from "./workbookRevision";

const workbookSessionId = globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`;
const MAX_SHEETS = 100;
const MAX_OBJECTS = 50;

export interface WorkbookTarget {
  workbookSessionId: string;
  documentUrl: string;
  sheetId: string;
  sheetName: string;
}

export interface OfficeCapabilities {
  excelApi11: boolean;
  selectedRanges: boolean;
  usedRangeOrNull: boolean;
  pivotTables: boolean;
  filters: boolean;
}

function supported(version: string): boolean {
  try { return Office.context.requirements.isSetSupported("ExcelApi", version); }
  catch { return false; }
}

export function officeCapabilities(): OfficeCapabilities {
  return {
    excelApi11: supported("1.1"),
    selectedRanges: supported("1.9"),
    usedRangeOrNull: supported("1.4"),
    pivotTables: supported("1.8"),
    filters: supported("1.9")
  };
}

function documentUrl(): string {
  try { return String(Office.context.document.url ?? ""); }
  catch { return ""; }
}

export function currentWorkbookIdentity() {
  return { workbookSessionId, documentUrl: documentUrl() };
}

export async function getActiveContext() {
  const capabilities = officeCapabilities();
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    sheet.load(["id", "name"]);
    await ctx.sync();

    // Проверка в Excel 24 сентября 2026 года (S6): после create_chart Excel
    // выделяет диаграмму, и активная ячейка и выделение отказывают
    // с InvalidSelection. Прежде это роняло каждую следующую просьбу ещё до
    // модели, пока пользователь не щёлкнет по ячейке. Выделение не ячейки —
    // сведение о книге, а не ошибка: так и сообщается.
    let activeCell: string | null = null;
    let selectedAreas: string[] = [];
    let selectionNote: string | undefined;
    try {
      const cell = ctx.workbook.getActiveCell();
      cell.load("address");
      const selection: Excel.Range | Excel.RangeAreas = capabilities.selectedRanges
        ? ctx.workbook.getSelectedRanges()
        : ctx.workbook.getSelectedRange();
      selection.load("address");
      await ctx.sync();
      activeCell = cell.address;
      selectedAreas = String(selection.address).split(",").map((part) => part.trim()).filter(Boolean);
    } catch (error: any) {
      if (error?.code !== "InvalidSelection") throw error;
      let chart: string | null = null;
      try {
        const active = (ctx.workbook as any).getActiveChartOrNullObject?.();
        if (active) {
          active.load(["isNullObject", "name"]);
          await ctx.sync();
          if (!active.isNullObject) chart = String(active.name);
        }
      } catch { /* сборка без getActiveChartOrNullObject: имя диаграммы неизвестно */ }
      selectionNote = chart
        ? `Выделена диаграмма «${chart}», а не ячейки: активной ячейки и выделенного диапазона нет. Адреса берите из просьбы или из обзора листа.`
        : "Выделены не ячейки (диаграмма, фигура или другой объект): активной ячейки и выделенного диапазона нет. Адреса берите из просьбы или из обзора листа.";
    }

    return {
      workbook: { sessionId: workbookSessionId, documentUrl: documentUrl(), identityConfirmed: true },
      activeSheet: { id: sheet.id, name: sheet.name },
      activeCell,
      selectedAreas,
      ...(selectionNote ? { selectionNote } : {}),
      readAt: new Date().toISOString(),
      revision: getWorkbookRevision(),
      revisionCoverage: getRevisionCoverage(),
      capabilities
    };
  });
}

export async function captureTarget(sheetName?: string): Promise<WorkbookTarget> {
  return Excel.run(async (ctx) => {
    const sheet = sheetName?.trim()
      ? ctx.workbook.worksheets.getItem(sheetName.trim())
      : ctx.workbook.worksheets.getActiveWorksheet();
    sheet.load(["id", "name"]);
    await ctx.sync();
    return {
      workbookSessionId,
      documentUrl: documentUrl(),
      sheetId: sheet.id,
      sheetName: sheet.name
    };
  });
}

export function assertWorkbookTarget(target: WorkbookTarget): void {
  if (target.workbookSessionId !== workbookSessionId || target.documentUrl !== documentUrl()) {
    throw new Error("Открытая книга изменилась после предпросмотра. Создайте новый план.");
  }
}

export async function listSheets() {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/id,name,position,visibility");
    await ctx.sync();
    const limited = sheets.items.slice(0, MAX_SHEETS);
    for (const sheet of limited) sheet.protection.load("protected");
    await ctx.sync();
    return {
      state: sheets.items.length ? "available" : "empty",
      incomplete: sheets.items.length > MAX_SHEETS,
      total: sheets.items.length,
      sheets: limited.map((sheet) => ({
        id: sheet.id,
        name: sheet.name,
        position: sheet.position,
        visibility: sheet.visibility,
        protected: sheet.protection.protected
      })),
      readAt: new Date().toISOString()
    };
  });
}

export async function getSheetOverview(sheetName?: string) {
  const capabilities = officeCapabilities();
  return Excel.run(async (ctx) => {
    const sheet = sheetName?.trim()
      ? ctx.workbook.worksheets.getItem(sheetName.trim())
      : ctx.workbook.worksheets.getActiveWorksheet();
    sheet.load(["id", "name", "visibility"]);
    sheet.protection.load("protected");
    const used = capabilities.usedRangeOrNull ? sheet.getUsedRangeOrNullObject(true) : sheet.getUsedRange(true);
    used.load(["address", "rowCount", "columnCount", "isNullObject"]);
    sheet.tables.load("items/name");
    sheet.charts.load("items/name,chartType");
    sheet.names.load("items/name,type,visible");
    ctx.workbook.names.load("items/name,type,visible,formula");
    if (capabilities.pivotTables) sheet.pivotTables.load("items/name");
    await ctx.sync();

    const tables = sheet.tables.items.slice(0, MAX_OBJECTS);
    const tableRanges = tables.map((table) => ({ table, range: table.getRange() }));
    for (const item of tableRanges) item.range.load("address");
    await ctx.sync();
    const truncated = (count: number) => count > MAX_OBJECTS;
    return {
      state: used.isNullObject ? "empty" : "available",
      sheet: { id: sheet.id, name: sheet.name, visibility: sheet.visibility, protected: sheet.protection.protected },
      usedRange: used.isNullObject ? null : { address: used.address, rowCount: used.rowCount, columnCount: used.columnCount },
      tables: tableRanges.map(({ table, range }) => ({ name: table.name, address: range.address })),
      names: sheet.names.items.slice(0, MAX_OBJECTS).map((item) => ({ name: item.name, type: item.type, visible: item.visible })),
      workbookNames: ctx.workbook.names.items.slice(0, MAX_OBJECTS).map((item) => ({
        name: item.name,
        type: item.type,
        visible: item.visible,
        formula: item.formula
      })),
      charts: sheet.charts.items.slice(0, MAX_OBJECTS).map((chart) => ({ name: chart.name, type: chart.chartType })),
      pivots: capabilities.pivotTables
        ? sheet.pivotTables.items.slice(0, MAX_OBJECTS).map((pivot) => ({ name: pivot.name }))
        : [],
      incomplete: truncated(sheet.tables.items.length) || truncated(sheet.names.items.length) ||
        truncated(ctx.workbook.names.items.length) ||
        truncated(sheet.charts.items.length) || (capabilities.pivotTables && truncated(sheet.pivotTables.items.length)),
      readAt: new Date().toISOString()
    };
  });
}
