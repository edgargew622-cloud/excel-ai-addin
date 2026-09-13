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
import { validateToolArgs, type ToolName } from "./toolSchemas";

export class ToolError extends Error {}

const A1 = /^\$?[A-Za-z]{1,3}\$?\d{1,7}(:\$?[A-Za-z]{1,3}\$?\d{1,7})?$/;
const MAX_IO_CELLS = 20_000;
const MAX_ROWS_PER_STRUCTURAL_OP = 1000;
const MAX_EXACT_FORMAT_UNDO_CELLS = 500;

function checkAddress(address: unknown): string {
  if (typeof address !== "string" || !A1.test(address.trim())) {
    throw new ToolError(
      `Адрес "${String(address)}" не в A1-нотации. Ожидается вид B2:D20, имя листа передаётся отдельно в поле sheet.`
    );
  }
  return address.trim();
}

function sheetOf(ctx: Excel.RequestContext, name?: string) {
  const n = (name ?? "").trim();
  return n ? ctx.workbook.worksheets.getItem(n) : ctx.workbook.worksheets.getActiveWorksheet();
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

async function get_range_values(a: { sheet?: string; address: string }) {
  const address = checkAddress(a.address);
  return Excel.run(async (ctx) => {
    const sheet = sheetOf(ctx, a.sheet);
    const range = sheet.getRange(address);

    // Сначала загружаем только размеры: нельзя сначала выкачать миллион ячеек,
    // а потом обнаружить, что диапазон слишком большой.
    range.load(["address", "rowCount", "columnCount"]);
    sheet.load("name");
    await ctx.sync();

    const cells = range.rowCount * range.columnCount;
    if (cells > MAX_IO_CELLS) {
      throw new ToolError(
        `Диапазон ${address} содержит ${cells} ячеек — слишком много для одного запроса. Читай частями до ${MAX_IO_CELLS} ячеек.`
      );
    }

    range.load(["values", "formulas", "numberFormat"]);
    await ctx.sync();

    return {
      sheet: sheet.name,
      address: range.address,
      rowCount: range.rowCount,
      columnCount: range.columnCount,
      values: range.values,
      formulas: range.formulas,
      numberFormat: range.numberFormat
    };
  });
}

async function set_range_values(a: {
  sheet?: string;
  address: string;
  values: unknown[][];
  isFormula?: boolean;
}) {
  const address = checkAddress(a.address);
  if (!Array.isArray(a.values) || !a.values.length || !Array.isArray(a.values[0])) {
    throw new ToolError("values должен быть непустым двумерным массивом.");
  }

  return Excel.run(async (ctx) => {
    const sheet = sheetOf(ctx, a.sheet);
    const range = sheet.getRange(address);
    range.load(["rowCount", "columnCount", "address"]);
    sheet.load("name");
    await ctx.sync();

    const rows = a.values.length;
    const cols = a.values[0].length;
    if (rows * cols > MAX_IO_CELLS) {
      throw new ToolError(`За одну запись разрешено не более ${MAX_IO_CELLS} ячеек.`);
    }
    if (rows !== range.rowCount || cols !== range.columnCount) {
      throw new ToolError(
        `Размер не совпадает: диапазон ${range.address} это ${range.rowCount}×${range.columnCount}, ` +
          `а values это ${rows}×${cols}.`
      );
    }
    if (a.values.some((r) => !Array.isArray(r) || r.length !== cols)) {
      throw new ToolError("Строки в values разной длины.");
    }

    const undoEnabled = isCustomUndoAvailable();
    const before = undoEnabled ? await captureContent(ctx, sheet.name, address) : null;
    if (a.isFormula) range.formulas = a.values as any[][];
    else range.values = a.values as any[][];
    await ctx.sync();
    let undoRecorded = false;
    if (before) {
      const after = await captureContent(ctx, sheet.name, address);
      undoRecorded = push(guardedContentUndo(a.isFormula ? "запись формул" : "запись значений", before, after));
    }

    return {
      ok: true,
      sheet: sheet.name,
      address: range.address,
      written: `${rows}×${cols}`,
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: "Custom undo недоступен: монитор структурных изменений Excel не активен." })
    };
  });
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

type Handler = (args: any) => Promise<unknown>;

const HANDLERS: Record<ToolName, Handler> = {
  get_range_values,
  set_range_values,
  insert_rows,
  delete_rows,
  sort_range,
  apply_filter,
  create_pivot_table,
  create_chart,
  format_range
};

export async function runTool(name: string, args: unknown): Promise<unknown> {
  const handler = HANDLERS[name as ToolName];
  if (!handler) {
    throw new ToolError(`Инструмента "${name}" не существует. Доступны: ${Object.keys(HANDLERS).join(", ")}.`);
  }
  const validation = validateToolArgs(name, args);
  if (!validation.ok) throw new ToolError(validation.error);
  return handler(args);
}
