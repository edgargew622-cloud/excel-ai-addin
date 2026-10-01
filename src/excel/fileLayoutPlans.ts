/**
 * Перенос PDF «как есть» (01.10.2026): таблица или бланк из PDF становится
 * новым листом с той же сеткой — столбцы и строки тех же размеров,
 * объединённые ячейки, рамки, жирный текст, выравнивание.
 *
 * Разметку строит сервер (server/src/files/pdfLayout.ts) прямо из файла, модель
 * её не пересказывает и исказить не может. Лист всегда новый: перенос ничего
 * не затирает, а отмена просто удаляет его, если на нём ничего не меняли.
 * Значения сверяются с разметкой после записи.
 */

import {
  assertPlanWorkbook,
  deepFreeze,
  preflightToolArgs,
  ToolError,
  ToolExecutionError,
  valuesForLiteralWrite
} from "./excelTools";
import { contains, parseA1Rect } from "./a1";
import { columnLetters } from "./formulaFill";
import { checkSheetName, freeSheetName } from "./sheetRules";
import { action, getStructuralRevision, isCustomUndoAvailable, push } from "./undo";
import { captureTarget, type WorkbookTarget } from "./workbookContext";
import { fetchFileLayout, type FileLayout } from "../taskpane/api/files";

const OPS_PER_SYNC = 1_500;
const TOLERANCE = 1e-9;

export interface FileLayoutPlan {
  readonly kind: "import_file_layout";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly fileId: string;
  readonly fileName: string;
  readonly sheetName: string;
  readonly layout: FileLayout;
  readonly rows: number;
  readonly columns: number;
  readonly merges: number;
  readonly borderedEdges: number;
  readonly preview: readonly string[];
  readonly undoAvailable: boolean;
  readonly undoNote?: string;
  readonly createdAt: string;
}

/** Что и куда писать: значения, форматы, объединения, оформление, рамки. Чистая функция — её проверяют тесты. */
export function layoutWrite(layout: FileLayout) {
  const rows = layout.rowHeights.length;
  const columns = layout.columnWidths.length;
  const values: (string | number)[][] = Array.from({ length: rows }, () => Array(columns).fill(""));
  const formats: string[][] = Array.from({ length: rows }, () => Array(columns).fill("General"));
  const merges: string[] = [];
  const address = (r: number, c: number, rowSpan = 1, colSpan = 1) =>
    rowSpan === 1 && colSpan === 1
      ? `${columnLetters(c + 1)}${r + 1}`
      : `${columnLetters(c + 1)}${r + 1}:${columnLetters(c + colSpan)}${r + rowSpan}`;
  for (const cell of layout.cells) {
    if (cell.r >= rows || cell.c >= columns) continue;
    values[cell.r][cell.c] = cell.value;
    formats[cell.r][cell.c] = cell.numberFormat;
    if (cell.rowSpan > 1 || cell.colSpan > 1) merges.push(address(cell.r, cell.c, cell.rowSpan, Math.min(cell.colSpan, columns - cell.c)));
  }
  const borders: Array<{ cell: string; edge: "EdgeTop" | "EdgeBottom" | "EdgeLeft" | "EdgeRight"; weight: "Thin" | "Medium" }> = [];
  for (const [r, c, weight] of layout.hEdges) {
    if (c >= columns) continue;
    if (r < rows) borders.push({ cell: address(r, c), edge: "EdgeTop", weight });
    else if (r === rows) borders.push({ cell: address(r - 1, c), edge: "EdgeBottom", weight });
  }
  for (const [r, c, weight] of layout.vEdges) {
    if (r >= rows) continue;
    if (c < columns) borders.push({ cell: address(r, c), edge: "EdgeLeft", weight });
    else if (c === columns) borders.push({ cell: address(r, c - 1), edge: "EdgeRight", weight });
  }
  const block = `A1:${columnLetters(columns)}${rows}`;
  return { rows, columns, values, formats, merges, borders, block, address };
}

export async function prepareFileLayoutPlan(args: unknown): Promise<FileLayoutPlan> {
  preflightToolArgs("import_file_layout", args);
  const a = args as { fileId: string; newSheet?: string };
  const { file, layout } = await fetchFileLayout(a.fileId);
  const write = layoutWrite(layout);
  if (!write.rows || !write.columns) throw new ToolError(`В «${file}» не нашлось ничего, что можно разложить по ячейкам.`);
  const target = await captureTarget();
  const prepared = await Excel.run(async (ctx) => {
    const all = ctx.workbook.worksheets;
    all.load("items/name");
    await ctx.sync();
    const existing = all.items.map((item) => item.name);
    const wanted = a.newSheet?.trim() || `Копия ${file.replace(/\.pdf$/i, "")}`;
    let sheetName: string;
    try {
      sheetName = checkSheetName(wanted.replace(/[\\/?*[\]:]/g, " ").slice(0, 31).trim(), existing);
    } catch (error: any) {
      if (a.newSheet?.trim()) {
        const message = String(error?.message ?? error);
        throw new ToolError(/уже есть/.test(message) ? `${message} Свободно, например, «${freeSheetName(wanted, existing)}».` : message);
      }
      sheetName = freeSheetName(wanted.replace(/[\\/?*[\]:]/g, " ").slice(0, 28).trim() || "Копия PDF", existing);
    }
    const lines = [...layout.cells].sort((p, q) => p.r - q.r || p.c - q.c).filter((cell) => cell.text);
    const preview: string[] = [];
    for (let r = 0, i = 0; preview.length < 6 && i < lines.length; r++) {
      const row = lines.filter((cell) => cell.r === r).map((cell) => cell.text.replace(/\n/g, " "));
      if (row.length) preview.push(row.join(" | ").slice(0, 160));
      i += row.length;
      if (r > write.rows) break;
    }
    const undo = isCustomUndoAvailable();
    return {
      kind: "import_file_layout" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      target,
      fileId: a.fileId,
      fileName: file,
      sheetName,
      layout,
      rows: write.rows,
      columns: write.columns,
      merges: write.merges.length,
      borderedEdges: write.borders.length,
      preview,
      undoAvailable: undo,
      ...(undo ? {} : { undoNote: "Отмена недоступна: монитор изменений Excel не активен." }),
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

function sameValue(actual: unknown, expected: string | number): boolean {
  if (typeof expected === "number") return typeof actual === "number" && Math.abs(actual - expected) <= TOLERANCE * Math.max(1, Math.abs(expected));
  return String(actual ?? "") === expected;
}

export async function executeFileLayoutPlan(plan: FileLayoutPlan) {
  assertPlanWorkbook(plan);
  const write = layoutWrite(plan.layout);
  return Excel.run(async (ctx) => {
    const all = ctx.workbook.worksheets;
    all.load("items/name");
    await ctx.sync();
    if (all.items.some((item) => item.name.trim().toLowerCase() === plan.sheetName.toLowerCase())) {
      throw new ToolExecutionError(`Лист «${plan.sheetName}» появился после предпросмотра. Перенос не выполнялся — выберите другое имя.`, "failed_before_write");
    }
    let sheet: Excel.Worksheet;
    try {
      sheet = ctx.workbook.worksheets.add(plan.sheetName);
      sheet.load(["id", "name"]);
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(`Excel отказал в создании листа «${plan.sheetName}»: ${error?.message ?? error}.`, "unknown");
    }
    const sheetId = sheet.id;
    let ops = 0;
    const tick = async () => { if (++ops % OPS_PER_SYNC === 0) await ctx.sync(); };
    const problems: string[] = [];
    try {
      // Как бланк: без служебной сетки Excel.
      try { sheet.showGridlines = false; } catch { /* старый Excel */ }
      const block = sheet.getRange(write.block);
      block.format.font.name = "Arial";
      block.format.verticalAlignment = "Center" as any;
      plan.layout.columnWidths.forEach((width, c) => { sheet.getRangeByIndexes(0, c, 1, 1).format.columnWidth = width; });
      plan.layout.rowHeights.forEach((height, r) => { sheet.getRangeByIndexes(r, 0, 1, 1).format.rowHeight = height; });
      await ctx.sync();
      // Сначала формат («@» — текст), потом значения: код «007» и номер «0042»
      // не должны стать числами, а дата из PDF — текстом.
      const rowsPerSync = Math.max(1, Math.floor(5_000 / write.columns));
      for (let start = 0; start < write.rows; start += rowsPerSync) {
        const count = Math.min(rowsPerSync, write.rows - start);
        const chunk = sheet.getRangeByIndexes(start, 0, count, write.columns);
        chunk.numberFormat = write.formats.slice(start, start + count) as any[][];
        chunk.values = valuesForLiteralWrite(write.values.slice(start, start + count)) as any[][];
        await ctx.sync();
      }
      for (const area of write.merges) { sheet.getRange(area).merge(false); await tick(); }
      await ctx.sync();
      for (const cell of plan.layout.cells) {
        if (cell.r >= write.rows || cell.c >= write.columns) continue;
        const range = sheet.getRange(write.address(cell.r, cell.c, cell.rowSpan, Math.min(cell.colSpan, write.columns - cell.c)));
        if (cell.bold) range.format.font.bold = true;
        range.format.font.size = cell.size;
        range.format.horizontalAlignment = cell.align as any;
        if (cell.wrap) range.format.wrapText = true;
        await tick();
      }
      for (const border of write.borders) {
        const edge = sheet.getRange(border.cell).format.borders.getItem(border.edge as any);
        edge.style = "Continuous" as any;
        edge.weight = border.weight as any;
        await tick();
      }
      await ctx.sync();
    } catch (error: any) {
      problems.push(`оформление прервалось: ${error?.message ?? error}`);
    }

    // Сверка: каждое значение — с разметкой; объединения — по числу областей.
    const written = sheet.getRange(write.block);
    written.load(["values", "formulas"]);
    await ctx.sync();
    const values = written.values as unknown[][];
    let mismatched = 0;
    plan.layout.cells.forEach((cell) => {
      if (cell.r >= write.rows || cell.c >= write.columns) return;
      if (!sameValue(values[cell.r]?.[cell.c], cell.value)) {
        mismatched += 1;
        if (mismatched <= 6) problems.push(`${write.address(cell.r, cell.c)}: ${JSON.stringify(values[cell.r]?.[cell.c])} вместо ${JSON.stringify(cell.value)}`);
      }
    });
    let mergedFound: number | null = null;
    try {
      const merged = written.getMergedAreasOrNullObject();
      merged.load(["isNullObject", "areaCount"]);
      await ctx.sync();
      mergedFound = merged.isNullObject ? 0 : merged.areaCount;
      if (mergedFound !== write.merges.length) problems.push(`объединений ${mergedFound} вместо ${write.merges.length}`);
    } catch { /* ExcelApi без getMergedAreas — объединения не сверить */ }
    try { sheet.activate(); sheet.getRange("A1").select(); await ctx.sync(); } catch { /* не страшно */ }

    let undoRecorded = false;
    if (plan.undoAvailable) {
      const signature = JSON.stringify(written.formulas);
      undoRecorded = push(action(`копия «${plan.fileName}» на листе ${plan.sheetName}`, async () => {
        const revision = getStructuralRevision();
        await Excel.run(async (undoCtx) => {
          const undoSheet = undoCtx.workbook.worksheets.getItemOrNullObject(sheetId);
          undoSheet.load("isNullObject");
          await undoCtx.sync();
          if (undoSheet.isNullObject) throw new Error(`Листа «${plan.sheetName}» уже нет. Отменять нечего.`);
          const range = undoSheet.getRange(write.block);
          range.load("formulas");
          const used = undoSheet.getUsedRangeOrNullObject(true);
          used.load(["isNullObject", "address"]);
          await undoCtx.sync();
          const usedRect = used.isNullObject ? null : parseA1Rect(String(used.address).replace(/^.*!/, ""));
          const blockRect = parseA1Rect(write.block)!;
          if (JSON.stringify(range.formulas) !== signature || (usedRect && !contains(blockRect, usedRect))) {
            throw new Error(`На листе «${plan.sheetName}» после переноса что-то изменили. Отмена остановлена, чтобы не потерять правку; удалите лист сами, если он не нужен.`);
          }
          if (getStructuralRevision() !== revision) throw new Error("Структура книги изменилась во время отмены. Отмена остановлена.");
          undoSheet.delete();
          await undoCtx.sync();
        });
      }));
    }

    const result = {
      sheet: plan.sheetName,
      file: plan.fileName,
      pages: plan.layout.pages,
      size: `${write.rows} строк × ${write.columns} столбцов`,
      merges: write.merges.length,
      borderedEdges: write.borders.length,
      cells: plan.layout.cells.filter((cell) => cell.text).length,
      ...(plan.layout.warnings.length ? { warnings: plan.layout.warnings } : {}),
      note: "Сетка, объединения, рамки, жирный текст и размеры взяты из PDF; значения сверены с файлом. Не переносятся картинки, печати, подписи как изображения, цвета текста и заливки — назови это пользователю, если в файле они были.",
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: plan.undoNote ?? "Автоматическая отмена этой операции недоступна." })
    };
    if (problems.length) {
      throw new ToolExecutionError(
        `Копия «${plan.fileName}» записана на лист «${plan.sheetName}», но расходится с файлом: ${problems.slice(0, 8).join("; ")}. ` +
        (undoRecorded ? "Лист можно убрать кнопкой «Отменить»." : "Проверьте лист."),
        "applied"
      );
    }
    return { ok: true, executionState: "verified", ...result };
  });
}
