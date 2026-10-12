/**
 * Объединить таблицы с одинаковой шапкой в одну (12.10.2026): «Данные» за
 * 2021–2022 и скрытый лист «Данные 2023» — в одну базу для сводных и
 * дашборда. Замена «Добавить запросы» из Power Query, которого надстройкам нет.
 *
 * Значения читает и пишет панель, не модель: строк может быть десятки тысяч.
 * Столбцы сопоставляются по заголовкам (порядок может отличаться). Результат —
 * новый лист значениями, по желанию столбец «Источник» с именем листа.
 * Сверка: число строк и сумма каждого числового столбца равны сумме по источникам.
 */

import { assertPlanWorkbook, checkAddress, deepFreeze, fitNewTable, preflightToolArgs, ToolError, ToolExecutionError, valuesForLiteralWrite } from "./excelTools";
import { checkSheetName, freeSheetName } from "./sheetRules";
import { action, getStructuralRevision, isCustomUndoAvailable, push } from "./undo";
import { captureTarget, type WorkbookTarget } from "./workbookContext";
import { columnLetters } from "./formulaFill";

export const MAX_APPEND_CELLS = 500_000;
const WRITE_CELLS_PER_SYNC = 5_000;
const norm = (value: unknown) => String(value ?? "").trim().toLowerCase();
const isBlank = (value: unknown) => value === "" || value === null || value === undefined;

export interface AppendSource { sheet: string; header: readonly unknown[]; rows: readonly (readonly unknown[])[]; formats: readonly (readonly unknown[])[] }

/** Склейка без Excel: столбцы — по шапке первого источника; пустые строки пропускаются. */
export function appendRows(sources: readonly AppendSource[], sourceColumn: string | null) {
  const header = sources[0].header.map((value) => String(value ?? "").trim());
  const problems: string[] = [];
  const out: unknown[][] = [];
  const formats: unknown[][] = [];
  const counts: number[] = [];
  for (const source of sources) {
    const map = header.map((name) => source.header.findIndex((value) => norm(value) === norm(name)));
    const missing = header.filter((_, i) => map[i] < 0);
    const extra = source.header.filter((value) => !isBlank(value) && !header.some((name) => norm(name) === norm(value)));
    if (missing.length || extra.length) {
      problems.push(`«${source.sheet}»: ${missing.length ? `нет столбцов ${missing.map((name) => `«${name}»`).join(", ")}` : ""}${missing.length && extra.length ? "; " : ""}${extra.length ? `лишние столбцы ${extra.map((name) => `«${name}»`).join(", ")}` : ""}`);
      continue;
    }
    let count = 0;
    source.rows.forEach((row, r) => {
      if (row.every(isBlank)) return;
      out.push([...map.map((index) => row[index] ?? ""), ...(sourceColumn ? [source.sheet] : [])]);
      formats.push([...map.map((index) => source.formats[r]?.[index] ?? "General"), ...(sourceColumn ? ["General"] : [])]);
      count++;
    });
    counts.push(count);
  }
  const sums = header.map((_, c) => out.reduce((total: number, row) => total + (typeof row[c] === "number" ? row[c] as number : 0), 0));
  return { header: [...header, ...(sourceColumn ? [sourceColumn] : [])], rows: out, formats, counts, sums, problems };
}

export interface AppendPlan {
  readonly kind: "append_ranges";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly destSheet: string;
  readonly sources: readonly { sheet: string; address: string; rows: number }[];
  readonly header: readonly unknown[];
  readonly rows: readonly (readonly unknown[])[];
  readonly formats: readonly (readonly unknown[])[];
  readonly sums: readonly number[];
  readonly preview: readonly string[];
  readonly undoAvailable: boolean;
  readonly createdAt: string;
}

export async function prepareAppendPlan(args: unknown): Promise<AppendPlan> {
  preflightToolArgs("append_ranges", args);
  const a = args as { sources: { sheet: string; address?: string }[]; newSheet?: string; sourceColumn?: string | boolean };
  if (!Array.isArray(a.sources) || a.sources.length < 2) throw new ToolError("Нужно хотя бы две таблицы (sources): лист и, если нужно, область с шапкой.");
  const target = await captureTarget(undefined);
  const prepared = await Excel.run(async (ctx) => {
    const all = ctx.workbook.worksheets;
    all.load("items/name,items/visibility");
    await ctx.sync();
    const sources: AppendSource[] = [];
    const described: { sheet: string; address: string; rows: number }[] = [];
    let cells = 0;
    for (const item of a.sources) {
      const sheet = all.items.find((w) => norm(w.name) === norm(item.sheet));
      if (!sheet) throw new ToolError(`Листа «${item.sheet}» нет. Есть: ${all.items.map((w) => `«${w.name}»`).join(", ")}.`);
      const range = item.address?.trim() ? sheet.getRange(checkAddress(item.address)) : sheet.getUsedRange(true);
      range.load(["address", "rowCount", "columnCount"]);
      await ctx.sync();
      cells += range.rowCount * range.columnCount;
      if (cells > MAX_APPEND_CELLS) throw new ToolError(`Вместе больше ${MAX_APPEND_CELLS.toLocaleString("ru-RU")} ячеек — за раз столько не объединить.`);
      range.load(["values", "numberFormat"]);
      await ctx.sync();
      const values = range.values as unknown[][];
      const formats = range.numberFormat as unknown[][];
      if (values.length < 2) throw new ToolError(`На «${sheet.name}» (${range.address}) нет строк под шапкой.`);
      sources.push({ sheet: sheet.name, header: values[0], rows: values.slice(1), formats: formats.slice(1) });
      described.push({ sheet: sheet.name + (sheet.visibility !== "Visible" ? " (скрытый)" : ""), address: String(range.address).replace(/^.*!/, ""), rows: values.length - 1 });
    }
    const sourceColumn = a.sourceColumn === false ? null : typeof a.sourceColumn === "string" && a.sourceColumn.trim() ? a.sourceColumn.trim() : "Источник";
    const joined = appendRows(sources, sourceColumn);
    if (joined.problems.length) throw new ToolError(`Шапки не совпадают — объединять нельзя: ${joined.problems.join("; ")}. Операция не выполнялась.`);
    const wanted = a.newSheet?.trim() || "Все данные";
    let destSheet: string;
    try {
      destSheet = checkSheetName(wanted, all.items.map((w) => w.name));
    } catch (error: any) {
      const message = String(error?.message ?? error);
      throw new ToolError(/уже есть/.test(message) ? `${message} Свободно, например, «${freeSheetName(wanted, all.items.map((w) => w.name))}».` : message);
    }
    return {
      kind: "append_ranges" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      target,
      destSheet,
      sources: described.map((item, i) => ({ ...item, rows: joined.counts[i] })),
      header: joined.header,
      rows: joined.rows,
      formats: joined.formats,
      sums: joined.sums,
      preview: [
        ...described.map((item, i) => `«${item.sheet}» ${item.address}: ${joined.counts[i]} строк`),
        `Итого ${joined.rows.length} строк на новом листе «${destSheet}»${sourceColumn ? `, столбец «${sourceColumn}» — откуда строка` : ""}. Источники не меняются.`,
        `Столбцы: ${joined.header.map((name) => `«${name}»`).join(", ")}`
      ],
      undoAvailable: isCustomUndoAvailable(),
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executeAppendPlan(plan: AppendPlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const all = ctx.workbook.worksheets;
    all.load("items/name");
    await ctx.sync();
    if (all.items.some((w) => norm(w.name) === norm(plan.destSheet))) {
      throw new ToolExecutionError(`Лист «${plan.destSheet}» появился после предпросмотра. Ничего не записано.`, "failed_before_write");
    }
    let sheet: Excel.Worksheet;
    try {
      sheet = ctx.workbook.worksheets.add(plan.destSheet);
      sheet.load(["id", "name"]);
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(`Excel отказал в создании листа «${plan.destSheet}»: ${error?.message ?? error}.`, "unknown");
    }
    const width = plan.header.length;
    const lines = [plan.header, ...plan.rows];
    const perSync = Math.max(1, Math.floor(WRITE_CELLS_PER_SYNC / width));
    try {
      for (let start = 0; start < lines.length; start += perSync) {
        const part = lines.slice(start, start + perSync);
        const chunk = sheet.getRangeByIndexes(start, 0, part.length, width);
        chunk.values = valuesForLiteralWrite(part as unknown[][]) as any[][];
        // Форматы — построчно из источников: даты остаются датами.
        if (start > 0 || part.length > 1) {
          const body = start === 0 ? part.slice(1) : part;
          const formats = plan.formats.slice(Math.max(0, start - 1), Math.max(0, start - 1) + body.length);
          if (body.length) sheet.getRangeByIndexes(start === 0 ? 1 : start, 0, body.length, width).numberFormat = formats.map((row) => row.map((f) => String(f || "General"))) as any;
        }
        await ctx.sync();
      }
      sheet.getRangeByIndexes(0, 0, 1, width).format.font.bold = true;
      sheet.freezePanes.freezeRows(1);
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(`Запись на лист «${plan.destSheet}» прервалась: ${error?.message ?? error}. Часть строк могла записаться.`, "unknown");
    }
    await fitNewTable(ctx, sheet, 0, width);
    const address = `A1:${columnLetters(width)}${lines.length}`;
    const written = sheet.getRange(address);
    written.load(["values", "formulas"]);
    await ctx.sync();
    const back = (written.values as unknown[][]).slice(1);
    const problems: string[] = [];
    if (back.length !== plan.rows.length) problems.push(`строк ${back.length} вместо ${plan.rows.length}`);
    plan.sums.forEach((expected, c) => {
      const got = back.reduce((total: number, row) => total + (typeof row[c] === "number" ? row[c] as number : 0), 0);
      if (Math.abs(got - expected) > 1e-6 * Math.max(1, Math.abs(expected))) problems.push(`сумма «${plan.header[c]}» ${got} вместо ${expected}`);
    });
    let undoRecorded = false;
    if (plan.undoAvailable) {
      const signature = JSON.stringify(written.formulas);
      const sheetId = sheet.id;
      undoRecorded = push(action(`объединение на листе ${plan.destSheet}`, async () => {
        const revision = getStructuralRevision();
        await Excel.run(async (undoCtx) => {
          const undoSheet = undoCtx.workbook.worksheets.getItem(sheetId);
          const range = undoSheet.getRange(address);
          range.load("formulas");
          await undoCtx.sync();
          if (JSON.stringify(range.formulas) !== signature) throw new Error(`Лист «${plan.destSheet}» изменили после объединения. Отмена остановлена.`);
          if (getStructuralRevision() !== revision) throw new Error("Структура книги изменилась во время отмены. Отмена остановлена.");
          undoSheet.delete();
          await undoCtx.sync();
        });
      }));
    }
    if (problems.length) throw new ToolExecutionError(`Данные записаны на «${plan.destSheet}», но сверка расходится: ${problems.join("; ")}.`, "applied");
    return {
      ok: true,
      executionState: "verified",
      sheet: plan.destSheet,
      address,
      rows: plan.rows.length,
      tableAddress: address,
      tableNote: `Для таблицы Excel и сводных бери ровно ${address} — шапка и все ${plan.rows.length} строк; не пересчитывай адрес сам.`,
      sources: plan.sources,
      note: "Сверено: число строк и суммы числовых столбцов равны сумме по источникам. Дальше — create_table по этому листу и сводные по таблице.",
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: "Автоматическая отмена недоступна: лист удаляется вручную." })
    };
  });
}
