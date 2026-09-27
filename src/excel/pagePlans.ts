/**
 * Параметры страницы для печати (этап 8, 8.3.3): ориентация, поля, масштаб
 * «вписать в N страниц», область печати, повтор строк заголовков, колонтитулы.
 *
 * Замер 28 сентября 2026 года (Office 2021): всё пишется и читается обратно
 * через `worksheet.pageLayout`. Поля Excel хранит в пунктах; «вписать» и
 * масштаб взаимоисключающие — при вписывании scale читается как null, а 0
 * страниц («авто») — тоже как null. Пустая строка снимает область печати и
 * повтор заголовков. Меняется только то, что попросили; прежние значения
 * запоминаются, по ним работает отмена.
 */

import { parseA1Rect } from "./a1";
import {
  assertPlanWorkbook,
  checkAddress,
  deepFreeze,
  preflightToolArgs,
  ToolError,
  ToolExecutionError
} from "./excelTools";
import { action, getStructuralRevision, isCustomUndoAvailable, push } from "./undo";
import { captureTarget, type WorkbookTarget } from "./workbookContext";

export const POINTS_PER_CM = 72 / 2.54;

type Side = "left" | "right" | "top" | "bottom";
type Slot = "left" | "center" | "right";
const SIDES: readonly Side[] = ["left", "right", "top", "bottom"];
const SLOTS: readonly Slot[] = ["left", "center", "right"];

export interface PageState {
  orientation: "Portrait" | "Landscape";
  /** Поля в пунктах, как их хранит Excel. */
  margins: Record<Side, number>;
  /** scale — масштаб в процентах; wide/tall — вписать в N страниц (null — авто). */
  zoom: { scale: number | null; wide: number | null; tall: number | null };
  /** Область печати без имени листа, например A1:F40; null — вся занятая область. */
  printArea: string | null;
  /** Повторяемые строки, например 1:1; null — нет. */
  titleRows: string | null;
  header: Record<Slot, string>;
  footer: Record<Slot, string>;
}

/** Что попросили поменять — только эти поля и сверяются. */
export interface PageRequest {
  orientation?: "Portrait" | "Landscape";
  margins?: Partial<Record<Side, number>>;
  zoom?: { scale: number } | { wide: number | null; tall: number | null };
  printArea?: string | null;
  titleRows?: string | null;
  header?: Partial<Record<Slot, string>>;
  footer?: Partial<Record<Slot, string>>;
}

export interface PageLayoutPlan {
  readonly kind: "set_page_layout";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly request: PageRequest;
  readonly before: PageState;
  readonly expected: PageState;
  readonly preview: readonly string[];
  readonly undoAvailable: boolean;
  readonly undoNote?: string;
  readonly createdAt: string;
}

const bare = (address: string | null | undefined) =>
  address ? address.slice(address.lastIndexOf("!") + 1).replace(/\$/g, "").toUpperCase() : null;

const cm = (points: number) => (points / POINTS_PER_CM).toFixed(2).replace(".", ",");

function describeZoom(zoom: PageState["zoom"]): string {
  if (zoom.wide !== null || zoom.tall !== null) {
    const wide = zoom.wide === null ? "авто" : String(zoom.wide);
    const tall = zoom.tall === null ? "авто" : String(zoom.tall);
    return `вписать: ${wide} стр. в ширину × ${tall} в высоту`;
  }
  return `${zoom.scale ?? 100}%`;
}

/** Строки «1», «1:2», «$1:$1» → «1:1», «1:2». */
function parseTitleRows(text: string): string {
  const match = /^\$?(\d+)(?::\$?(\d+))?$/.exec(text.trim());
  if (!match) throw new ToolError(`printTitleRows — номера строк, например «1» или «1:2»; получено «${text}».`);
  const first = Number(match[1]);
  const last = Number(match[2] ?? match[1]);
  if (first < 1 || last < first || last > 1_048_576) throw new ToolError(`Строки заголовков «${text}» не годятся: нужны номера по возрастанию, от 1.`);
  return `${first}:${last}`;
}

export function parsePageRequest(args: Record<string, unknown>): PageRequest {
  const request: PageRequest = {};
  if (args.orientation !== undefined) request.orientation = args.orientation === "landscape" ? "Landscape" : "Portrait";
  if (args.marginsCm !== undefined) {
    const raw = args.marginsCm as Record<string, unknown>;
    const margins: Partial<Record<Side, number>> = {};
    for (const side of SIDES) {
      if (raw?.[side] === undefined) continue;
      const value = Number(raw[side]);
      if (!Number.isFinite(value) || value < 0 || value > 20) throw new ToolError(`Поле ${side} — от 0 до 20 см; получено ${String(raw[side])}.`);
      margins[side] = Math.round(value * POINTS_PER_CM * 100) / 100;
    }
    if (!Object.keys(margins).length) throw new ToolError("marginsCm — хотя бы одно из left, right, top, bottom.");
    request.margins = margins;
  }
  const wantsFit = args.fitToPagesWide !== undefined || args.fitToPagesTall !== undefined;
  if (wantsFit && args.scale !== undefined) throw new ToolError("scale не сочетается с fitToPagesWide/fitToPagesTall: Excel либо масштабирует, либо вписывает.");
  if (args.scale !== undefined) {
    const scale = Number(args.scale);
    if (!Number.isInteger(scale) || scale < 10 || scale > 400) throw new ToolError("scale — целое от 10 до 400 процентов.");
    request.zoom = { scale };
  } else if (wantsFit) {
    const pages = (value: unknown) => {
      if (value === undefined || value === 0) return null;
      const number = Number(value);
      if (!Number.isInteger(number) || number < 0 || number > 100) throw new ToolError("Число страниц для вписывания — целое от 1 до 100; 0 — авто.");
      return number;
    };
    const wide = pages(args.fitToPagesWide);
    const tall = pages(args.fitToPagesTall);
    if (wide === null && tall === null) throw new ToolError("Для вписывания нужна хотя бы одна сторона: fitToPagesWide или fitToPagesTall.");
    request.zoom = { wide, tall };
  }
  if (args.printArea !== undefined) {
    const text = String(args.printArea).trim();
    if (!text || text.toLowerCase() === "none") request.printArea = null;
    else {
      const address = checkAddress(text);
      const rect = parseA1Rect(address);
      if (!rect || rect.kind !== "cells") throw new ToolError(`Область печати — прямоугольник ячеек, например A1:F40; получено «${text}».`);
      request.printArea = address.toUpperCase().replace(/\$/g, "");
    }
  }
  if (args.printTitleRows !== undefined) {
    const text = String(args.printTitleRows).trim();
    request.titleRows = !text || text.toLowerCase() === "none" ? null : parseTitleRows(text);
  }
  for (const part of ["header", "footer"] as const) {
    if (args[part] === undefined) continue;
    const raw = args[part] as Record<string, unknown>;
    const texts: Partial<Record<Slot, string>> = {};
    for (const slot of SLOTS) {
      if (raw?.[slot] === undefined) continue;
      const text = String(raw[slot]);
      if (text.length > 255) throw new ToolError(`Текст колонтитула длиннее 255 знаков: Excel его не примет.`);
      texts[slot] = text;
    }
    if (!Object.keys(texts).length) throw new ToolError(`${part} — хотя бы одно из left, center, right.`);
    request[part] = texts;
  }
  if (!Object.keys(request).length) throw new ToolError("Не указано, что менять: ориентацию, поля, масштаб, область печати, строки заголовков или колонтитулы.");
  return request;
}

export function applyPageRequest(before: PageState, request: PageRequest): PageState {
  return {
    orientation: request.orientation ?? before.orientation,
    margins: { ...before.margins, ...request.margins },
    zoom: request.zoom
      ? "scale" in request.zoom ? { scale: request.zoom.scale, wide: null, tall: null } : { scale: null, wide: request.zoom.wide, tall: request.zoom.tall }
      : before.zoom,
    printArea: request.printArea !== undefined ? request.printArea : before.printArea,
    titleRows: request.titleRows !== undefined ? request.titleRows : before.titleRows,
    header: { ...before.header, ...request.header },
    footer: { ...before.footer, ...request.footer }
  };
}

/** Строки предпросмотра: было → станет, только по запрошенному. */
export function describePageChange(before: PageState, after: PageState, request: PageRequest): string[] {
  const lines: string[] = [];
  const orientation = (value: string) => (value === "Landscape" ? "альбомная" : "книжная");
  if (request.orientation) lines.push(`Ориентация: ${orientation(before.orientation)} → ${orientation(after.orientation)}`);
  if (request.margins) {
    const names: Record<Side, string> = { left: "слева", right: "справа", top: "сверху", bottom: "снизу" };
    lines.push(`Поля, см: ${SIDES.filter((side) => request.margins![side] !== undefined).map((side) => `${names[side]} ${cm(before.margins[side])} → ${cm(after.margins[side])}`).join("; ")}`);
  }
  if (request.zoom) lines.push(`Масштаб: ${describeZoom(before.zoom)} → ${describeZoom(after.zoom)}`);
  if (request.printArea !== undefined) lines.push(`Область печати: ${before.printArea ?? "весь лист"} → ${after.printArea ?? "весь лист"}`);
  if (request.titleRows !== undefined) lines.push(`Повтор строк на каждой странице: ${before.titleRows ?? "нет"} → ${after.titleRows ?? "нет"}`);
  const slotNames: Record<Slot, string> = { left: "слева", center: "по центру", right: "справа" };
  for (const part of ["header", "footer"] as const) {
    const texts = request[part];
    if (!texts) continue;
    for (const slot of SLOTS) {
      if (texts[slot] === undefined) continue;
      lines.push(`${part === "header" ? "Верхний" : "Нижний"} колонтитул ${slotNames[slot]}: «${before[part][slot]}» → «${after[part][slot]}»`);
    }
  }
  if ([request.header, request.footer].some((texts) => texts && Object.values(texts).some((text) => /&[PNDTFA]/.test(text ?? "")))) {
    lines.push("Коды колонтитула: &P — номер страницы, &N — всего страниц, &D — дата, &T — время, &F — файл, &A — лист.");
  }
  return lines;
}

/** Чем прочитанное расходится с ожидаемым — только по запрошенным полям. */
export function pageMismatches(actual: PageState, expected: PageState, request: PageRequest): string[] {
  const problems: string[] = [];
  if (request.orientation && actual.orientation !== expected.orientation) problems.push(`ориентация ${actual.orientation} вместо ${expected.orientation}`);
  for (const side of SIDES) {
    if (request.margins?.[side] === undefined) continue;
    if (Math.abs(actual.margins[side] - expected.margins[side]) > 0.1) problems.push(`поле ${side} ${cm(actual.margins[side])} см вместо ${cm(expected.margins[side])}`);
  }
  if (request.zoom && describeZoom(actual.zoom) !== describeZoom(expected.zoom)) problems.push(`масштаб «${describeZoom(actual.zoom)}» вместо «${describeZoom(expected.zoom)}»`);
  if (request.printArea !== undefined && actual.printArea !== expected.printArea) problems.push(`область печати ${actual.printArea ?? "нет"} вместо ${expected.printArea ?? "нет"}`);
  if (request.titleRows !== undefined && actual.titleRows !== expected.titleRows) problems.push(`строки заголовков ${actual.titleRows ?? "нет"} вместо ${expected.titleRows ?? "нет"}`);
  for (const part of ["header", "footer"] as const) {
    for (const slot of SLOTS) {
      if (request[part]?.[slot] === undefined) continue;
      if (actual[part][slot] !== expected[part][slot]) problems.push(`${part === "header" ? "верхний" : "нижний"} колонтитул (${slot}) «${actual[part][slot]}» вместо «${expected[part][slot]}»`);
    }
  }
  return problems;
}

export async function readPageState(ctx: Excel.RequestContext, sheet: Excel.Worksheet): Promise<PageState> {
  const layout = sheet.pageLayout;
  layout.load(["orientation", "leftMargin", "rightMargin", "topMargin", "bottomMargin", "zoom"]);
  const area = layout.getPrintAreaOrNullObject();
  area.load(["isNullObject", "address"]);
  const titles = layout.getPrintTitleRowsOrNullObject();
  titles.load(["isNullObject", "address"]);
  const texts = layout.headersFooters.defaultForAllPages;
  texts.load(["leftHeader", "centerHeader", "rightHeader", "leftFooter", "centerFooter", "rightFooter"]);
  await ctx.sync();
  const zoom = (layout.zoom ?? {}) as { scale?: number | null; horizontalFitToPages?: number | null; verticalFitToPages?: number | null };
  const pages = (value: number | null | undefined) => (value ? value : null);
  return {
    orientation: String(layout.orientation) === "Landscape" ? "Landscape" : "Portrait",
    margins: { left: layout.leftMargin, right: layout.rightMargin, top: layout.topMargin, bottom: layout.bottomMargin },
    zoom: { scale: zoom.scale ?? null, wide: pages(zoom.horizontalFitToPages), tall: pages(zoom.verticalFitToPages) },
    printArea: area.isNullObject ? null : bare(area.address),
    titleRows: titles.isNullObject ? null : bare(titles.address),
    header: { left: texts.leftHeader ?? "", center: texts.centerHeader ?? "", right: texts.rightHeader ?? "" },
    footer: { left: texts.leftFooter ?? "", center: texts.centerFooter ?? "", right: texts.rightFooter ?? "" }
  };
}

/** Записывает в Excel поля state, названные в request. */
function writePage(sheet: Excel.Worksheet, state: PageState, request: PageRequest) {
  const layout = sheet.pageLayout;
  if (request.orientation) layout.orientation = state.orientation as any;
  if (request.margins) {
    if (request.margins.left !== undefined) layout.leftMargin = state.margins.left;
    if (request.margins.right !== undefined) layout.rightMargin = state.margins.right;
    if (request.margins.top !== undefined) layout.topMargin = state.margins.top;
    if (request.margins.bottom !== undefined) layout.bottomMargin = state.margins.bottom;
  }
  if (request.zoom) {
    layout.zoom = (state.zoom.wide !== null || state.zoom.tall !== null
      ? { horizontalFitToPages: state.zoom.wide ?? 0, verticalFitToPages: state.zoom.tall ?? 0 }
      : { scale: state.zoom.scale ?? 100 }) as any;
  }
  if (request.printArea !== undefined) layout.setPrintArea(state.printArea ?? "");
  if (request.titleRows !== undefined) layout.setPrintTitleRows(state.titleRows ? state.titleRows.split(":").map((row) => `$${row}`).join(":") : "");
  const texts = layout.headersFooters.defaultForAllPages;
  for (const slot of SLOTS) {
    if (request.header?.[slot] !== undefined) (texts as any)[`${slot}Header`] = state.header[slot];
    if (request.footer?.[slot] !== undefined) (texts as any)[`${slot}Footer`] = state.footer[slot];
  }
}

export async function preparePageLayoutPlan(args: unknown): Promise<PageLayoutPlan> {
  preflightToolArgs("set_page_layout", args);
  const a = (args ?? {}) as { sheet?: string } & Record<string, unknown>;
  const request = parsePageRequest(a);
  const target = await captureTarget(a.sheet);
  const prepared = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    sheet.load(["id", "name"]);
    try { sheet.protection?.load("protected"); } catch { /* среда без сведений о защите */ }
    await ctx.sync();
    if ((sheet as any).protection?.protected) throw new ToolError(`Лист ${sheet.name} защищён: параметры страницы на нём не поменять. Снимите защиту.`);
    const before = await readPageState(ctx, sheet);
    const expected = applyPageRequest(before, request);
    if (!pageMismatches(before, expected, request).length) {
      throw new ToolError(`На листе ${sheet.name} уже так: ${describePageChange(before, expected, request).join("; ")}. Менять нечего.`);
    }
    const undo = isCustomUndoAvailable();
    return {
      kind: "set_page_layout" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      target: { ...target, sheetName: sheet.name },
      request,
      before,
      expected,
      preview: describePageChange(before, expected, request),
      undoAvailable: undo,
      ...(undo ? {} : { undoNote: "Отмена недоступна: монитор изменений Excel не активен." }),
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executePageLayoutPlan(plan: PageLayoutPlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(plan.target.sheetId);
    sheet.load(["id", "name"]);
    const current = await readPageState(ctx, sheet);
    if (JSON.stringify(current) !== JSON.stringify(plan.before)) {
      throw new ToolExecutionError(`Параметры страницы листа ${sheet.name} изменились после предпросмотра. Операция не выполнялась — сделайте новый предпросмотр.`, "failed_before_write");
    }
    try {
      writePage(sheet, plan.expected, plan.request);
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(`Excel отказал в изменении параметров страницы листа ${sheet.name}: ${error?.message ?? error}. Перечитайте параметры страницы.`, "unknown");
    }
    const after = await readPageState(ctx, sheet);
    const problems = pageMismatches(after, plan.expected, plan.request);

    let undoRecorded = false;
    if (plan.undoAvailable) {
      const sheetId = plan.target.sheetId;
      const { before, request } = plan;
      undoRecorded = push(action(`параметры страницы листа ${sheet.name}`, async () => {
        const revision = getStructuralRevision();
        await Excel.run(async (undoCtx) => {
          const undoSheet = undoCtx.workbook.worksheets.getItem(sheetId);
          const now = await readPageState(undoCtx, undoSheet);
          const changed = pageMismatches(now, after, request);
          if (changed.length) throw new Error(`Параметры страницы изменили после операции агента (${changed[0]}). Отмена остановлена, чтобы не затереть более свежую настройку.`);
          if (getStructuralRevision() !== revision) throw new Error("Структура книги изменилась во время отмены. Отмена остановлена.");
          writePage(undoSheet, before, request);
          await undoCtx.sync();
          const back = await readPageState(undoCtx, undoSheet);
          const left = pageMismatches(back, before, request);
          if (left.length) throw new Error(`Отмена вернула не всё: ${left.join("; ")}.`);
        });
      }));
    }

    if (problems.length) {
      throw new ToolExecutionError(
        `Параметры страницы листа ${sheet.name} записаны, но обратное чтение расходится с планом: ${problems.join("; ")}. ` +
        (undoRecorded ? "Их можно вернуть кнопкой «Отменить»." : "Проверьте параметры страницы."),
        "applied"
      );
    }
    return {
      ok: true,
      executionState: "verified",
      sheet: sheet.name,
      changes: plan.preview,
      note: "Параметры страницы проверены обратным чтением. Как ляжет печать на бумагу, панель не видит — это покажет предварительный просмотр Excel.",
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: plan.undoNote ?? "Автоматическая отмена этой операции недоступна." })
    };
  });
}
