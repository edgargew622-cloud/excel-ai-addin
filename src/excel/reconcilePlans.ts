/**
 * reconcile_ranges — сверка двух таблиц книги (этап 11, PLAN_11_RECONCILE.md).
 *
 * Считает ядро (reconcileModel.ts), а не модель. Результат — новый лист:
 * итог по разделам, пары с основанием и ссылками на исходные строки,
 * «Вероятно» со столбцом «Решение», строки без пары с ближайшим кандидатом,
 * параметры сверки. «Повтори сверку» читает параметры и решения с прошлого
 * листа и пишет новый — прошлый остаётся для сравнения.
 *
 * Источники не меняются. Записанное перечитывается и сверяется ячейка в
 * ячейку; отмена убирает лист результата.
 */

import { assertPlanWorkbook, checkAddress, deepFreeze, fitNewTable, setColumnChars, preflightToolArgs, ToolError, ToolExecutionError, valuesForLiteralWrite } from "./excelTools";
import { columnLetters } from "./formulaFill";
import {
  DEFAULT_TOLERANCES,
  detectColumns,
  normalizeName,
  readSide,
  reconcile,
  SECTION_ORDER,
  SECTION_TITLE,
  toDay,
  toNumber,
  type MatchItem,
  type Section,
  type SideColumns,
  type SideRow,
  type Tolerances
} from "./reconcileModel";
import { checkSheetName } from "./sheetRules";
import { action, getStructuralRevision, isCustomUndoAvailable, push } from "./undo";
import { captureTarget, type WorkbookTarget } from "./workbookContext";

const PANEL_VERSION = typeof __PANEL_VERSION__ === "string" ? __PANEL_VERSION__ : "разработка";

export const MAX_RECONCILE_ROWS = 20_000;
const WIDTH = 14;
const PARAMS_MARK = "ПАРАМЕТРЫ СВЕРКИ (для повтора, не менять)";
const WRITE_CELLS_PER_SYNC = 5_000;

interface SourceRef { sheet: string; address: string; firstRow: number; firstColumn: number; header: string[]; signature: string }

interface ColumnNames { date?: string; amount?: string; debit?: string; credit?: string; names?: string[]; inns?: string[]; texts?: string[]; docs?: string[] }

export interface ReconcilePlan {
  readonly kind: "reconcile_ranges";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly left: SourceRef;
  readonly right: SourceRef;
  readonly leftColumns: ColumnNames;
  readonly rightColumns: ColumnNames;
  readonly tolerances: Tolerances;
  readonly destSheet: string;
  readonly grid: readonly (readonly unknown[])[];
  /** Номера строк листа результата: заголовки разделов, шапки, строки «Итога». */
  readonly layout: { titles: number[]; heads: number[]; summary: [number, number]; probable?: [number, number]; params: number; dateCells: [number, number][]; moneyColumns: number[] };
  readonly counts: Record<Section, number>;
  readonly totals: { left: number; right: number; leftRows: number; rightRows: number };
  readonly decisionsApplied: number;
  readonly repeatOf?: string;
  readonly preview: readonly string[];
  readonly undoAvailable: boolean;
  readonly createdAt: string;
}

/** Короткая подпись содержимого: источники не должны поменяться между предпросмотром и записью. */
function signatureOf(values: readonly (readonly unknown[])[]): string {
  let hash = 2166136261;
  const text = JSON.stringify(values);
  for (let i = 0; i < text.length; i++) { hash ^= text.charCodeAt(i); hash = Math.imul(hash, 16777619); }
  return `${values.length}:${(hash >>> 0).toString(16)}`;
}

function namesOf(columns: SideColumns, header: readonly string[]): ColumnNames {
  const one = (index?: number) => (index === undefined ? undefined : header[index]);
  const many = (list: number[]) => list.map((index) => header[index]);
  return {
    ...(columns.date !== undefined ? { date: one(columns.date) } : {}),
    ...(columns.amount !== undefined ? { amount: one(columns.amount) } : {}),
    ...(columns.debit !== undefined ? { debit: one(columns.debit) } : {}),
    ...(columns.credit !== undefined ? { credit: one(columns.credit) } : {}),
    ...(columns.names.length ? { names: many(columns.names) } : {}),
    ...(columns.inns.length ? { inns: many(columns.inns) } : {}),
    ...(columns.texts.length ? { texts: many(columns.texts) } : {}),
    ...(columns.docs.length ? { docs: many(columns.docs) } : {})
  };
}

/** Столбцы, названные пользователем или моделью, — по заголовку; остальное подбирается само. */
function columnsFor(header: string[], rows: unknown[][], given: ColumnNames | undefined, side: string): SideColumns {
  const auto = detectColumns(header, rows);
  if (!given) return auto;
  const find = (name: string) => {
    const index = header.findIndex((item) => item.trim().toLowerCase() === name.trim().toLowerCase());
    if (index < 0) throw new ToolError(`В ${side} таблице нет столбца «${name}». Заголовки: ${header.map((item) => `«${item}»`).join(", ")}.`);
    return index;
  };
  const result: SideColumns = { ...auto, names: [...auto.names], inns: [...auto.inns], texts: [...auto.texts], docs: [...auto.docs] };
  for (const key of ["date", "amount", "debit", "credit"] as const) {
    if (given[key]) { result[key] = find(given[key]!); if (key === "amount") { delete result.debit; delete result.credit; } }
  }
  for (const key of ["names", "inns", "texts", "docs"] as const) if (given[key]) result[key] = given[key]!.map(find);
  return result;
}

const sectionCount = (items: MatchItem[], section: Section) => items.filter((item) => item.section === section).length;

interface Built { grid: unknown[][]; layout: ReconcilePlan["layout"] }

/** Лист результата: заголовок, итог, разделы, параметры. */
function buildSheet(
  items: MatchItem[], left: { ref: SourceRef; rows: unknown[][]; side: SideRow[]; columns: SideColumns }, right: { ref: SourceRef; rows: unknown[][]; side: SideRow[]; columns: SideColumns },
  tolerances: Tolerances, params: unknown
): Built {
  const grid: unknown[][] = [];
  const blank = () => Array<unknown>(WIDTH).fill("");
  const row = (...cells: unknown[]) => { const line = blank(); cells.forEach((cell, i) => { line[i] = cell; }); grid.push(line); return grid.length; };
  const titles: number[] = []; const heads: number[] = []; const dateCells: [number, number][] = [];
  const refOf = (ref: SourceRef, index: number) => `${ref.sheet}!${ref.firstRow + 1 + index}`;
  const nameOf = (side: { rows: unknown[][]; columns: SideColumns }, index: number) => {
    const values = side.columns.names.map((column) => String(side.rows[index][column] ?? "")).filter(Boolean);
    const texts = side.columns.texts.map((column) => String(side.rows[index][column] ?? "")).filter(Boolean);
    return [...values, ...texts].join(" · ").slice(0, 200);
  };
  const describe = (side: typeof left, indexes: number[]) => {
    if (!indexes.length) return ["", "", "", ""];
    const rows = indexes.map((index) => side.side[index]);
    const sum = rows.reduce((total, item) => total + (item.amount ?? 0), 0);
    const day = rows[0].day;
    return [indexes.map((index) => refOf(side.ref, index)).join("; "), day ?? "", Math.round(sum * 100) / 100, indexes.map((index) => nameOf(side, index)).join(" | ")];
  };
  row(`Сверка: ${left.ref.sheet}!${left.ref.address} ↔ ${right.ref.sheet}!${right.ref.address}`);
  titles.push(grid.length);
  row();
  const summaryStart = row("Раздел", "Строк слева", "Сумма слева", "Строк справа", "Сумма справа");
  heads.push(summaryStart);
  for (const section of SECTION_ORDER) {
    const list = items.filter((item) => item.section === section);
    if (!list.length) continue;
    const sum = (side: typeof left, key: "left" | "right") => Math.round(list.flatMap((item) => item[key]).reduce((total, index) => total + (side.side[index].amount ?? 0), 0) * 100) / 100;
    row(SECTION_TITLE[section].split(" — ")[0], list.flatMap((item) => item.left).length, sum(left, "left"), list.flatMap((item) => item.right).length, sum(right, "right"));
  }
  const totalLeft = Math.round(left.side.reduce((total, item) => total + (item.amount ?? 0), 0) * 100) / 100;
  const totalRight = Math.round(right.side.reduce((total, item) => total + (item.amount ?? 0), 0) * 100) / 100;
  const summaryEnd = row("Всего", left.side.length, totalLeft, right.side.length, totalRight);
  if (left.side.some((item) => item.direction) || right.side.some((item) => item.direction)) {
    row("Суммы — по модулю: списания и поступления складываются; направление видно в исходной строке.");
  }
  let probable: [number, number] | undefined;
  for (const section of SECTION_ORDER) {
    const list = items.filter((item) => item.section === section);
    if (!list.length) continue;
    row();
    // Записей и строк бывает разное число: сочетание — одна запись на 2–3 строки.
    const leftCount = list.flatMap((item) => item.left).length;
    const rightCount = list.flatMap((item) => item.right).length;
    const rowsNote = (leftCount && leftCount !== list.length) || (rightCount && rightCount !== list.length) ? `; строк слева ${leftCount}, справа ${rightCount}` : "";
    row(`${SECTION_TITLE[section]} (${list.length}${rowsNote})`);
    titles.push(grid.length);
    const paired = section !== "leftOnly" && section !== "rightOnly" && section !== "fees";
    heads.push(row("№", paired ? "Основание" : "Причина", paired ? "Оценка" : "", "Разница", "Слева: строки", "Слева: дата", "Слева: сумма", "Слева: контрагент / назначение", "Справа: строки", "Справа: дата", "Справа: сумма", "Справа: контрагент / назначение", "Решение", "Ключ"));
    const first = grid.length + 1;
    list.forEach((item, n) => {
      const l = describe(left, item.left);
      const r = describe(right, item.right);
      // Ближайший кандидат — к основанию, а не в узкий столбец «Оценка» (Книга111: C раздувался до 50 знаков).
      const reason = !paired && item.nearest ? `${item.reason}; ближайшее — ${item.nearest}` : item.reason;
      const at = row(n + 1, reason, paired ? item.score : "", paired ? item.diff : "", ...l, ...r, "", section === "probable" ? item.key : "");
      if (l[1] !== "") dateCells.push([at, 6]);
      if (r[1] !== "") dateCells.push([at, 10]);
    });
    if (section === "probable") probable = [first, grid.length];
  }
  row();
  const paramsRow = row(PARAMS_MARK, JSON.stringify(params));
  row("Первая таблица", `${left.ref.sheet}!${left.ref.address}`, "столбцы", JSON.stringify(namesOf(left.columns, left.ref.header)));
  row("Вторая таблица", `${right.ref.sheet}!${right.ref.address}`, "столбцы", JSON.stringify(namesOf(right.columns, right.ref.header)));
  row("Допуски", `сумма ±${tolerances.amount}; дата ±${tolerances.days} дн.; название от ${Math.round(tolerances.name * 100)}%; комиссия до ${tolerances.fee}; сочетания ${tolerances.groups ? "да" : "нет"}`);
  row("Когда", new Date().toLocaleString("ru-RU"), "am.AI", PANEL_VERSION);
  return { grid, layout: { titles, heads, summary: [summaryStart, summaryEnd], ...(probable ? { probable } : {}), params: paramsRow, dateCells, moneyColumns: [3, 6, 10] } };
}

async function readSource(ctx: Excel.RequestContext, sheetName: string, address: string, label: string) {
  const sheet = ctx.workbook.worksheets.getItemOrNullObject(sheetName);
  sheet.load(["isNullObject", "name"]);
  await ctx.sync();
  if (sheet.isNullObject) throw new ToolError(`Листа «${sheetName}» для ${label} таблицы нет.`);
  const range = sheet.getRange(checkAddress(address));
  range.load(["address", "values", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
  await ctx.sync();
  if (range.rowCount < 2) throw new ToolError(`В ${label} таблице нужна шапка и хотя бы одна строка.`);
  if (range.rowCount - 1 > MAX_RECONCILE_ROWS) throw new ToolError(`В ${label} таблице ${range.rowCount - 1} строк, за раз — до ${MAX_RECONCILE_ROWS}. Сверьте частями, например по месяцам.`);
  const values = range.values as unknown[][];
  // Пустые строки внизу области — не строки данных.
  let last = values.length - 1;
  while (last > 0 && values[last].every((cell) => cell === "" || cell === null)) last--;
  const used = values.slice(0, last + 1);
  const header = used[0].map((cell, i) => String(cell ?? "").trim() || `Столбец ${columnLetters(range.columnIndex + i + 1)}`);
  const ref: SourceRef = {
    sheet: sheet.name,
    address: String(range.address).replace(/^.*!/, ""),
    firstRow: range.rowIndex + 1,
    firstColumn: range.columnIndex,
    header,
    signature: signatureOf(used)
  };
  return { ref, rows: used.slice(1) };
}

/** Параметры и решения с листа прошлой сверки. */
async function readPrevious(ctx: Excel.RequestContext, name: string) {
  const sheet = ctx.workbook.worksheets.getItemOrNullObject(name);
  sheet.load("isNullObject");
  await ctx.sync();
  if (sheet.isNullObject) throw new ToolError(`Листа прошлой сверки «${name}» нет.`);
  const used = sheet.getUsedRange(true);
  used.load("values");
  await ctx.sync();
  const values = used.values as unknown[][];
  const mark = values.find((line) => line[0] === PARAMS_MARK);
  if (!mark) throw new ToolError(`На листе «${name}» нет блока параметров сверки — это не лист reconcile_ranges или блок изменён.`);
  let params: any;
  try { params = JSON.parse(String(mark[1])); } catch { throw new ToolError(`Параметры на листе «${name}» повреждены — сверку можно запустить заново с теми же таблицами.`); }
  const decisions: Record<string, boolean> = {};
  for (const line of values) {
    const key = String(line[13] ?? "");
    const decision = String(line[12] ?? "").trim().toLowerCase();
    if (!key) continue;
    if (/^(да|yes|\+|1|верно|ок|ok)$/.test(decision)) decisions[key] = true;
    else if (/^(нет|no|-|0|неверно)$/.test(decision)) decisions[key] = false;
  }
  return { params, decisions };
}

export async function prepareReconcilePlan(args: unknown): Promise<ReconcilePlan> {
  preflightToolArgs("reconcile_ranges", args);
  const a = args as {
    sheet?: string; leftSheet?: string; leftAddress?: string; rightSheet?: string; rightAddress?: string;
    leftColumns?: ColumnNames; rightColumns?: ColumnNames;
    amountTolerance?: number; dayTolerance?: number; nameSimilarity?: number; feeTolerance?: number; groups?: boolean;
    resultSheet?: string; repeat?: string;
  };
  const target = await captureTarget(a.sheet);
  const prepared = await Excel.run(async (ctx) => {
    const active = ctx.workbook.worksheets.getItem(target.sheetId);
    active.load("name");
    const all = ctx.workbook.worksheets;
    all.load("items/name");
    await ctx.sync();
    let decisions: Record<string, boolean> = {};
    let base = a;
    if (a.repeat?.trim()) {
      const previous = await readPrevious(ctx, a.repeat.trim());
      decisions = previous.decisions;
      base = { ...previous.params, ...Object.fromEntries(Object.entries(a).filter(([key, value]) => key !== "repeat" && value !== undefined)) };
    }
    if (!base.leftAddress || !base.rightAddress) throw new ToolError("Укажите обе таблицы: leftAddress и rightAddress (и листы leftSheet, rightSheet, если они не на активном листе).");
    const left = await readSource(ctx, base.leftSheet?.trim() || active.name, base.leftAddress, "первой");
    const right = await readSource(ctx, base.rightSheet?.trim() || active.name, base.rightAddress, "второй");
    const leftColumns = columnsFor(left.ref.header, left.rows, base.leftColumns, "первой");
    const rightColumns = columnsFor(right.ref.header, right.rows, base.rightColumns, "второй");
    const needAmount = (columns: SideColumns, ref: SourceRef, label: string) => {
      if (columns.amount === undefined && columns.debit === undefined && columns.credit === undefined) {
        throw new ToolError(`В ${label} таблице не нашёлся столбец суммы. Заголовки: ${ref.header.map((item) => `«${item}»`).join(", ")}. Назовите его: ${label === "первой" ? "leftColumns" : "rightColumns"}.amount.`);
      }
    };
    needAmount(leftColumns, left.ref, "первой");
    needAmount(rightColumns, right.ref, "второй");
    const tolerances: Tolerances = {
      amount: Math.max(0, Number(base.amountTolerance ?? DEFAULT_TOLERANCES.amount)),
      days: Math.max(0, Math.round(Number(base.dayTolerance ?? DEFAULT_TOLERANCES.days))),
      name: Math.min(1, Math.max(0.3, Number(base.nameSimilarity ?? DEFAULT_TOLERANCES.name * 100) / 100)),
      fee: Math.max(0, Number(base.feeTolerance ?? DEFAULT_TOLERANCES.fee)),
      groups: base.groups !== false
    };
    const leftSide = readSide(left.rows, leftColumns);
    const rightSide = readSide(right.rows, rightColumns);
    const result = reconcile({
      left: leftSide, right: rightSide, tolerances, decisions,
      label: (side, index) => { const ref = side === "left" ? left.ref : right.ref; return `${ref.sheet}!${ref.firstRow + 1 + index}`; }
    });
    if (!result.coverage.ok) throw new ToolError("Сверка не учла каждую строку ровно один раз — это ошибка панели. Операция не выполнялась.");
    const params = {
      leftSheet: left.ref.sheet, leftAddress: left.ref.address, rightSheet: right.ref.sheet, rightAddress: right.ref.address,
      leftColumns: namesOf(leftColumns, left.ref.header), rightColumns: namesOf(rightColumns, right.ref.header),
      amountTolerance: tolerances.amount, dayTolerance: tolerances.days, nameSimilarity: Math.round(tolerances.name * 100), feeTolerance: tolerances.fee, groups: tolerances.groups
    };
    const built = buildSheet(result.items, { ref: left.ref, rows: left.rows, side: leftSide, columns: leftColumns }, { ref: right.ref, rows: right.rows, side: rightSide, columns: rightColumns }, tolerances, params);
    const names = all.items.map((item) => item.name);
    let destSheet = a.resultSheet?.trim() || "";
    if (!destSheet) {
      const now = new Date();
      const stem = `Сверка ${String(now.getDate()).padStart(2, "0")}.${String(now.getMonth() + 1).padStart(2, "0")}`;
      destSheet = stem;
      for (let n = 2; names.some((name) => name.toLowerCase() === destSheet.toLowerCase()); n++) destSheet = `${stem} (${n})`;
    }
    destSheet = checkSheetName(destSheet, names);
    const counts = Object.fromEntries(SECTION_ORDER.map((section) => [section, sectionCount(result.items, section)])) as Record<Section, number>;
    const totals = {
      left: Math.round(leftSide.reduce((s, item) => s + (item.amount ?? 0), 0) * 100) / 100,
      right: Math.round(rightSide.reduce((s, item) => s + (item.amount ?? 0), 0) * 100) / 100,
      leftRows: leftSide.length,
      rightRows: rightSide.length
    };
    const probableExamples = result.items.filter((item) => item.section === "probable").slice(0, 3).map((item) => `«Вероятно»: строки ${item.left.map((i) => i + 1).join("+")} слева и ${item.right.map((i) => i + 1).join("+")} справа — ${item.reason}`);
    const decisionsApplied = result.items.filter((item) => item.section === "confirmed").length + Object.values(decisions).filter((value) => value === false).length;
    const preview = [
      `Первая: ${left.ref.sheet}!${left.ref.address} — ${leftSide.length} строк, ${totals.left}. Столбцы: ${JSON.stringify(namesOf(leftColumns, left.ref.header))}`,
      `Вторая: ${right.ref.sheet}!${right.ref.address} — ${rightSide.length} строк, ${totals.right}. Столбцы: ${JSON.stringify(namesOf(rightColumns, right.ref.header))}`,
      `Допуски: сумма ±${tolerances.amount}, дата ±${tolerances.days} дн., название от ${Math.round(tolerances.name * 100)}%, комиссия до ${tolerances.fee}${tolerances.groups ? ", сочетания строк" : ""}`,
      SECTION_ORDER.filter((section) => counts[section]).map((section) => `${SECTION_TITLE[section].split(" — ")[0]}: ${counts[section]}`).join("; "),
      ...probableExamples,
      ...(a.repeat ? [`Повтор сверки «${a.repeat}»: учтено решений пользователя — ${decisionsApplied}`] : []),
      `Результат — новый лист «${destSheet}»; источники не меняются.`
    ];
    return {
      kind: "reconcile_ranges" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      target: { ...target, sheetName: active.name },
      left: left.ref,
      right: right.ref,
      leftColumns: namesOf(leftColumns, left.ref.header),
      rightColumns: namesOf(rightColumns, right.ref.header),
      tolerances,
      destSheet,
      grid: built.grid,
      layout: built.layout,
      counts,
      totals,
      decisionsApplied,
      ...(a.repeat ? { repeatOf: a.repeat } : {}),
      preview,
      undoAvailable: isCustomUndoAvailable(),
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executeReconcilePlan(plan: ReconcilePlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    // Источники не поменялись после предпросмотра.
    for (const ref of [plan.left, plan.right]) {
      const range = ctx.workbook.worksheets.getItem(ref.sheet).getRange(ref.address);
      range.load("values");
      await ctx.sync();
      const values = range.values as unknown[][];
      let last = values.length - 1;
      while (last > 0 && values[last].every((cell) => cell === "" || cell === null)) last--;
      if (signatureOf(values.slice(0, last + 1)) !== ref.signature) {
        throw new ToolExecutionError(`Таблица ${ref.sheet}!${ref.address} изменилась после предпросмотра. Ничего не записано — запустите сверку заново.`, "failed_before_write");
      }
    }
    const all = ctx.workbook.worksheets;
    all.load("items/name");
    await ctx.sync();
    if (all.items.some((item) => item.name.toLowerCase() === plan.destSheet.toLowerCase())) {
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
    const rows = plan.grid.length;
    const perSync = Math.max(1, Math.floor(WRITE_CELLS_PER_SYNC / WIDTH));
    try {
      // Блок параметров пишется после подгонки ширины: Excel подбирает ширину по
      // всему столбцу, и JSON параметров раздувал столбцы B и D (Книга111).
      const paramsAt = plan.layout.params - 1;
      for (let start = 0; start < paramsAt; start += perSync) {
        const part = plan.grid.slice(start, Math.min(start + perSync, paramsAt));
        sheet.getRangeByIndexes(start, 0, part.length, WIDTH).values = valuesForLiteralWrite(part as unknown[][]) as any[][];
        await ctx.sync();
      }
      // Оформление: заголовки разделов, шапки, деньги, даты, «Решение» — список «да/нет».
      const at = (row: number, column: number, height = 1, width = 1) => sheet.getRangeByIndexes(row - 1, column, height, width);
      at(1, 0).format.font.bold = true;
      at(1, 0).format.font.size = 13;
      for (const row of plan.layout.titles.slice(1)) { const r = at(row, 0, 1, WIDTH); r.format.font.bold = true; r.format.fill.color = "#D9E1F2"; }
      for (const row of plan.layout.heads) { const r = at(row, 0, 1, WIDTH); r.format.font.bold = true; r.format.fill.color = "#F2F2F2"; r.format.wrapText = true; }
      at(plan.layout.summary[1], 0, 1, 5).format.font.bold = true;
      for (const column of plan.layout.moneyColumns) at(1, column, rows, 1).numberFormat = Array.from({ length: rows }, () => ["#,##0.00"]) as any;
      const summaryRows = plan.layout.summary[1] - plan.layout.summary[0] + 1;
      for (const column of [2, 4]) at(plan.layout.summary[0], column, summaryRows, 1).numberFormat = Array.from({ length: summaryRows }, () => ["#,##0.00"]) as any;
      at(plan.layout.summary[0], 1, plan.layout.summary[1] - plan.layout.summary[0] + 1, 1).numberFormat = Array.from({ length: plan.layout.summary[1] - plan.layout.summary[0] + 1 }, () => ["0"]) as any;
      at(plan.layout.summary[0], 3, plan.layout.summary[1] - plan.layout.summary[0] + 1, 1).numberFormat = Array.from({ length: plan.layout.summary[1] - plan.layout.summary[0] + 1 }, () => ["0"]) as any;
      for (const [row, column] of plan.layout.dateCells) at(row, column - 1).numberFormat = [["dd.mm.yyyy"]];
      if (plan.layout.probable) {
        const [from, to] = plan.layout.probable;
        const decision = at(from, 12, to - from + 1, 1);
        decision.format.fill.color = "#FFF2CC";
        try { decision.dataValidation.rule = { list: { inCellDropDown: true, source: "да,нет" } } as any; } catch { /* без списка — тоже можно вписать */ }
      }
      sheet.getRangeByIndexes(0, 13, rows, 1).format.font.color = "#A6A6A6";
      await ctx.sync();
      // Ширина — по содержимому (даты и суммы видны целиком), а не заданная
      // заранее: в «Книге11» даты в узких столбцах выглядели «####».
      await fitNewTable(ctx, sheet, 1, WIDTH - 2);
      // Столбец A — по названиям разделов итога (Excel по части области давал 91 знак);
      // длинные заголовки разделов видны поверх пустых соседних ячеек. Служебный ключ N — узкий.
      const [from, to] = plan.layout.summary;
      const labels = plan.grid.slice(from - 1, to).map((line) => String(line[0] ?? "").length);
      await setColumnChars(ctx, sheet, 0, Math.min(34, Math.max(12, ...labels) + 2));
      await setColumnChars(ctx, sheet, WIDTH - 1, 8);
      const tail = plan.grid.slice(paramsAt);
      sheet.getRangeByIndexes(paramsAt, 0, tail.length, WIDTH).values = valuesForLiteralWrite(tail as unknown[][]) as any[][];
      sheet.getRangeByIndexes(paramsAt, 0, tail.length, WIDTH).format.wrapText = false;
      await ctx.sync();
      // Служебная строка параметров — в одну строку, а не столбиком на полэкрана.
      const paramsLine = sheet.getRangeByIndexes(plan.layout.params - 1, 0, 1, WIDTH);
      paramsLine.format.wrapText = false;
      paramsLine.format.rowHeight = 15;
      paramsLine.format.font.color = "#A6A6A6";
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(`Запись листа «${plan.destSheet}» прервалась: ${error?.message ?? error}. Часть могла записаться.`, "unknown");
    }
    const address = `A1:${columnLetters(WIDTH)}${rows}`;
    const written = sheet.getRange(address);
    written.load(["values", "formulas"]);
    await ctx.sync();
    const back = written.values as unknown[][];
    const problems: string[] = [];
    plan.grid.forEach((line, r) => line.forEach((expected, c) => {
      const actual = back[r]?.[c];
      const same = typeof expected === "number" ? typeof actual === "number" && Math.abs(actual - expected) < 1e-6 : String(actual ?? "") === String(expected ?? "");
      if (!same && problems.length < 6) problems.push(`${columnLetters(c + 1)}${r + 1}: ${JSON.stringify(actual)} вместо ${JSON.stringify(expected)}`);
    }));

    let undoRecorded = false;
    if (plan.undoAvailable) {
      const sheetId = sheet.id;
      const signature = JSON.stringify(written.formulas);
      undoRecorded = push(action(`сверка на листе ${plan.destSheet}`, async () => {
        const revision = getStructuralRevision();
        await Excel.run(async (undoCtx) => {
          const undoSheet = undoCtx.workbook.worksheets.getItem(sheetId);
          const range = undoSheet.getRange(address);
          range.load("formulas");
          await undoCtx.sync();
          // Пометки «да/нет» в «Решении» — правка пользователя: лист с ними не удаляем молча.
          if (JSON.stringify(range.formulas) !== signature) throw new Error(`На листе «${plan.destSheet}» есть правки после сверки (например, решения). Отмена остановлена; лист можно удалить вручную.`);
          if (getStructuralRevision() !== revision) throw new Error("Структура книги изменилась во время отмены. Отмена остановлена.");
          undoSheet.delete();
          await undoCtx.sync();
        });
      }));
    }
    if (problems.length) {
      throw new ToolExecutionError(`Лист «${plan.destSheet}» записан, но ячейки расходятся с расчётом: ${problems.join("; ")}.` + (undoRecorded ? " Его можно убрать кнопкой «Отменить»." : ""), "applied");
    }
    return {
      ok: true,
      executionState: "verified",
      sheet: plan.destSheet,
      address,
      first: `${plan.left.sheet}!${plan.left.address}`,
      second: `${plan.right.sheet}!${plan.right.address}`,
      columns: { first: plan.leftColumns, second: plan.rightColumns },
      tolerances: plan.tolerances,
      sections: Object.fromEntries(SECTION_ORDER.filter((section) => plan.counts[section]).map((section) => [SECTION_TITLE[section].split(" — ")[0], plan.counts[section]])),
      totals: plan.totals,
      ...(plan.repeatOf ? { repeatOf: plan.repeatOf, decisionsApplied: plan.decisionsApplied } : {}),
      note: "Каждая строка обеих таблиц — ровно в одном разделе (проверено); лист сверен с расчётом ячейка в ячейку. " +
        "«Вероятно» не доказано: пользователь ставит «да» или «нет» в столбце «Решение», затем «повтори сверку» (repeat) учтёт решения. " +
        "Даты на листе — числа с форматом даты; сумма сторон — в блоке итога.",
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: "Автоматическая отмена недоступна: лист удаляется вручную." })
    };
  });
}

/** Для тестов и предпросмотра. */
export const _internal = { buildSheet, signatureOf, normalizeName, toDay, toNumber };
