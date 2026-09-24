/**
 * Цветовая конвенция финансовой модели (этап 7, 7.4.4).
 *
 * Цвет текста ставится по тому, что в ячейке на самом деле: введённое число,
 * формула на этом листе, формула со ссылкой на другой лист или книгу,
 * контрольные ячейки, которые назвал пользователь. Палитра — настройка
 * задачи, а не стандарт: по умолчанию синий, чёрный и зелёный, и ответ это
 * говорит. Текст (подписи) и пустые ячейки не трогаются.
 *
 * Условные правила и статические цвета не смешиваются незаметно: если на
 * области есть правило, задающее цвет текста, предпросмотр его называет —
 * там, где оно срабатывает, видно его цвет, а не цвет конвенции.
 */

import { contains, parseA1Rect } from "./a1";
import {
  assertPlanWorkbook,
  assertTargetWritable,
  checkAddress,
  deepFreeze,
  preflightToolArgs,
  rangeOf,
  ToolError,
  ToolExecutionError
} from "./excelTools";
import { functionNamesIn } from "./functionProbe";
import { columnLetters } from "./formulaFill";
import { normalizeColor } from "./formatProps";
import { describeRule } from "./ruleOrderPlans";
import { formulaReferences } from "./rowOps";
import { readRuleSnapshots } from "./sheetFormatPlans";
import { captureExactFormat, exactFormatUndo, isCustomUndoAvailable, push } from "./undo";
import { captureTarget, type WorkbookTarget } from "./workbookContext";

export type CellRole = "input" | "formula" | "link" | "check";

export const DEFAULT_PALETTE: Record<CellRole, string> = {
  input: "#0000FF",
  formula: "#000000",
  link: "#008000",
  check: "#C00000"
};

const ROLE_TEXT: Record<CellRole, string> = {
  input: "вход",
  formula: "формула на этом листе",
  link: "ссылка на другой лист или книгу",
  check: "контрольная ячейка"
};

/** Предел: цвет ставится и снимается для отмены поячеечно. */
export const MAX_CONVENTION_CELLS = 2_000;

/**
 * Роль ячейки по её содержимому. null — не красится: текст, пусто, логическое.
 *
 * Формула без ссылок и функций (`=1000`, `=12*4`) — это введённое число,
 * записанное формулой: оно вход, а не расчёт.
 */
export function cellRole(formula: unknown, sheet: string): Exclude<CellRole, "check"> | null {
  if (typeof formula === "number") return "input";
  if (typeof formula !== "string" || !formula.startsWith("=")) return null;
  const bare = formula.replace(/"(?:[^"]|"")*"/g, '""');
  if (/\[[^\]]+\][^!]*!/.test(bare)) return "link";
  const references = formulaReferences(formula);
  if (references.some((reference) => reference.sheet !== null && reference.sheet.trim().toLowerCase() !== sheet.trim().toLowerCase())) return "link";
  if (!references.length && !functionNamesIn(formula).length && !/[A-Za-z_Ѐ-ӿ]/.test(bare.slice(1))) return "input";
  return "formula";
}

/**
 * Годы в шапке — подписи, а не входы. Проверка 24 сентября 2026 года: шапка
 * «2025, 2026» записана числами, и конвенция окрасила её как входы.
 * Годом считается целое 1900–2100, и только если все числа первой строки
 * области такие: строку с настоящими входами эта проверка не заденет.
 */
export function headerYears(firstRow: readonly unknown[]): number[] {
  const numbers = firstRow.map((value, index) => ({ value, index })).filter((item) => typeof item.value === "number");
  if (!numbers.length || !numbers.every((item) => Number.isInteger(item.value) && (item.value as number) >= 1900 && (item.value as number) <= 2100)) return [];
  return numbers.map((item) => item.index);
}

export interface ConventionCell {
  cell: string;
  role: CellRole;
  before: string | null;
}

export interface ConventionPlan {
  readonly kind: "apply_color_convention";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly resolvedAddress: string;
  readonly palette: Record<CellRole, string>;
  readonly paletteNote: string;
  readonly counts: Record<CellRole, number>;
  readonly cells: readonly ConventionCell[];
  /** Ячейки, где уже стоял другой, не чёрный цвет текста: он будет заменён. */
  readonly overwritten: readonly string[];
  readonly skipped: number;
  /** Числа-годы первой строки, принятые за подписи. */
  readonly yearLabels: readonly string[];
  readonly conditionalNote?: string;
  readonly signature: string;
  readonly undoAvailable: boolean;
  readonly undoNote: string;
  readonly createdAt: string;
}

function parsePalette(raw: unknown): { palette: Record<CellRole, string>; custom: boolean } {
  const palette = { ...DEFAULT_PALETTE };
  if (raw === undefined) return { palette, custom: false };
  if (!raw || typeof raw !== "object") throw new ToolError("palette — объект с цветами input, formula, link, check в HEX.");
  for (const [key, value] of Object.entries(raw as Record<string, unknown>)) {
    if (!(key in palette)) throw new ToolError(`В palette неизвестная роль «${key}»: есть input, formula, link, check.`);
    try { palette[key as CellRole] = normalizeColor(String(value)); } catch { throw new ToolError(`Цвет «${String(value)}» для ${key} — не HEX вида #RRGGBB.`); }
  }
  return { palette, custom: true };
}

export async function prepareConventionPlan(args: unknown): Promise<ConventionPlan> {
  preflightToolArgs("apply_color_convention", args);
  const a = args as { sheet?: string; address: string; checks?: string; palette?: unknown };
  const { palette, custom } = parsePalette(a.palette);
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
    assertTargetWritable(sheet, range, "format");
    const total = range.rowCount * range.columnCount;
    if (total > MAX_CONVENTION_CELLS) {
      throw new ToolError(`Цвета конвенции ставятся на область до ${MAX_CONVENTION_CELLS} ячеек; ${range.address} содержит ${total}. Разбейте модель на части.`);
    }
    const resolvedAddress = String(range.address).replace(/^.*!/, "");
    const outer = parseA1Rect(resolvedAddress)!;
    const checks = a.checks ? parseA1Rect(checkAddress(a.checks)) : null;
    if (a.checks && (!checks || !contains(outer, checks))) throw new ToolError(`Контрольные ячейки ${a.checks} должны лежать внутри ${resolvedAddress}.`);

    range.load("formulas");
    const colors = Array.from({ length: range.rowCount }, (_, r) => Array.from({ length: range.columnCount }, (_, c) => {
      const cell = range.getCell(r, c);
      cell.format.font.load("color");
      return cell;
    }));
    await ctx.sync();

    const cells: ConventionCell[] = [];
    const overwritten: string[] = [];
    let skipped = 0;
    const years = new Set(range.rowCount > 1 ? headerYears((range.formulas as unknown[][])[0]) : []);
    const yearLabels: string[] = [];
    (range.formulas as unknown[][]).forEach((row, r) => row.forEach((formula, c) => {
      const rowNumber = range.rowIndex + r + 1;
      const columnNumber = range.columnIndex + c + 1;
      if (r === 0 && years.has(c)) { yearLabels.push(`${columnLetters(columnNumber)}${rowNumber}`); skipped += 1; return; }
      const base = cellRole(formula, sheet.name);
      if (!base) { skipped += 1; return; }
      const inChecks = checks && rowNumber >= checks.rowStart && rowNumber <= checks.rowEnd && columnNumber >= checks.columnStart && columnNumber <= checks.columnEnd;
      const role: CellRole = inChecks ? "check" : base;
      const before = (colors[r][c].format.font.color ?? null) as string | null;
      const name = `${columnLetters(columnNumber)}${rowNumber}`;
      cells.push({ cell: name, role, before });
      if (before && before.toUpperCase() !== "#000000" && before.toUpperCase() !== palette[role].toUpperCase()) overwritten.push(name);
    }));
    if (!cells.length) throw new ToolError(`В ${resolvedAddress} нет чисел и формул: красить нечего. Подписи и пустые ячейки конвенция не трогает.`);
    const counts = { input: 0, formula: 0, link: 0, check: 0 };
    for (const item of cells) counts[item.role] += 1;

    // Правила условного форматирования с цветом текста перекрывают статический цвет.
    let conditionalNote: string | undefined;
    try {
      const rules = await readRuleSnapshots(ctx, range);
      const coloring = rules.filter((rule) => rule.fontColor);
      if (coloring.length) {
        conditionalNote =
          `На области есть правила условного форматирования с цветом текста: ${coloring.map(describeRule).join("; ")}. ` +
          "Там, где они срабатывают, будет виден цвет правила, а не цвет конвенции. Правила эта операция не меняет.";
      }
    } catch { /* правила не прочитались — предупреждать не о чем достоверно */ }

    const undo = isCustomUndoAvailable();
    return {
      kind: "apply_color_convention" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      target: { ...target, sheetName: sheet.name },
      resolvedAddress,
      palette,
      paletteNote: custom
        ? "Палитра задана для этой задачи."
        : "Палитра по умолчанию: входы синим, формулы чёрным, ссылки на другие листы зелёным, контрольные ячейки тёмно-красным. Это выбор, а не стандарт — её можно заменить.",
      counts,
      cells,
      overwritten: overwritten.slice(0, 20),
      skipped,
      yearLabels,
      ...(conditionalNote ? { conditionalNote } : {}),
      signature: JSON.stringify(range.formulas),
      undoAvailable: undo,
      undoNote: undo ? "Отмена вернёт прежний цвет текста каждой ячейки." : "Отмена недоступна: монитор изменений Excel не активен.",
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executeConventionPlan(plan: ConventionPlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(plan.target.sheetId);
    sheet.load("name");
    const range = sheet.getRange(plan.resolvedAddress);
    range.load("formulas");
    await ctx.sync();
    const where = `${sheet.name}!${plan.resolvedAddress}`;
    if (JSON.stringify(range.formulas) !== plan.signature) {
      throw new ToolExecutionError(`Содержимое ${where} изменилось после предпросмотра: роли ячеек могли смениться. Цвета не ставились — сделайте новый предпросмотр.`, "failed_before_write");
    }
    const before = await captureExactFormat(ctx, sheet.name, plan.resolvedAddress, { keys: ["fontColor"] });
    try {
      for (const item of plan.cells) sheet.getRange(item.cell).format.font.color = plan.palette[item.role];
      await ctx.sync();
    } catch (error: any) {
      throw new ToolExecutionError(`Не удалось определить, какие цвета успели встать на ${where}: ${error?.message ?? error}. Перечитайте область.`, "unknown");
    }
    const after = await captureExactFormat(ctx, sheet.name, plan.resolvedAddress, { keys: ["fontColor"] });
    const undoRecorded = plan.undoAvailable ? push(exactFormatUndo("цвета конвенции", before, after)) : false;

    const origin = parseA1Rect(plan.resolvedAddress)!;
    const wrong: string[] = [];
    for (const item of plan.cells) {
      const cell = parseA1Rect(item.cell)!;
      const actual = after.cells[cell.rowStart - origin.rowStart]?.[cell.columnStart - origin.columnStart]?.fontColor;
      if (String(actual ?? "").toUpperCase() !== plan.palette[item.role].toUpperCase()) wrong.push(`${item.cell}: ${String(actual)} вместо ${plan.palette[item.role]}`);
    }
    if (wrong.length) {
      throw new ToolExecutionError(
        `Цвета на ${where} поставлены, но обратное чтение расходится: ${wrong.slice(0, 8).join(", ")}${wrong.length > 8 ? ` и ещё ${wrong.length - 8}` : ""}. ${undoRecorded ? "Прежние цвета вернёт «Отменить»." : "Проверьте область."}`,
        "applied"
      );
    }
    const byRole = (role: CellRole) => plan.cells.filter((item) => item.role === role).map((item) => item.cell);
    return {
      ok: true,
      executionState: "verified",
      address: where,
      palette: plan.palette,
      paletteNote: plan.paletteNote,
      counts: plan.counts,
      examples: Object.fromEntries((["input", "formula", "link", "check"] as CellRole[])
        .filter((role) => plan.counts[role])
        .map((role) => [ROLE_TEXT[role], byRole(role).slice(0, 6)])),
      skippedTextOrEmpty: plan.skipped,
      ...(plan.yearLabels.length ? { yearLabels: plan.yearLabels, yearNote: "Годы в первой строке приняты за подписи шапки и не окрашены." } : {}),
      ...(plan.overwritten.length ? { overwritten: plan.overwritten, overwrittenNote: "В этих ячейках прежний цвет текста заменён цветом конвенции." } : {}),
      ...(plan.conditionalNote ? { conditionalNote: plan.conditionalNote } : {}),
      undoable: undoRecorded,
      undoNote: plan.undoNote
    };
  });
}
