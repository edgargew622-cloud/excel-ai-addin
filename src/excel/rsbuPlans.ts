/**
 * build_rsbu_reports: бухгалтерский баланс (форма 1) и отчёт о финансовых
 * результатах (форма 2) по оборотно-сальдовой ведомости (24.10.2026).
 *
 * Строки — формулами со ссылками на ячейки ОСВ (видно, из каких счетов
 * число, и форма пересчитается, если поправить ведомость на месте), итоги —
 * суммами строк. Значения сверяются с расчётом панели; проверки: актив =
 * пассив на обе даты, чистая прибыль формы 2 против изменения счёта 99,
 * счета с остатками, не попавшие в разнос.
 */

import { assertPlanWorkbook, checkAddress, deepFreeze, fitNewTable, preflightToolArgs, ToolError, ToolExecutionError, valuesForLiteralWrite } from "./excelTools";
import { columnLetters } from "./formulaFill";
import { checkSheetName, freeSheetName } from "./sheetRules";
import { action, getStructuralRevision, isCustomUndoAvailable, push } from "./undo";
import { captureTarget, type WorkbookTarget } from "./workbookContext";
import { BALANCE, computeForm, detectOsv, osvTotals, pickAccount, readOsv, RESULTS, unmappedAccounts, toStart, type ComputedLine, type FormLine, type OsvColumns, type OsvRow, type Term } from "./rsbuModel";

const quote = (sheet: string) => `'${sheet.replace(/'/g, "''")}'`;

interface SheetPlan { name: string; grid: unknown[][]; expected: { row: number; col: number; value: number }[]; bold: number[]; heads: number[]; checks: number[] }

/** Формула строки: ссылки на ячейки ОСВ со знаками; итог — сумма строк формы. */
function lineFormula(line: FormLine, rows: readonly OsvRow[], cols: OsvColumns, osv: { sheet: string; rowIndex: number; columnIndex: number }, start: boolean, rowOfCode: Map<string, number>, column: string): string {
  if (line.sumOf) {
    const parts = line.sumOf.filter((part) => rowOfCode.has(part.code)).map((part) => `${part.sign < 0 ? "-" : "+"}${column}${rowOfCode.get(part.code)}`);
    return parts.length ? "=" + parts.join("").replace(/^\+/, "") : "=0";
  }
  const terms: Term[] = start ? toStart(line.terms) : line.terms;
  const fieldCol: Record<string, number> = { sd: cols.startDebit, sc: cols.startCredit, td: cols.turnDebit, tc: cols.turnCredit, ed: cols.endDebit, ec: cols.endCredit };
  const refs: string[] = [];
  for (const term of terms) {
    const pick = pickAccount(rows, term.account);
    for (const row of pick.rows) {
      const cell = `${quote(osv.sheet)}!${columnLetters(osv.columnIndex + fieldCol[term.field] + 1)}${osv.rowIndex + row.row + 1}`;
      refs.push(`${term.sign < 0 ? "-" : "+"}N(${cell})`);
    }
  }
  return refs.length ? "=" + refs.join("").replace(/^\+/, "") : "=0";
}

function sheetFor(name: string, title: string, form: readonly FormLine[], computed: ComputedLine[], rows: readonly OsvRow[], cols: OsvColumns, osv: { sheet: string; rowIndex: number; columnIndex: number }, labels: { now: string; before: string | null }, org: string): SheetPlan {
  const grid: unknown[][] = [];
  const bold: number[] = []; const heads: number[] = []; const checks: number[] = [];
  const width = labels.before ? 5 : 4;
  const row = (...cells: unknown[]) => { const line = Array.from({ length: width }, (_, i) => cells[i] ?? ""); grid.push(line); return grid.length; };
  row(title);
  row(org ? `Организация: ${org}` : "Организация: ________________");
  row(`Источник: оборотно-сальдовая ведомость, лист «${osv.sheet}». Единица измерения: как в ведомости.`);
  row();
  const head = row(...(labels.before ? ["Наименование показателя", "Код", labels.now, labels.before, "Из каких счетов"] : ["Наименование показателя", "Код", labels.now, "Из каких счетов"]));
  heads.push(head);
  const rowOfCode = new Map<string, number>();
  const expected: SheetPlan["expected"] = [];
  computed.forEach((item, index) => {
    const line = item.line;
    if (line.section === "head") { heads.push(row(line.name)); return; }
    const at = grid.length + 1;
    rowOfCode.set(line.code, at);
    const now = lineFormula(line, rows, cols, osv, false, rowOfCode, "C");
    const before = labels.before ? lineFormula(line, rows, cols, osv, true, rowOfCode, "D") : null;
    const explain = line.sumOf ? line.sumOf.map((part, i) => `${i ? (part.sign < 0 ? " − " : " + ") : part.sign < 0 ? "−" : ""}${part.code}`).join("") : (line.note ?? line.terms.map((t) => `${t.sign < 0 ? "−" : ""}${t.account}`).join(" "));
    const missing = line.terms.length && !item.used.length && item.now === 0 ? " (в ведомости нет)" : "";
    row(line.name, line.code, now, ...(before !== null ? [before] : []), explain + missing);
    expected.push({ row: at, col: 2, value: item.now });
    if (labels.before && item.before !== null) expected.push({ row: at, col: 3, value: item.before });
    if (line.bold) bold.push(at);
    void index;
  });
  if (form === BALANCE) {
    row();
    const check = row("Проверка: актив − пассив (должно быть 0)", "", `=C${rowOfCode.get("1600")}-C${rowOfCode.get("1700")}`, ...(labels.before ? [`=D${rowOfCode.get("1600")}-D${rowOfCode.get("1700")}`] : []), "1600 − 1700");
    checks.push(check);
  }
  return { name, grid, expected, bold, heads, checks };
}

export interface RsbuPlan {
  readonly kind: "build_rsbu_reports";
  readonly id: string;
  readonly target: WorkbookTarget;
  readonly osvSheet: string;
  readonly osvAddress: string;
  readonly sheets: readonly SheetPlan[];
  readonly summary: Record<string, unknown>;
  readonly preview: readonly string[];
  readonly undoAvailable: boolean;
  readonly createdAt: string;
}

export async function prepareRsbuPlan(args: unknown): Promise<RsbuPlan> {
  preflightToolArgs("build_rsbu_reports", args);
  const a = args as { sheet?: string; address?: string; organization?: string; periodEnd?: string; periodStart?: string; forms?: string[] };
  const target = await captureTarget(a.sheet);
  const prepared = await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    sheet.load("name");
    const range = a.address?.trim() ? sheet.getRange(checkAddress(a.address)) : sheet.getUsedRange(true);
    range.load(["address", "values", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
    const all = ctx.workbook.worksheets;
    all.load("items/name");
    await ctx.sync();
    if (range.rowCount * range.columnCount > 200_000) throw new ToolError("Ведомость больше 200 000 ячеек — укажите область address.");
    const values = range.values as unknown[][];
    const cols = detectOsv(values);
    if (!cols) {
      throw new ToolError(`На листе «${sheet.name}» не нашлась шапка оборотно-сальдовой ведомости: нужен столбец «Счёт» и группы «Сальдо на начало периода», «Обороты за период», «Сальдо на конец периода» с «Дебет / Кредит». Укажите лист с ОСВ (из 1С: Отчёты → Оборотно-сальдовая ведомость, сохранить в Excel).`);
    }
    const rows = readOsv(values, cols);
    if (rows.length < 3) throw new ToolError("В ведомости меньше трёх счетов — строить не из чего.");
    const totals = osvTotals(rows);
    const balanced = Math.abs(totals.endDebit - totals.endCredit) < 1 && Math.abs(totals.startDebit - totals.startCredit) < 1;
    const forms = (a.forms?.length ? a.forms : ["1", "2"]).map(String);
    const want1 = forms.some((f) => /1|баланс/i.test(f));
    const want2 = forms.some((f) => /2|финанс|результ/i.test(f));
    const withStart = rows.some((row) => row.sd || row.sc);
    const balance = computeForm(BALANCE, rows, withStart);
    const results = computeForm(RESULTS, rows, false);
    const byCode = (list: ComputedLine[], code: string) => list.find((item) => item.line.code === code);
    const assets = byCode(balance, "1600")!; const liabilities = byCode(balance, "1700")!;
    const has90 = rows.some((row) => row.account.startsWith("90."));
    const r99 = pickAccount(rows, "99").rows;
    const profit99 = r99.length ? Math.round((r99.reduce((s, r) => s + r.ec - r.ed, 0) - r99.reduce((s, r) => s + r.sc - r.sd, 0)) * 100) / 100 : null;
    const net = byCode(results, "2400")!.now;
    const unmapped = unmappedAccounts(rows, [BALANCE, RESULTS]);
    const labelNow = a.periodEnd?.trim() ? `На ${a.periodEnd.trim()}` : "На отчётную дату";
    const labelBefore = withStart ? (a.periodStart?.trim() ? `На ${a.periodStart.trim()}` : "На начало периода") : null;
    const names = all.items.map((item) => item.name);
    const pickName = (wanted: string) => { try { return checkSheetName(wanted, names); } catch { return freeSheetName(wanted, names); } };
    const osv = { sheet: sheet.name, rowIndex: range.rowIndex, columnIndex: range.columnIndex };
    const org = a.organization?.trim() ?? "";
    const sheets: SheetPlan[] = [];
    if (want1) {
      const name = pickName("Бухгалтерский баланс"); names.push(name);
      sheets.push(sheetFor(name, "Бухгалтерский баланс (форма по ОКУД 0710001)", BALANCE, balance, rows, cols, osv, { now: labelNow, before: labelBefore }, org));
    }
    if (want2) {
      if (!has90) throw new ToolError("В ведомости нет субсчетов 90 (90.01, 90.02…) — форму 2 по ней не построить. Сформируйте ОСВ с субсчетами («Детализация по субсчетам») или попросите только форму 1.");
      const name = pickName("Фин. результаты"); names.push(name);
      sheets.push(sheetFor(name, "Отчёт о финансовых результатах (форма по ОКУД 0710002)", RESULTS, results, rows, cols, osv, { now: a.periodEnd?.trim() ? `За период по ${a.periodEnd.trim()}` : "За отчётный период", before: null }, org));
    }
    const fmt = (n: number) => n.toLocaleString("ru-RU", { maximumFractionDigits: 2 });
    const summary = {
      accounts: rows.length,
      osvBalanced: balanced,
      osvTotals: totals,
      balance: { assets: assets.now, liabilities: liabilities.now, difference: Math.round((assets.now - liabilities.now) * 100) / 100, ...(withStart ? { assetsStart: assets.before, liabilitiesStart: liabilities.before } : {}) },
      ...(want2 ? { netProfit: net, ...(profit99 !== null ? { profitBy99: profit99, profitCheck: Math.abs(profit99 - net) < 1 ? "сходится" : `расходится на ${fmt(Math.round((net - profit99) * 100) / 100)}` } : {}) } : {}),
      unmappedAccounts: unmapped
    };
    const preview = [
      `Ведомость «${sheet.name}» ${String(range.address).replace(/^.*!/, "")}: ${rows.length} счетов и субсчетов; итоги ведомости ${balanced ? "сходятся (Дт = Кт)" : `НЕ сходятся: Дт ${fmt(totals.endDebit)} / Кт ${fmt(totals.endCredit)}`}.`,
      ...(want1 ? [`Форма 1 на листе «${sheets[0].name}»: актив ${fmt(assets.now)}, пассив ${fmt(liabilities.now)}${Math.abs(assets.now - liabilities.now) < 1 ? " — сходится" : ` — расхождение ${fmt(assets.now - liabilities.now)}`}${withStart ? `; на начало: ${fmt(assets.before ?? 0)} / ${fmt(liabilities.before ?? 0)}` : ""}.`] : []),
      ...(want2 ? [`Форма 2 на листе «${sheets[sheets.length - 1].name}»: выручка ${fmt(byCode(results, "2110")!.now)}, чистая прибыль ${fmt(net)}${profit99 !== null ? ` (по счёту 99 — ${fmt(profit99)})` : ""}.`] : []),
      ...(unmapped.length ? [`Счета с остатками, которых нет в разносе: ${unmapped.join(", ")} — их нужно разнести вручную.`] : []),
      "Строки — формулами со ссылками на ячейки ведомости. Разнос — типовой; долгосрочная часть займов и вложений, проценты в 91 — по аналитике, вручную."
    ];
    return {
      kind: "build_rsbu_reports" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      target: { ...target, sheetName: sheet.name },
      osvSheet: sheet.name,
      osvAddress: String(range.address).replace(/^.*!/, ""),
      sheets,
      summary,
      preview,
      undoAvailable: isCustomUndoAvailable(),
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executeRsbuPlan(plan: RsbuPlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const all = ctx.workbook.worksheets;
    all.load("items/name");
    await ctx.sync();
    const taken = plan.sheets.filter((item) => all.items.some((w) => w.name.toLowerCase() === item.name.toLowerCase()));
    if (taken.length) throw new ToolExecutionError(`Листы ${taken.map((t) => `«${t.name}»`).join(", ")} появились после предпросмотра. Ничего не записано.`, "failed_before_write");
    const created: { id: string; name: string; address: string }[] = [];
    const problems: string[] = [];
    try {
      for (const item of plan.sheets) {
        const sheet = ctx.workbook.worksheets.add(item.name);
        sheet.load(["id", "name"]);
        const width = item.grid[0].length;
        const area = sheet.getRangeByIndexes(0, 0, item.grid.length, width);
        // Формулы — как формулы, текст — как текст.
        area.formulas = item.grid.map((line) => line.map((cell) => (typeof cell === "string" && cell.startsWith("=") ? cell : valuesForLiteralWrite([[cell]])[0][0]))) as any;
        sheet.getRangeByIndexes(0, 0, 1, 1).format.font.set({ bold: true, size: 13 });
        for (const head of item.heads) {
          const r = sheet.getRangeByIndexes(head - 1, 0, 1, width);
          r.format.font.bold = true;
          r.format.fill.color = head === item.heads[0] ? "#1F4E79" : "#E8EEF5";
          if (head === item.heads[0]) { r.format.font.color = "#FFFFFF"; r.format.wrapText = true; }
        }
        for (const b of item.bold) sheet.getRangeByIndexes(b - 1, 0, 1, width).format.font.bold = true;
        const numbers = sheet.getRangeByIndexes(item.heads[0], 2, item.grid.length - item.heads[0], width === 5 ? 2 : 1);
        numbers.numberFormat = Array.from({ length: item.grid.length - item.heads[0] }, () => Array.from({ length: width === 5 ? 2 : 1 }, () => "#,##0.00;(#,##0.00);\"–\"")) as any;
        sheet.getRangeByIndexes(item.heads[0] - 1, 0, item.grid.length - item.heads[0] + 1, width).format.borders.getItem("InsideHorizontal").style = "Continuous";
        for (const check of item.checks) sheet.getRangeByIndexes(check - 1, 0, 1, width).format.font.italic = true;
        sheet.freezePanes.freezeRows(item.heads[0]);
        await ctx.sync();
        await fitNewTable(ctx, sheet, 0, width);
        // Сверка с расчётом панели.
        const back = area;
        back.load("values");
        await ctx.sync();
        const values = back.values as unknown[][];
        for (const want of item.expected) {
          const got = values[want.row - 1]?.[want.col];
          if (typeof got !== "number" || Math.abs(got - want.value) > 0.01) problems.push(`«${item.name}» ${columnLetters(want.col + 1)}${want.row}: ${String(got)} вместо ${want.value}`);
        }
        created.push({ id: sheet.id, name: sheet.name, address: `A1:${columnLetters(width)}${item.grid.length}` });
      }
    } catch (error: any) {
      throw new ToolExecutionError(`Построение отчётности прервалось: ${error?.message ?? error}. Созданные листы: ${created.map((c) => c.name).join(", ") || "нет"}.`, created.length ? "applied" : "unknown");
    }
    let undoRecorded = false;
    if (plan.undoAvailable) {
      undoRecorded = push(action(`отчётность по ОСВ (${created.map((c) => c.name).join(", ")})`, async () => {
        const revision = getStructuralRevision();
        await Excel.run(async (undoCtx) => {
          if (getStructuralRevision() !== revision) throw new Error("Структура книги изменилась во время отмены. Отмена остановлена.");
          for (const item of created) undoCtx.workbook.worksheets.getItemOrNullObject(item.id).delete();
          await undoCtx.sync();
        });
      }));
    }
    if (problems.length) throw new ToolExecutionError(`Листы построены, но формулы дали не то, что посчитала панель: ${problems.slice(0, 5).join("; ")}.`, "applied");
    return {
      ok: true,
      executionState: "verified",
      sheets: created.map((c) => c.name),
      ...plan.summary,
      note: "Строки сверены с расчётом панели. Это черновик по типовому разносу: перед сдачей бухгалтер проверяет долгосрочную часть займов и вложений, проценты в 91, оценочные обязательства и счета без разноса (unmappedAccounts). Назови пользователю итоги и все расхождения.",
      undoable: undoRecorded,
      ...(undoRecorded ? {} : { undoNote: "Автоматическая отмена недоступна: листы удаляются вручную." })
    };
  });
}
