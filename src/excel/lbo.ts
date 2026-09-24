/**
 * Модель LBO (этап 7, 7.5.2): покупка компании в долг, погашение долга из
 * свободного потока и доходность инвестора на выходе — на новом листе
 * формулами от блока допущений.
 *
 * План ставил LBO после проверенных связей долга, процентов и денежных
 * потоков — они проверены трёхотчётной моделью. Здесь те же правила: все
 * допущения обязательны, лист хранит валюту, единицы, период и источник,
 * каждая ячейка сверяется с расчётом панели.
 *
 * Контрольные равенства (каждое считает величину другим путём):
 * - источники = использование на входе: долг + капитал = цена + расходы;
 * - деньги на выходе = сумма свободных потоков − сумма погашений;
 * - долг на выходе = долг на входе − сумма погашений.
 *
 * Проценты — на долг на начало года: так в модели нет цикличности, и это
 * названо в упрощениях. Доходность считается по одному вложению и одному
 * выходу, без промежуточных дивидендов: IRR = MOIC^(1/лет) − 1.
 */

import { assertPlanWorkbook, deepFreeze, preflightToolArgs, ToolError, ToolExecutionError } from "./excelTools";
import { columnLetters } from "./formulaFill";
import { AMOUNT, line as makeLine, modelMismatches, SHARE, text, writeModelSheet, type ModelCell } from "./modelSheet";
import { checkSheetName } from "./sheetRules";
import { isCustomUndoAvailable } from "./undo";
import { currentWorkbookIdentity } from "./workbookContext";

export interface LboAssumptions {
  ebitda0: number;
  entryMultiple: number;
  debtMultiple: number;
  fees: number;
  ebitdaGrowth: number;
  daPct: number;
  capexPct: number;
  nwcPct: number;
  taxRate: number;
  interestRate: number;
  cashSweep: number;
  exitMultiple: number;
}

type Key = keyof LboAssumptions;

export const LBO_INPUTS: readonly { key: Key; label: string; kind: "amount" | "share" | "multiple" }[] = [
  { key: "ebitda0", label: "EBITDA года входа", kind: "amount" },
  { key: "entryMultiple", label: "Цена входа, множитель EV/EBITDA", kind: "multiple" },
  { key: "debtMultiple", label: "Долг на входе, множитель к EBITDA", kind: "multiple" },
  { key: "fees", label: "Расходы на сделку", kind: "amount" },
  { key: "ebitdaGrowth", label: "Рост EBITDA в год", kind: "share" },
  { key: "daPct", label: "Амортизация, доля EBITDA", kind: "share" },
  { key: "capexPct", label: "Капвложения, доля EBITDA", kind: "share" },
  { key: "nwcPct", label: "Оборотный капитал, доля прироста EBITDA", kind: "share" },
  { key: "taxRate", label: "Ставка налога", kind: "share" },
  { key: "interestRate", label: "Ставка по долгу", kind: "share" },
  { key: "cashSweep", label: "Доля свободного потока на погашение долга", kind: "share" },
  { key: "exitMultiple", label: "Цена выхода, множитель EV/EBITDA", kind: "multiple" }
];

export interface LboRequest {
  sheet: string;
  currency: string;
  units: string;
  source: string;
  entryYear: number;
  years: number;
  assumptions: LboAssumptions;
}

export function parseLboRequest(args: Record<string, any>): LboRequest {
  const raw = (args.assumptions ?? {}) as Record<string, unknown>;
  const missing = LBO_INPUTS.filter((input) => typeof raw[input.key] !== "number" || !Number.isFinite(raw[input.key] as number));
  const meta: string[] = [];
  if (typeof args.currency !== "string" || !args.currency.trim()) meta.push("валюта (currency)");
  if (typeof args.units !== "string" || !args.units.trim()) meta.push("единицы (units)");
  if (typeof args.source !== "string" || !args.source.trim()) meta.push("источник допущений (source)");
  if (missing.length || meta.length) {
    throw new ToolError(
      `Модель LBO не строится, пока не заданы все допущения. Не хватает: ${[...meta, ...missing.map((input) => `${input.label} (${input.key})`)].join("; ")}. ` +
        "Спроси их у пользователя — значения по умолчанию не подставляются."
    );
  }
  const a = raw as unknown as LboAssumptions;
  const shares = LBO_INPUTS.filter((input) => input.kind === "share" && input.key !== "ebitdaGrowth" && ((a[input.key] < 0) || a[input.key] > 1));
  if (shares.length) throw new ToolError(`Доли задаются от 0 до 1 (0,1 — это 10 %): ${shares.map((input) => `${input.key} = ${a[input.key]}`).join(", ")}.`);
  if (a.ebitdaGrowth <= -1) throw new ToolError("Рост EBITDA должен быть больше −100 %.");
  if (a.ebitda0 <= 0) throw new ToolError("EBITDA года входа должна быть положительной: LBO строится на прибыльной компании.");
  const multiples = LBO_INPUTS.filter((input) => input.kind === "multiple" && (a[input.key] < 0 || a[input.key] > 50));
  if (multiples.length || a.entryMultiple === 0 || a.exitMultiple === 0) throw new ToolError("Множители цены — положительные числа (например, 8 — это 8× EBITDA), долг — от 0.");
  if (a.fees < 0) throw new ToolError("Расходы на сделку не бывают отрицательными.");
  const equity = a.ebitda0 * a.entryMultiple + a.fees - a.ebitda0 * a.debtMultiple;
  if (equity <= 0) {
    throw new ToolError(`Долг (${a.ebitda0 * a.debtMultiple}) покрывает всю цену с расходами (${a.ebitda0 * a.entryMultiple + a.fees}): вложение инвестора было бы ${equity}. Уточни структуру сделки у пользователя.`);
  }
  const years = Number(args.years);
  const entryYear = Number(args.entryYear);
  if (!Number.isInteger(years) || years < 3 || years > 10) throw new ToolError("years — срок владения, целое от 3 до 10.");
  if (!Number.isInteger(entryYear) || entryYear < 1900 || entryYear > 2200) throw new ToolError("entryYear — год входа в сделку, например 2025.");
  return { sheet: String(args.sheet ?? "").trim(), currency: args.currency.trim(), units: args.units.trim(), source: args.source.trim(), entryYear, years, assumptions: a };
}

/* --- раскладка ------------------------------------------------------------------ */

export interface LboLayout {
  rows: ModelCell[][];
  columns: number;
  /** Номера строк листа с контрольными равенствами. */
  checkRows: number[];
  equity0: number;
  exitEquity: number;
  moic: number;
  irr: number;
  debtLeft: number;
  lowCoverage: string[];
}

type RowKey = "ebitda" | "da" | "ebit" | "interest" | "ebt" | "tax" | "ni" | "capex" | "dnwc" | "fcf" | "repay" | "debt" | "cash" | "coverage";

export function buildLbo(request: LboRequest): LboLayout {
  const a = request.assumptions;
  const n = request.years;
  const columns = Math.max(n + 2, 3);
  const rows: ModelCell[][] = [];
  const line = (label: string, cells: ModelCell[] = []) => makeLine(columns, label, cells);
  const blank = (): ModelCell => ({ formula: "", expected: "" });
  const cell = (formula: string, expected: number | string, format = AMOUNT): ModelCell => ({ formula, expected, format });

  rows.push(line("Модель LBO"));
  rows.push(line("Валюта и единицы", [text(`${request.currency}, ${request.units}`)]));
  rows.push(line("Период", [text(`вход в ${request.entryYear}, владение ${n} лет, выход в ${request.entryYear + n}`)]));
  rows.push(line("Источник допущений", [text(request.source)]));
  rows.push(line(""));
  rows.push(line("Допущения", [text("значение")]));
  const inputRow: Partial<Record<Key, number>> = {};
  for (const input of LBO_INPUTS) {
    inputRow[input.key] = rows.length + 1;
    rows.push(line(input.label, [{ formula: a[input.key], expected: a[input.key], format: input.kind === "share" ? SHARE : input.kind === "multiple" ? '0.0"x"' : AMOUNT, input: true }]));
  }
  const ref = (key: Key) => `$B$${inputRow[key]}`;
  rows.push(line(""));

  // Вход в сделку.
  const entryEv = a.ebitda0 * a.entryMultiple;
  const debt0 = a.ebitda0 * a.debtMultiple;
  const equity0 = entryEv + a.fees - debt0;
  const entry: Record<string, number> = {};
  const entryLine = (key: string, label: string, value: ModelCell) => { entry[key] = rows.length + 1; rows.push(line(label, [value])); };
  rows.push(line("Вход в сделку"));
  entryLine("ev", "Цена компании (EV)", cell(`=${ref("ebitda0")}*${ref("entryMultiple")}`, entryEv));
  entryLine("fees", "Расходы на сделку", cell(`=${ref("fees")}`, a.fees));
  entryLine("debt", "Долг", cell(`=${ref("ebitda0")}*${ref("debtMultiple")}`, debt0));
  entryLine("equity", "Вложение инвестора", cell(`=B${entry.ev}+B${entry.fees}-B${entry.debt}`, equity0));
  rows.push(line(""));

  // Годы владения.
  const order: RowKey[] = ["ebitda", "da", "ebit", "interest", "ebt", "tax", "ni", "capex", "dnwc", "fcf", "repay", "debt", "cash", "coverage"];
  const headerRow = rows.length + 1;
  const rowNo = Object.fromEntries(order.map((key, index) => [key, headerRow + 1 + index])) as Record<RowKey, number>;
  const col = (t: number) => columnLetters(2 + t);
  const at = (key: RowKey, t: number) => `${col(t)}${rowNo[key]}`;
  const v = Object.fromEntries(order.map((key) => [key, new Array(n + 1).fill(0)])) as Record<RowKey, number[]>;
  v.ebitda[0] = a.ebitda0;
  v.debt[0] = debt0;
  v.cash[0] = 0;
  const coverage: (number | string)[] = [""];
  for (let t = 1; t <= n; t++) {
    v.ebitda[t] = v.ebitda[t - 1] * (1 + a.ebitdaGrowth);
    v.da[t] = v.ebitda[t] * a.daPct;
    v.ebit[t] = v.ebitda[t] - v.da[t];
    v.interest[t] = v.debt[t - 1] * a.interestRate;
    v.ebt[t] = v.ebit[t] - v.interest[t];
    v.tax[t] = Math.max(0, v.ebt[t]) * a.taxRate;
    v.ni[t] = v.ebt[t] - v.tax[t];
    v.capex[t] = v.ebitda[t] * a.capexPct;
    v.dnwc[t] = (v.ebitda[t] - v.ebitda[t - 1]) * a.nwcPct;
    v.fcf[t] = v.ni[t] + v.da[t] - v.capex[t] - v.dnwc[t];
    v.repay[t] = Math.min(v.debt[t - 1], Math.max(0, v.fcf[t] * a.cashSweep));
    v.debt[t] = v.debt[t - 1] - v.repay[t];
    v.cash[t] = v.cash[t - 1] + v.fcf[t] - v.repay[t];
    coverage.push(v.interest[t] === 0 ? "" : v.ebitda[t] / v.interest[t]);
  }
  const labels: Record<RowKey, string> = {
    ebitda: "EBITDA", da: "Амортизация", ebit: "Операционная прибыль (EBIT)", interest: "Проценты (на долг начала года)", ebt: "Прибыль до налога",
    tax: "Налог", ni: "Чистая прибыль", capex: "Капвложения", dnwc: "Прирост оборотного капитала", fcf: "Свободный поток до погашения долга",
    repay: "Погашение долга", debt: "Долг на конец года", cash: "Деньги на конец года", coverage: "Покрытие процентов (EBITDA / проценты)"
  };
  const formula = (key: RowKey, t: number): string => {
    const c = (k: RowKey) => at(k, t);
    const p = (k: RowKey) => at(k, t - 1);
    if (t === 0) return key === "ebitda" ? `=${ref("ebitda0")}` : key === "debt" ? `=B${entry.debt}` : key === "cash" ? "=0" : "";
    switch (key) {
      case "ebitda": return `=${p("ebitda")}*(1+${ref("ebitdaGrowth")})`;
      case "da": return `=${c("ebitda")}*${ref("daPct")}`;
      case "ebit": return `=${c("ebitda")}-${c("da")}`;
      case "interest": return `=${p("debt")}*${ref("interestRate")}`;
      case "ebt": return `=${c("ebit")}-${c("interest")}`;
      case "tax": return `=MAX(0,${c("ebt")})*${ref("taxRate")}`;
      case "ni": return `=${c("ebt")}-${c("tax")}`;
      case "capex": return `=${c("ebitda")}*${ref("capexPct")}`;
      case "dnwc": return `=(${c("ebitda")}-${p("ebitda")})*${ref("nwcPct")}`;
      case "fcf": return `=${c("ni")}+${c("da")}-${c("capex")}-${c("dnwc")}`;
      case "repay": return `=MIN(${p("debt")},MAX(0,${c("fcf")}*${ref("cashSweep")}))`;
      case "debt": return `=${p("debt")}-${c("repay")}`;
      case "cash": return `=${p("cash")}+${c("fcf")}-${c("repay")}`;
      case "coverage": return `=IF(${c("interest")}=0,"",${c("ebitda")}/${c("interest")})`;
    }
  };
  rows.push(line("Показатель", Array.from({ length: n + 1 }, (_, t) => text(t === 0 ? `${request.entryYear} вход` : `${request.entryYear + t}`))));
  for (const key of order) {
    rows.push(line(labels[key], Array.from({ length: n + 1 }, (_, t) => {
      const f = formula(key, t);
      if (!f) return blank();
      return cell(f, key === "coverage" ? coverage[t] : v[key][t], key === "coverage" ? '0.0"x"' : AMOUNT);
    })));
  }
  rows.push(line(""));

  // Выход и доходность.
  const last = col(n);
  const exitEv = v.ebitda[n] * a.exitMultiple;
  const exitEquity = exitEv - v.debt[n] + v.cash[n];
  const moic = exitEquity / equity0;
  const irr = moic > 0 ? moic ** (1 / n) - 1 : -1;
  const exit: Record<string, number> = {};
  const exitLine = (key: string, label: string, value: ModelCell) => { exit[key] = rows.length + 1; rows.push(line(label, [value])); };
  rows.push(line(`Выход в ${request.entryYear + n}`));
  exitLine("ev", "Цена компании на выходе", cell(`=${last}${rowNo.ebitda}*${ref("exitMultiple")}`, exitEv));
  exitLine("debt", "Долг на выходе", cell(`=${last}${rowNo.debt}`, v.debt[n]));
  exitLine("cash", "Деньги на выходе", cell(`=${last}${rowNo.cash}`, v.cash[n]));
  exitLine("equity", "Капитал инвестора на выходе", cell(`=B${exit.ev}-B${exit.debt}+B${exit.cash}`, exitEquity));
  exitLine("moic", "Кратность денег (MOIC)", cell(`=B${exit.equity}/B${entry.equity}`, moic, '0.00"x"'));
  exitLine("irr", "Доходность (IRR)", cell(`=IF(B${exit.moic}<=0,-1,B${exit.moic}^(1/${n})-1)`, irr, SHARE));
  rows.push(line(""));

  // Контрольные равенства: каждое считает величину другим путём.
  rows.push(line("Контроль (должно быть 0)"));
  const checkRows: number[] = [];
  const checkLine = (label: string, formulaText: string) => { checkRows.push(rows.length + 1); rows.push(line(label, [cell(formulaText, 0, "0.000000")])); };
  checkLine("Источники − использование на входе", `=ROUND(B${entry.debt}+B${entry.equity}-B${entry.ev}-B${entry.fees},6)`);
  checkLine("Деньги на выходе − (Σ потоков − Σ погашений)", `=ROUND(B${exit.cash}-(SUM(${col(1)}${rowNo.fcf}:${last}${rowNo.fcf})-SUM(${col(1)}${rowNo.repay}:${last}${rowNo.repay})),6)`);
  checkLine("Долг на выходе − (долг на входе − Σ погашений)", `=ROUND(B${exit.debt}-(B${entry.debt}-SUM(${col(1)}${rowNo.repay}:${last}${rowNo.repay})),6)`);

  const lowCoverage = Array.from({ length: n }, (_, k) => k + 1)
    .filter((t) => typeof coverage[t] === "number" && (coverage[t] as number) < 2)
    .map((t) => `${request.entryYear + t}`);
  return { rows, columns, checkRows, equity0, exitEquity, moic, irr, debtLeft: v.debt[n], lowCoverage };
}

/* --- план ------------------------------------------------------------------------ */

export interface LboPlan {
  readonly kind: "build_lbo_model";
  readonly id: string;
  readonly workbook: { workbookSessionId: string; documentUrl: string };
  readonly request: LboRequest;
  readonly layout: LboLayout;
  readonly simplifications: string;
  readonly undoAvailable: boolean;
  readonly undoNote: string;
  readonly createdAt: string;
}

const SIMPLIFICATIONS =
  "Упрощения LBO: один транш долга; проценты на долг начала года (без цикличности); погашение — заданная доля свободного потока, не больше остатка долга; " +
  "амортизация, капвложения и оборотный капитал — доли EBITDA; налог на положительную прибыль; без промежуточных дивидендов, IRR = MOIC^(1/лет) − 1; " +
  "расходы на сделку платит инвестор.";

export async function prepareLboPlan(args: unknown): Promise<LboPlan> {
  preflightToolArgs("build_lbo_model", args);
  const request = parseLboRequest(args as Record<string, any>);
  const prepared = await Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    let sheet: string;
    try { sheet = checkSheetName(request.sheet, sheets.items.map((item) => item.name)); } catch (error: any) { throw new ToolError(`${error.message} Модель строится на новом листе.`); }
    const layout = buildLbo({ ...request, sheet });
    const undo = isCustomUndoAvailable();
    return {
      kind: "build_lbo_model" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      workbook: currentWorkbookIdentity(),
      request: { ...request, sheet },
      layout,
      simplifications: SIMPLIFICATIONS,
      undoAvailable: undo,
      undoNote: undo ? "Отмена удалит лист модели, если его не меняли после операции." : "Отмена недоступна: монитор изменений Excel не активен.",
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executeLboPlan(plan: LboPlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const { sheetName, values, undoRecorded } = await writeModelSheet(ctx, {
      sheetName: plan.request.sheet,
      rows: plan.layout.rows,
      columns: plan.layout.columns,
      undoAvailable: plan.undoAvailable,
      label: "модель LBO"
    });
    const mismatches = modelMismatches(plan.layout.rows, values);
    const checks = plan.layout.checkRows.map((row) => values[row - 1]?.[1]);
    const failed = checks.some((value) => value !== 0);
    if (mismatches.length || failed) {
      throw new ToolExecutionError(
        `Модель записана на лист «${sheetName}», но ${failed ? `контрольные равенства не сходятся (строки ${plan.layout.checkRows.join(", ")}: ${JSON.stringify(checks)})` : ""}` +
          `${failed && mismatches.length ? "; " : ""}${mismatches.length ? `значения расходятся с расчётом панели: ${mismatches.slice(0, 8).join("; ")}` : ""}. ` +
          `Готовой она не считается. ${undoRecorded ? "Лист уберёт «Отменить»." : ""}`,
        "applied"
      );
    }
    const { layout, request } = plan;
    return {
      ok: true,
      executionState: "verified",
      sheet: sheetName,
      period: `вход ${request.entryYear}, выход ${request.entryYear + request.years}`,
      currency: `${request.currency}, ${request.units}`,
      investorEquity: layout.equity0,
      exitEquity: layout.exitEquity,
      moic: layout.moic,
      irr: layout.irr,
      debtAtExit: layout.debtLeft,
      control: { rows: layout.checkRows, values: checks, note: "Источники = использование; деньги и долг на выходе сходятся с суммами потоков и погашений. Сверено." },
      checkedCells: layout.rows.length * layout.columns,
      ...(layout.lowCoverage.length
        ? { lowCoverage: layout.lowCoverage, lowCoverageNote: "В эти годы EBITDA покрывает проценты меньше чем вдвое: долговая нагрузка высокая. Скажи это пользователю." }
        : {}),
      ...(layout.debtLeft > 0 ? { debtNote: `К выходу долг погашен не полностью: остаток ${layout.debtLeft}. Он вычтен из капитала на выходе.` } : {}),
      simplifications: plan.simplifications,
      note: "Модель — формулы от блока допущений (синий текст). Панель проверила расчёт, но не допущения: решение о сделке принимает человек.",
      undoable: undoRecorded,
      undoNote: plan.undoNote
    };
  });
}
