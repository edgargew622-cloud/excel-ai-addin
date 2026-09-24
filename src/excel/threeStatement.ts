/**
 * Трёхотчётная модель (этап 7, 7.5.2–7.5.4): отчёт о прибылях, баланс,
 * движение денег — на новом листе, формулами от блока допущений.
 *
 * Правила плана, которые здесь исполняются:
 * - модель строится на формулах, а не на числах из чата: каждое допущение
 *   стоит в своей ячейке, отчёты ссылаются на него;
 * - лист хранит валюту, единицы, период и источник допущений;
 * - недостающее допущение — отказ до записи со списком, чего не хватает:
 *   значения по умолчанию модель не подставляет;
 * - баланс на начало обязан сходиться ещё до записи;
 * - после записи панель сверяет каждую ячейку со своим расчётом, а
 *   контрольное равенство «активы = обязательства + капитал» — по каждому
 *   году. Без этого «готово» не выдаётся.
 *
 * Упрощения названы в ответе: один темп роста, статьи оборотного капитала —
 * доли выручки, долг гасится равными суммами, проценты — на долг на начало
 * года, налог только с прибыли, денежный остаток без процентов.
 */

import { checkSheetName } from "./sheetRules";
import { assertPlanWorkbook, deepFreeze, preflightToolArgs, ToolError, ToolExecutionError } from "./excelTools";
import { columnLetters } from "./formulaFill";
import { AMOUNT, line as makeLine, modelMismatches as rowMismatches, SHARE, text, writeModelSheet, type ModelCell } from "./modelSheet";
import { isCustomUndoAvailable } from "./undo";
import { currentWorkbookIdentity } from "./workbookContext";

export interface Assumptions {
  revenue0: number;
  growth: number;
  cogsPct: number;
  opexPct: number;
  daPct: number;
  capexPct: number;
  receivablesPct: number;
  inventoryPct: number;
  payablesPct: number;
  taxRate: number;
  interestRate: number;
  repayment: number;
  payout: number;
  cash0: number;
  ppe0: number;
  receivables0: number;
  inventory0: number;
  payables0: number;
  debt0: number;
  equity0: number;
}

type Key = keyof Assumptions;

/** Допущения в порядке блока на листе: подпись, единица, доля ли это. */
export const INPUTS: readonly { key: Key; label: string; share: boolean }[] = [
  { key: "revenue0", label: "Выручка последнего фактического года", share: false },
  { key: "growth", label: "Рост выручки в год", share: true },
  { key: "cogsPct", label: "Себестоимость, доля выручки", share: true },
  { key: "opexPct", label: "Операционные расходы, доля выручки", share: true },
  { key: "daPct", label: "Амортизация, доля выручки", share: true },
  { key: "capexPct", label: "Капвложения, доля выручки", share: true },
  { key: "receivablesPct", label: "Дебиторская задолженность, доля выручки", share: true },
  { key: "inventoryPct", label: "Запасы, доля выручки", share: true },
  { key: "payablesPct", label: "Кредиторская задолженность, доля выручки", share: true },
  { key: "taxRate", label: "Ставка налога на прибыль", share: true },
  { key: "interestRate", label: "Ставка по долгу", share: true },
  { key: "repayment", label: "Погашение долга в год", share: false },
  { key: "payout", label: "Доля прибыли на дивиденды", share: true },
  { key: "cash0", label: "Деньги на начало", share: false },
  { key: "ppe0", label: "Основные средства на начало", share: false },
  { key: "receivables0", label: "Дебиторская задолженность на начало", share: false },
  { key: "inventory0", label: "Запасы на начало", share: false },
  { key: "payables0", label: "Кредиторская задолженность на начало", share: false },
  { key: "debt0", label: "Долг на начало", share: false },
  { key: "equity0", label: "Капитал на начало", share: false }
];

export interface ModelRequest {
  sheet: string;
  currency: string;
  units: string;
  source: string;
  firstYear: number;
  years: number;
  assumptions: Assumptions;
}

const MAX_YEARS = 10;

/** Проверка до Excel: всё ли задано и сходится ли баланс на начало. */
export function parseModelRequest(args: Record<string, any>): ModelRequest {
  const raw = (args.assumptions ?? {}) as Record<string, unknown>;
  const missing = INPUTS.filter((input) => typeof raw[input.key] !== "number" || !Number.isFinite(raw[input.key] as number));
  const meta: string[] = [];
  if (typeof args.currency !== "string" || !args.currency.trim()) meta.push("валюта (currency)");
  if (typeof args.units !== "string" || !args.units.trim()) meta.push("единицы (units)");
  if (typeof args.source !== "string" || !args.source.trim()) meta.push("источник допущений (source)");
  if (missing.length || meta.length) {
    throw new ToolError(
      `Модель не строится, пока не заданы все допущения. Не хватает: ${[...meta, ...missing.map((input) => `${input.label} (${input.key})`)].join("; ")}. ` +
        "Спроси их у пользователя — значения по умолчанию не подставляются."
    );
  }
  const a = raw as unknown as Assumptions;
  const shareProblems = INPUTS.filter((input) => input.share && input.key !== "growth" && ((a[input.key] as number) < 0 || (a[input.key] as number) > 1));
  if (shareProblems.length) {
    throw new ToolError(`Доли задаются от 0 до 1 (0,25 — это 25 %): ${shareProblems.map((input) => `${input.key} = ${a[input.key]}`).join(", ")}.`);
  }
  if (a.growth <= -1) throw new ToolError("Рост выручки должен быть больше −100 % (growth > −1).");
  const negative = (["revenue0", "repayment", "cash0", "ppe0", "receivables0", "inventory0", "payables0", "debt0"] as Key[]).filter((key) => a[key] < 0);
  if (negative.length) throw new ToolError(`Эти величины не бывают отрицательными: ${negative.join(", ")}.`);
  const assets = a.cash0 + a.ppe0 + a.receivables0 + a.inventory0;
  const liabilities = a.payables0 + a.debt0 + a.equity0;
  if (Math.abs(assets - liabilities) > 1e-6 * Math.max(1, Math.abs(assets))) {
    throw new ToolError(
      `Баланс на начало не сходится: активы ${assets} (деньги, основные средства, дебиторка, запасы), ` +
        `обязательства и капитал ${liabilities} (кредиторка, долг, капитал); разница ${assets - liabilities}. Уточни входы у пользователя.`
    );
  }
  const years = Number(args.years);
  const firstYear = Number(args.firstYear);
  if (!Number.isInteger(years) || years < 1 || years > MAX_YEARS) throw new ToolError(`years — целое от 1 до ${MAX_YEARS}.`);
  if (!Number.isInteger(firstYear) || firstYear < 1900 || firstYear > 2200) throw new ToolError("firstYear — первый прогнозный год, например 2026.");
  return {
    sheet: String(args.sheet ?? "").trim(),
    currency: args.currency.trim(),
    units: args.units.trim(),
    source: args.source.trim(),
    firstYear,
    years,
    assumptions: a
  };
}

/* --- раскладка ------------------------------------------------------------------ */

export type { ModelCell };

export interface ModelLayout {
  rows: ModelCell[][];
  columns: number;
  /** Строки контроля баланса, с 1 — номер строки листа. */
  checkRow: number;
  /** Годы, где деньги на конец отрицательны: модели не хватает финансирования. */
  negativeCash: string[];
}

type RowKey =
  | "rev" | "cogs" | "gross" | "opex" | "ebitda" | "da" | "ebit" | "interest" | "ebt" | "tax" | "ni"
  | "cash" | "rec" | "inv" | "ppe" | "assets" | "pay" | "debt" | "equity" | "le"
  | "cfNi" | "cfDa" | "cfRec" | "cfInv" | "cfPay" | "cfo" | "capex" | "cfi" | "cfDebt" | "div" | "cff" | "net" | "open" | "close"
  | "check";

/**
 * Лист модели: шапка с валютой и источником, блок допущений, три отчёта
 * по годам и контроль. Столбец B — последний фактический год (год 0),
 * C и дальше — прогноз. Ожидаемые значения считаются тем же путём, что
 * формулы, и после записи сверяются с Excel.
 */
export function buildModel(request: ModelRequest): ModelLayout {
  const a = request.assumptions;
  const n = request.years;
  const columns = n + 2;
  const rows: ModelCell[][] = [];
  const blank = (): ModelCell => ({ formula: "", expected: "" });
  const line = (label: string, cells: ModelCell[] = []) => makeLine(columns, label, cells);

  rows.push(line("Трёхотчётная модель"));
  rows.push(line("Валюта и единицы", [text(`${request.currency}, ${request.units}`)]));
  rows.push(line("Период", [text(`${request.firstYear - 1} — факт, ${request.firstYear}–${request.firstYear + n - 1} — прогноз`)]));
  rows.push(line("Источник допущений", [text(request.source)]));
  rows.push(line(""));
  rows.push(line("Допущения", [text("значение")]));
  const inputRow: Record<string, number> = {};
  for (const input of INPUTS) {
    inputRow[input.key] = rows.length + 1;
    rows.push(line(input.label, [{ formula: a[input.key], expected: a[input.key], format: input.share ? SHARE : AMOUNT, input: true }]));
  }
  const ref = (key: Key) => `$B$${inputRow[key]}`;
  rows.push(line(""));

  // Номера строк отчётов известны заранее: формулы ссылаются вперёд и назад.
  const order: (RowKey | { title: string })[] = [
    { title: "Отчёт о прибылях и убытках" },
    "rev", "cogs", "gross", "opex", "ebitda", "da", "ebit", "interest", "ebt", "tax", "ni",
    { title: "Баланс на конец года" },
    "cash", "rec", "inv", "ppe", "assets", "pay", "debt", "equity", "le",
    { title: "Движение денег" },
    "cfNi", "cfDa", "cfRec", "cfInv", "cfPay", "cfo", "capex", "cfi", "cfDebt", "div", "cff", "net", "open", "close",
    { title: "Контроль" },
    "check"
  ];
  const headerRow = rows.length + 1;
  const rowNo: Partial<Record<RowKey, number>> = {};
  order.forEach((item, index) => { if (typeof item === "string") rowNo[item] = headerRow + 1 + index; });
  const col = (t: number) => columnLetters(2 + t);
  const at = (key: RowKey, t: number) => `${col(t)}${rowNo[key]}`;

  // Ожидаемые значения по годам: тот же расчёт, что в формулах.
  const v: Record<RowKey, number[]> = Object.fromEntries(order.filter((item): item is RowKey => typeof item === "string").map((key) => [key, new Array(n + 1).fill(0)])) as any;
  v.rev[0] = a.revenue0;
  v.cash[0] = a.cash0; v.rec[0] = a.receivables0; v.inv[0] = a.inventory0; v.ppe[0] = a.ppe0;
  v.pay[0] = a.payables0; v.debt[0] = a.debt0; v.equity[0] = a.equity0;
  v.assets[0] = a.cash0 + a.ppe0 + a.receivables0 + a.inventory0;
  v.le[0] = a.payables0 + a.debt0 + a.equity0;
  v.check[0] = 0;
  for (let t = 1; t <= n; t++) {
    v.rev[t] = v.rev[t - 1] * (1 + a.growth);
    v.cogs[t] = -v.rev[t] * a.cogsPct;
    v.gross[t] = v.rev[t] + v.cogs[t];
    v.opex[t] = -v.rev[t] * a.opexPct;
    v.ebitda[t] = v.gross[t] + v.opex[t];
    v.da[t] = -v.rev[t] * a.daPct;
    v.ebit[t] = v.ebitda[t] + v.da[t];
    v.interest[t] = -v.debt[t - 1] * a.interestRate;
    v.ebt[t] = v.ebit[t] + v.interest[t];
    v.tax[t] = -Math.max(0, v.ebt[t]) * a.taxRate;
    v.ni[t] = v.ebt[t] + v.tax[t];
    v.rec[t] = v.rev[t] * a.receivablesPct;
    v.inv[t] = v.rev[t] * a.inventoryPct;
    v.pay[t] = v.rev[t] * a.payablesPct;
    v.capex[t] = -v.rev[t] * a.capexPct;
    v.ppe[t] = v.ppe[t - 1] - v.capex[t] + v.da[t];
    v.debt[t] = Math.max(0, v.debt[t - 1] - a.repayment);
    v.div[t] = -Math.max(0, v.ni[t]) * a.payout;
    v.equity[t] = v.equity[t - 1] + v.ni[t] + v.div[t];
    v.cfNi[t] = v.ni[t];
    v.cfDa[t] = -v.da[t];
    v.cfRec[t] = v.rec[t - 1] - v.rec[t];
    v.cfInv[t] = v.inv[t - 1] - v.inv[t];
    v.cfPay[t] = v.pay[t] - v.pay[t - 1];
    v.cfo[t] = v.cfNi[t] + v.cfDa[t] + v.cfRec[t] + v.cfInv[t] + v.cfPay[t];
    v.cfi[t] = v.capex[t];
    v.cfDebt[t] = v.debt[t] - v.debt[t - 1];
    v.cff[t] = v.cfDebt[t] + v.div[t];
    v.net[t] = v.cfo[t] + v.cfi[t] + v.cff[t];
    v.open[t] = v.cash[t - 1];
    v.close[t] = v.open[t] + v.net[t];
    v.cash[t] = v.close[t];
    v.assets[t] = v.cash[t] + v.rec[t] + v.inv[t] + v.ppe[t];
    v.le[t] = v.pay[t] + v.debt[t] + v.equity[t];
    v.check[t] = 0;
  }

  const labels: Record<RowKey, string> = {
    rev: "Выручка", cogs: "Себестоимость", gross: "Валовая прибыль", opex: "Операционные расходы", ebitda: "EBITDA",
    da: "Амортизация", ebit: "Операционная прибыль (EBIT)", interest: "Проценты", ebt: "Прибыль до налога", tax: "Налог на прибыль", ni: "Чистая прибыль",
    cash: "Деньги", rec: "Дебиторская задолженность", inv: "Запасы", ppe: "Основные средства", assets: "Итого активы",
    pay: "Кредиторская задолженность", debt: "Долг", equity: "Капитал", le: "Итого обязательства и капитал",
    cfNi: "Чистая прибыль", cfDa: "Амортизация (обратно)", cfRec: "Изменение дебиторской задолженности", cfInv: "Изменение запасов",
    cfPay: "Изменение кредиторской задолженности", cfo: "Операционный поток", capex: "Капвложения", cfi: "Инвестиционный поток",
    cfDebt: "Изменение долга", div: "Дивиденды", cff: "Финансовый поток", net: "Изменение денег", open: "Деньги на начало", close: "Деньги на конец",
    check: "Активы − обязательства и капитал (должно быть 0)"
  };
  const formula = (key: RowKey, t: number): string => {
    const c = (k: RowKey) => at(k, t);
    const p = (k: RowKey) => at(k, t - 1);
    if (t === 0) {
      const opening: Partial<Record<RowKey, string>> = {
        rev: `=${ref("revenue0")}`, cash: `=${ref("cash0")}`, rec: `=${ref("receivables0")}`, inv: `=${ref("inventory0")}`,
        ppe: `=${ref("ppe0")}`, pay: `=${ref("payables0")}`, debt: `=${ref("debt0")}`, equity: `=${ref("equity0")}`,
        assets: `=${c("cash")}+${c("rec")}+${c("inv")}+${c("ppe")}`, le: `=${c("pay")}+${c("debt")}+${c("equity")}`,
        check: `=ROUND(${c("assets")}-${c("le")},6)`
      };
      return opening[key] ?? "";
    }
    switch (key) {
      case "rev": return `=${p("rev")}*(1+${ref("growth")})`;
      case "cogs": return `=-${c("rev")}*${ref("cogsPct")}`;
      case "gross": return `=${c("rev")}+${c("cogs")}`;
      case "opex": return `=-${c("rev")}*${ref("opexPct")}`;
      case "ebitda": return `=${c("gross")}+${c("opex")}`;
      case "da": return `=-${c("rev")}*${ref("daPct")}`;
      case "ebit": return `=${c("ebitda")}+${c("da")}`;
      case "interest": return `=-${p("debt")}*${ref("interestRate")}`;
      case "ebt": return `=${c("ebit")}+${c("interest")}`;
      case "tax": return `=-MAX(0,${c("ebt")})*${ref("taxRate")}`;
      case "ni": return `=${c("ebt")}+${c("tax")}`;
      case "cash": return `=${c("close")}`;
      case "rec": return `=${c("rev")}*${ref("receivablesPct")}`;
      case "inv": return `=${c("rev")}*${ref("inventoryPct")}`;
      case "ppe": return `=${p("ppe")}-${c("capex")}+${c("da")}`;
      case "assets": return `=${c("cash")}+${c("rec")}+${c("inv")}+${c("ppe")}`;
      case "pay": return `=${c("rev")}*${ref("payablesPct")}`;
      case "debt": return `=MAX(0,${p("debt")}-${ref("repayment")})`;
      case "equity": return `=${p("equity")}+${c("ni")}+${c("div")}`;
      case "le": return `=${c("pay")}+${c("debt")}+${c("equity")}`;
      case "cfNi": return `=${c("ni")}`;
      case "cfDa": return `=-${c("da")}`;
      case "cfRec": return `=${p("rec")}-${c("rec")}`;
      case "cfInv": return `=${p("inv")}-${c("inv")}`;
      case "cfPay": return `=${c("pay")}-${p("pay")}`;
      case "cfo": return `=SUM(${c("cfNi")}:${c("cfPay")})`;
      case "capex": return `=-${c("rev")}*${ref("capexPct")}`;
      case "cfi": return `=${c("capex")}`;
      case "cfDebt": return `=${c("debt")}-${p("debt")}`;
      case "div": return `=-MAX(0,${c("ni")})*${ref("payout")}`;
      case "cff": return `=${c("cfDebt")}+${c("div")}`;
      case "net": return `=${c("cfo")}+${c("cfi")}+${c("cff")}`;
      case "open": return `=${p("cash")}`;
      case "close": return `=${c("open")}+${c("net")}`;
      case "check": return `=ROUND(${c("assets")}-${c("le")},6)`;
    }
  };

  rows.push(line("Показатель", Array.from({ length: n + 1 }, (_, t) => text(t === 0 ? `${request.firstYear - 1} факт` : `${request.firstYear + t - 1} прогноз`))));
  for (const item of order) {
    if (typeof item !== "string") { rows.push(line(item.title)); continue; }
    rows.push(line(labels[item], Array.from({ length: n + 1 }, (_, t) => {
      const f = formula(item, t);
      return f ? { formula: f, expected: v[item][t], format: AMOUNT } : blank();
    })));
  }
  const negativeCash = Array.from({ length: n }, (_, i) => i + 1).filter((t) => v.cash[t] < 0).map((t) => `${request.firstYear + t - 1}`);
  return { rows, columns, checkRow: rowNo.check!, negativeCash };
}

/** Расхождения прочитанного с расчётом модели. */
export const modelMismatches = (layout: ModelLayout, values: readonly (readonly unknown[])[]) => rowMismatches(layout.rows, values);

/* --- план ------------------------------------------------------------------------ */

export interface ThreeStatementPlan {
  readonly kind: "build_three_statement_model";
  readonly id: string;
  readonly workbook: { workbookSessionId: string; documentUrl: string };
  readonly request: ModelRequest;
  readonly layout: ModelLayout;
  readonly address: string;
  readonly simplifications: string;
  readonly undoAvailable: boolean;
  readonly undoNote: string;
  readonly createdAt: string;
}

const SIMPLIFICATIONS =
  "Упрощения модели: один темп роста на все годы; дебиторка, запасы и кредиторка — доли выручки; амортизация и капвложения — доли выручки; " +
  "долг гасится равными суммами, проценты — на долг на начало года; налог только с положительной прибыли; проценты на остаток денег не начисляются.";

export async function prepareThreeStatementPlan(args: unknown): Promise<ThreeStatementPlan> {
  preflightToolArgs("build_three_statement_model", args);
  const request = parseModelRequest(args as Record<string, any>);
  const prepared = await Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    let sheet: string;
    try { sheet = checkSheetName(request.sheet, sheets.items.map((item) => item.name)); } catch (error: any) { throw new ToolError(`${error.message} Модель строится на новом листе.`); }
    const layout = buildModel({ ...request, sheet });
    const undo = isCustomUndoAvailable();
    return {
      kind: "build_three_statement_model" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      workbook: currentWorkbookIdentity(),
      request: { ...request, sheet },
      layout,
      address: `A1:${columnLetters(layout.columns)}${layout.rows.length}`,
      simplifications: SIMPLIFICATIONS,
      undoAvailable: undo,
      undoNote: undo ? "Отмена удалит лист модели, если его не меняли после операции." : "Отмена недоступна: монитор изменений Excel не активен.",
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executeThreeStatementPlan(plan: ThreeStatementPlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const { sheetName, values, undoRecorded } = await writeModelSheet(ctx, {
      sheetName: plan.request.sheet,
      rows: plan.layout.rows,
      columns: plan.layout.columns,
      undoAvailable: plan.undoAvailable,
      label: "трёхотчётная модель"
    });
    const mismatches = modelMismatches(plan.layout, values);
    const checks = (values[plan.layout.checkRow - 1] ?? []).slice(1);
    const unbalanced = checks.filter((value) => value !== 0);
    if (mismatches.length || unbalanced.length) {
      throw new ToolExecutionError(
        `Модель записана на лист «${sheetName}», но ${unbalanced.length ? `баланс не сходится (строка ${plan.layout.checkRow}: ${JSON.stringify(checks)})` : ""}` +
          `${unbalanced.length && mismatches.length ? "; " : ""}${mismatches.length ? `значения расходятся с расчётом панели: ${mismatches.slice(0, 8).join("; ")}` : ""}. ` +
          `Готовой она не считается. ${undoRecorded ? "Лист уберёт «Отменить»." : ""}`,
        "applied"
      );
    }
    const n = plan.request.years;
    const lastColumn = n + 1;
    const value = (label: string) => {
      const row = plan.layout.rows.findIndex((cells) => cells[0].expected === label);
      return row === -1 ? null : values[row]?.[lastColumn];
    };
    return {
      ok: true,
      executionState: "verified",
      sheet: sheetName,
      address: plan.address,
      period: `${plan.request.firstYear - 1} факт, ${plan.request.firstYear}–${plan.request.firstYear + n - 1} прогноз`,
      currency: `${plan.request.currency}, ${plan.request.units}`,
      control: { row: plan.layout.checkRow, values: checks, note: "Активы = обязательства + капитал в каждом году, включая начало. Сверено." },
      checkedCells: plan.layout.rows.length * plan.layout.columns,
      lastYear: { revenue: value("Выручка"), netIncome: value("Чистая прибыль"), cash: value("Деньги на конец"), debt: value("Долг") },
      ...(plan.layout.negativeCash.length
        ? { negativeCash: plan.layout.negativeCash, negativeCashNote: "В эти годы деньги на конец отрицательны: модели не хватает финансирования. Это не ошибка расчёта — скажи пользователю, что нужен кредит или другие допущения." }
        : {}),
      simplifications: plan.simplifications,
      note: "Модель — формулы от блока допущений (синий текст): их можно менять, отчёты пересчитаются. Допущения и результат подтверждает человек.",
      undoable: undoRecorded,
      undoNote: plan.undoNote
    };
  });
}
