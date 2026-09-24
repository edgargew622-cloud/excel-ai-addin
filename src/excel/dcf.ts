/**
 * Оценка DCF (этап 7, 7.5.2): свободный денежный поток по годам, остаточная
 * стоимость по Гордону, стоимость бизнеса и капитала, таблица
 * чувствительности к WACC и росту после прогноза — на новом листе формулами.
 *
 * Как и трёхотчётная модель: все допущения обязательны и стоят в ячейках,
 * лист хранит валюту, единицы, период и источник, каждая ячейка после записи
 * сверяется с расчётом панели. Контрольное равенство — центр таблицы
 * чувствительности совпадает со стоимостью бизнеса: таблица считает ту же
 * величину другой формулой, и расхождение выдало бы ошибку в ссылке.
 *
 * Панель не судит, верна ли оценка. Доля остаточной стоимости и множитель
 * EV/EBITDA называются, чтобы человек видел, на чём держится результат.
 */

import { assertPlanWorkbook, deepFreeze, preflightToolArgs, ToolError, ToolExecutionError } from "./excelTools";
import { columnLetters } from "./formulaFill";
import { AMOUNT, line as makeLine, modelMismatches, SHARE, text, writeModelSheet, type ModelCell } from "./modelSheet";
import { checkSheetName } from "./sheetRules";
import { isCustomUndoAvailable } from "./undo";
import { currentWorkbookIdentity } from "./workbookContext";

export interface DcfAssumptions {
  revenue0: number;
  growth: number;
  ebitMargin: number;
  taxRate: number;
  daPct: number;
  capexPct: number;
  nwcPct: number;
  wacc: number;
  terminalGrowth: number;
  netDebt: number;
  shares?: number;
}

type Key = keyof DcfAssumptions;

export const DCF_INPUTS: readonly { key: Key; label: string; share: boolean; optional?: true }[] = [
  { key: "revenue0", label: "Выручка последнего фактического года", share: false },
  { key: "growth", label: "Рост выручки в год", share: true },
  { key: "ebitMargin", label: "Операционная маржа (EBIT / выручка)", share: true },
  { key: "taxRate", label: "Ставка налога", share: true },
  { key: "daPct", label: "Амортизация, доля выручки", share: true },
  { key: "capexPct", label: "Капвложения, доля выручки", share: true },
  { key: "nwcPct", label: "Оборотный капитал, доля выручки", share: true },
  { key: "wacc", label: "Ставка дисконтирования (WACC)", share: true },
  { key: "terminalGrowth", label: "Рост после прогноза", share: true },
  { key: "netDebt", label: "Чистый долг", share: false },
  { key: "shares", label: "Число акций", share: false, optional: true }
];

export interface DcfRequest {
  sheet: string;
  currency: string;
  units: string;
  source: string;
  firstYear: number;
  years: number;
  assumptions: DcfAssumptions;
}

export function parseDcfRequest(args: Record<string, any>): DcfRequest {
  const raw = (args.assumptions ?? {}) as Record<string, unknown>;
  const missing = DCF_INPUTS.filter((input) => !input.optional && (typeof raw[input.key] !== "number" || !Number.isFinite(raw[input.key] as number)));
  const meta: string[] = [];
  if (typeof args.currency !== "string" || !args.currency.trim()) meta.push("валюта (currency)");
  if (typeof args.units !== "string" || !args.units.trim()) meta.push("единицы (units)");
  if (typeof args.source !== "string" || !args.source.trim()) meta.push("источник допущений (source)");
  if (missing.length || meta.length) {
    throw new ToolError(
      `Оценка не строится, пока не заданы все допущения. Не хватает: ${[...meta, ...missing.map((input) => `${input.label} (${input.key})`)].join("; ")}. ` +
        "Спроси их у пользователя — значения по умолчанию не подставляются."
    );
  }
  const a = raw as unknown as DcfAssumptions;
  const ranged = (["ebitMargin", "taxRate", "daPct", "capexPct", "nwcPct", "wacc"] as Key[]).filter((key) => (a[key] as number) < 0 || (a[key] as number) > 1);
  if (ranged.length) throw new ToolError(`Доли задаются от 0 до 1 (0,1 — это 10 %): ${ranged.map((key) => `${key} = ${a[key]}`).join(", ")}.`);
  if (a.growth <= -1) throw new ToolError("Рост выручки должен быть больше −100 % (growth > −1).");
  if (a.revenue0 <= 0) throw new ToolError("Выручка последнего фактического года должна быть положительной.");
  if (a.terminalGrowth >= a.wacc) {
    throw new ToolError(`Рост после прогноза (${a.terminalGrowth}) должен быть меньше WACC (${a.wacc}): иначе остаточная стоимость бесконечна или отрицательна. Уточни у пользователя.`);
  }
  if (a.shares !== undefined && !(typeof a.shares === "number" && a.shares > 0)) throw new ToolError("Число акций, если задано, — положительное число.");
  const years = Number(args.years);
  const firstYear = Number(args.firstYear);
  if (!Number.isInteger(years) || years < 3 || years > 10) throw new ToolError("years — целое от 3 до 10: прогноз короче трёх лет для DCF не годится.");
  if (!Number.isInteger(firstYear) || firstYear < 1900 || firstYear > 2200) throw new ToolError("firstYear — первый прогнозный год, например 2026.");
  return { sheet: String(args.sheet ?? "").trim(), currency: args.currency.trim(), units: args.units.trim(), source: args.source.trim(), firstYear, years, assumptions: a };
}

/* --- раскладка ------------------------------------------------------------------ */

export interface DcfLayout {
  rows: ModelCell[][];
  columns: number;
  /** Номер строки листа с контролем «центр чувствительности − EV». */
  checkRow: number;
  ev: number;
  equity: number;
  perShare: number | null;
  tvShare: number;
  evEbitda: number | null;
  negativeFcf: string[];
}

type RowKey = "rev" | "ebit" | "tax" | "da" | "capex" | "nwc" | "dnwc" | "fcf" | "t" | "df" | "pv";

const SENSITIVITY_WACC = [-0.01, 0, 0.01];
const SENSITIVITY_GROWTH = [-0.005, 0, 0.005];

export function buildDcf(request: DcfRequest): DcfLayout {
  const a = request.assumptions;
  const n = request.years;
  const columns = n + 2;
  const rows: ModelCell[][] = [];
  const line = (label: string, cells: ModelCell[] = []) => makeLine(columns, label, cells);
  const blank = (): ModelCell => ({ formula: "", expected: "" });
  const amount = (formula: string, expected: number, format = AMOUNT): ModelCell => ({ formula, expected, format });

  rows.push(line("Оценка DCF"));
  rows.push(line("Валюта и единицы", [text(`${request.currency}, ${request.units}`)]));
  rows.push(line("Период", [text(`${request.firstYear - 1} — факт, ${request.firstYear}–${request.firstYear + n - 1} — прогноз; дисконтирование на конец года`)]));
  rows.push(line("Источник допущений", [text(request.source)]));
  rows.push(line(""));
  rows.push(line("Допущения", [text("значение")]));
  const inputRow: Partial<Record<Key, number>> = {};
  for (const input of DCF_INPUTS) {
    if (input.optional && a[input.key] === undefined) continue;
    inputRow[input.key] = rows.length + 1;
    rows.push(line(input.label, [{ formula: a[input.key] as number, expected: a[input.key] as number, format: input.share ? SHARE : AMOUNT, input: true }]));
  }
  const ref = (key: Key) => `$B$${inputRow[key]}`;
  rows.push(line(""));

  const order: RowKey[] = ["rev", "ebit", "tax", "da", "capex", "nwc", "dnwc", "fcf", "t", "df", "pv"];
  const headerRow = rows.length + 1;
  const rowNo = Object.fromEntries(order.map((key, index) => [key, headerRow + 1 + index])) as Record<RowKey, number>;
  const col = (t: number) => columnLetters(2 + t);
  const at = (key: RowKey, t: number) => `${col(t)}${rowNo[key]}`;

  const v = Object.fromEntries(order.map((key) => [key, new Array(n + 1).fill(0)])) as Record<RowKey, number[]>;
  v.rev[0] = a.revenue0;
  v.nwc[0] = a.revenue0 * a.nwcPct;
  for (let t = 1; t <= n; t++) {
    v.rev[t] = v.rev[t - 1] * (1 + a.growth);
    v.ebit[t] = v.rev[t] * a.ebitMargin;
    v.tax[t] = -Math.max(0, v.ebit[t]) * a.taxRate;
    v.da[t] = v.rev[t] * a.daPct;
    v.capex[t] = -v.rev[t] * a.capexPct;
    v.nwc[t] = v.rev[t] * a.nwcPct;
    v.dnwc[t] = -(v.nwc[t] - v.nwc[t - 1]);
    v.fcf[t] = v.ebit[t] + v.tax[t] + v.da[t] + v.capex[t] + v.dnwc[t];
    v.t[t] = t;
    v.df[t] = 1 / (1 + a.wacc) ** t;
    v.pv[t] = v.fcf[t] * v.df[t];
  }
  const labels: Record<RowKey, string> = {
    rev: "Выручка", ebit: "Операционная прибыль (EBIT)", tax: "Налог на EBIT", da: "Амортизация", capex: "Капвложения",
    nwc: "Оборотный капитал", dnwc: "Изменение оборотного капитала", fcf: "Свободный денежный поток", t: "Год дисконтирования",
    df: "Коэффициент дисконтирования", pv: "Приведённый поток"
  };
  const formula = (key: RowKey, t: number): string => {
    const c = (k: RowKey) => at(k, t);
    if (t === 0) return key === "rev" ? `=${ref("revenue0")}` : key === "nwc" ? `=${c("rev")}*${ref("nwcPct")}` : "";
    switch (key) {
      case "rev": return `=${at("rev", t - 1)}*(1+${ref("growth")})`;
      case "ebit": return `=${c("rev")}*${ref("ebitMargin")}`;
      case "tax": return `=-MAX(0,${c("ebit")})*${ref("taxRate")}`;
      case "da": return `=${c("rev")}*${ref("daPct")}`;
      case "capex": return `=-${c("rev")}*${ref("capexPct")}`;
      case "nwc": return `=${c("rev")}*${ref("nwcPct")}`;
      case "dnwc": return `=-(${c("nwc")}-${at("nwc", t - 1)})`;
      case "fcf": return `=${c("ebit")}+${c("tax")}+${c("da")}+${c("capex")}+${c("dnwc")}`;
      case "t": return "";
      case "df": return `=1/(1+${ref("wacc")})^${c("t")}`;
      case "pv": return `=${c("fcf")}*${c("df")}`;
    }
  };
  rows.push(line("Показатель", Array.from({ length: n + 1 }, (_, t) => text(t === 0 ? `${request.firstYear - 1} факт` : `${request.firstYear + t - 1} прогноз`))));
  for (const key of order) {
    rows.push(line(labels[key], Array.from({ length: n + 1 }, (_, t) => {
      if (key === "t") return t === 0 ? blank() : { formula: t, expected: t, format: "0" };
      const f = formula(key, t);
      return f ? amount(f, v[key][t], key === "df" ? "0.0000" : AMOUNT) : blank();
    })));
  }
  rows.push(line(""));

  // Оценка: значения в столбце B.
  const last = col(n);
  const sumPv = v.pv.slice(1).reduce((sum, value) => sum + value, 0);
  const tv = v.fcf[n] * (1 + a.terminalGrowth) / (a.wacc - a.terminalGrowth);
  const pvTv = tv * v.df[n];
  const ev = sumPv + pvTv;
  const equity = ev - a.netDebt;
  const valueRow: Record<string, number> = {};
  const valueLine = (key: string, label: string, cell: ModelCell) => { valueRow[key] = rows.length + 1; rows.push(line(label, [cell])); };
  rows.push(line("Оценка"));
  valueLine("sumPv", "Сумма приведённых потоков", amount(`=SUM(${col(1)}${rowNo.pv}:${last}${rowNo.pv})`, sumPv));
  valueLine("tv", `Остаточная стоимость на конец ${request.firstYear + n - 1}`, amount(`=${last}${rowNo.fcf}*(1+${ref("terminalGrowth")})/(${ref("wacc")}-${ref("terminalGrowth")})`, tv));
  valueLine("pvTv", "Приведённая остаточная стоимость", amount(`=B${valueRow.tv}*${last}${rowNo.df}`, pvTv));
  valueLine("ev", "Стоимость бизнеса (EV)", amount(`=B${valueRow.sumPv}+B${valueRow.pvTv}`, ev));
  valueLine("netDebt", "Чистый долг", amount(`=${ref("netDebt")}`, a.netDebt));
  valueLine("equity", "Стоимость капитала", amount(`=B${valueRow.ev}-B${valueRow.netDebt}`, equity));
  let perShare: number | null = null;
  if (a.shares !== undefined) {
    perShare = equity / a.shares;
    valueLine("perShare", "Стоимость одной акции", amount(`=B${valueRow.equity}/${ref("shares")}`, perShare, "#,##0.00"));
  }
  const tvShare = pvTv / ev;
  const ebitdaLast = v.ebit[n] + v.da[n];
  const evEbitda = ebitdaLast > 0 ? ev / ebitdaLast : null;
  valueLine("tvShare", "Доля остаточной стоимости в EV", amount(`=B${valueRow.pvTv}/B${valueRow.ev}`, tvShare, SHARE));
  const ebitdaCells = `(${last}${rowNo.ebit}+${last}${rowNo.da})`;
  valueLine("evEbitda", `EV / EBITDA ${request.firstYear + n - 1}`, { formula: `=IF(${ebitdaCells}<=0,"",B${valueRow.ev}/${ebitdaCells})`, expected: evEbitda ?? "", format: "0.0" });
  rows.push(line(""));

  // Чувствительность: строки — WACC, столбцы — рост после прогноза.
  rows.push(line("Чувствительность EV: WACC по строкам, рост после прогноза по столбцам"));
  const gRow = rows.length + 1;
  rows.push(line("WACC \\ рост", SENSITIVITY_GROWTH.map((dg, j) => amount(`=${ref("terminalGrowth")}${dg ? (dg > 0 ? `+${dg}` : `${dg}`) : ""}`, a.terminalGrowth + dg, SHARE))));
  let center = 0;
  SENSITIVITY_WACC.forEach((dw, i) => {
    const w = a.wacc + dw;
    const wRef = `$A${rows.length + 1}`;
    const cells = SENSITIVITY_GROWTH.map((dg, j) => {
      const g = a.terminalGrowth + dg;
      const gRef = `${columnLetters(2 + j)}$${gRow}`;
      const terms = Array.from({ length: n }, (_, k) => `${col(k + 1)}$${rowNo.fcf}/(1+${wRef})^${k + 1}`).join("+");
      const value = w <= g ? "" : v.fcf.slice(1).reduce((sum, fcf, k) => sum + fcf / (1 + w) ** (k + 1), 0) + v.fcf[n] * (1 + g) / (w - g) / (1 + w) ** n;
      if (i === 1 && j === 1) center = value as number;
      return {
        formula: `=IF(${wRef}<=${gRef},"",${terms}+${last}$${rowNo.fcf}*(1+${gRef})/(${wRef}-${gRef})/(1+${wRef})^${n})`,
        expected: value,
        format: AMOUNT
      } as ModelCell;
    });
    rows.push(line("", cells));
    rows[rows.length - 1][0] = { formula: `=${ref("wacc")}${dw ? (dw > 0 ? `+${dw}` : `${dw}`) : ""}`, expected: w, format: SHARE };
  });
  const centerCell = `C${gRow + 2}`;
  const checkRow = rows.length + 1;
  rows.push(line("Контроль: центр таблицы − EV (должно быть 0)", [amount(`=ROUND(${centerCell}-B${valueRow.ev},6)`, Math.round((center - ev) * 1e6) / 1e6 || 0, "0.000000")]));

  const negativeFcf = Array.from({ length: n }, (_, k) => k + 1).filter((t) => v.fcf[t] < 0).map((t) => `${request.firstYear + t - 1}`);
  return { rows, columns, checkRow, ev, equity, perShare, tvShare, evEbitda, negativeFcf };
}

/* --- план ------------------------------------------------------------------------ */

export interface DcfPlan {
  readonly kind: "build_dcf_model";
  readonly id: string;
  readonly workbook: { workbookSessionId: string; documentUrl: string };
  readonly request: DcfRequest;
  readonly layout: DcfLayout;
  readonly simplifications: string;
  readonly undoAvailable: boolean;
  readonly undoNote: string;
  readonly createdAt: string;
}

const SIMPLIFICATIONS =
  "Упрощения оценки: один темп роста и одна маржа на весь прогноз; амортизация, капвложения и оборотный капитал — доли выручки; " +
  "налог на положительный EBIT; дисконтирование на конец года; остаточная стоимость по Гордону от потока последнего года; WACC постоянен.";

export async function prepareDcfPlan(args: unknown): Promise<DcfPlan> {
  preflightToolArgs("build_dcf_model", args);
  const request = parseDcfRequest(args as Record<string, any>);
  const prepared = await Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    let sheet: string;
    try { sheet = checkSheetName(request.sheet, sheets.items.map((item) => item.name)); } catch (error: any) { throw new ToolError(`${error.message} Оценка строится на новом листе.`); }
    const layout = buildDcf({ ...request, sheet });
    const undo = isCustomUndoAvailable();
    return {
      kind: "build_dcf_model" as const,
      id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      workbook: currentWorkbookIdentity(),
      request: { ...request, sheet },
      layout,
      simplifications: SIMPLIFICATIONS,
      undoAvailable: undo,
      undoNote: undo ? "Отмена удалит лист оценки, если его не меняли после операции." : "Отмена недоступна: монитор изменений Excel не активен.",
      createdAt: new Date().toISOString()
    };
  });
  return deepFreeze(prepared);
}

export async function executeDcfPlan(plan: DcfPlan) {
  assertPlanWorkbook(plan);
  return Excel.run(async (ctx) => {
    const { sheetName, values, undoRecorded } = await writeModelSheet(ctx, {
      sheetName: plan.request.sheet,
      rows: plan.layout.rows,
      columns: plan.layout.columns,
      undoAvailable: plan.undoAvailable,
      label: "оценка DCF"
    });
    const mismatches = modelMismatches(plan.layout.rows, values);
    const check = values[plan.layout.checkRow - 1]?.[1];
    if (mismatches.length || check !== 0) {
      throw new ToolExecutionError(
        `Оценка записана на лист «${sheetName}», но ${check !== 0 ? `центр таблицы чувствительности расходится со стоимостью бизнеса (строка ${plan.layout.checkRow}: ${String(check)})` : ""}` +
          `${check !== 0 && mismatches.length ? "; " : ""}${mismatches.length ? `значения расходятся с расчётом панели: ${mismatches.slice(0, 8).join("; ")}` : ""}. ` +
          `Готовой она не считается. ${undoRecorded ? "Лист уберёт «Отменить»." : ""}`,
        "applied"
      );
    }
    const { layout, request } = plan;
    return {
      ok: true,
      executionState: "verified",
      sheet: sheetName,
      period: `${request.firstYear - 1} факт, ${request.firstYear}–${request.firstYear + request.years - 1} прогноз`,
      currency: `${request.currency}, ${request.units}`,
      enterpriseValue: layout.ev,
      equityValue: layout.equity,
      ...(layout.perShare !== null ? { perShare: layout.perShare } : {}),
      terminalValueShare: layout.tvShare,
      evToEbitda: layout.evEbitda,
      control: { row: layout.checkRow, value: check, note: "Центр таблицы чувствительности совпадает со стоимостью бизнеса. Сверено." },
      checkedCells: layout.rows.length * layout.columns,
      ...(layout.tvShare > 0.75
        ? { terminalNote: `Остаточная стоимость — ${(layout.tvShare * 100).toFixed(0)} % оценки: результат держится на допущении о росте после прогноза и WACC. Скажи это пользователю.` }
        : {}),
      ...(layout.negativeFcf.length ? { negativeFcf: layout.negativeFcf, negativeFcfNote: "В эти годы свободный поток отрицателен." } : {}),
      simplifications: plan.simplifications,
      note: "Оценка — формулы от блока допущений (синий текст). Панель проверила расчёт, но не допущения: верность оценки подтверждает человек.",
      undoable: undoRecorded,
      undoNote: plan.undoNote
    };
  });
}
