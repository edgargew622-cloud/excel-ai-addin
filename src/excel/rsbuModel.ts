/**
 * Бухгалтерский баланс (форма 1) и отчёт о финансовых результатах (форма 2)
 * по оборотно-сальдовой ведомости из 1С (24.10.2026). Ядро без Excel и без
 * модели: разбор ОСВ, разнос счетов по строкам по типовой схеме, проверки.
 *
 * Это черновик отчётности по стандартному разносу: строки, которые зависят от
 * аналитики (долгосрочная или краткосрочная задолженность, проценты в 91),
 * разнесены по основному правилу, и об этом сказано в примечаниях. Отчёт для
 * сдачи проверяет бухгалтер.
 */

export type Cell = unknown;

export interface OsvColumns {
  account: number;
  name: number | null;
  startDebit: number; startCredit: number;
  turnDebit: number; turnCredit: number;
  endDebit: number; endCredit: number;
  /** Номер первой строки данных (0 — первая строка области). */
  firstRow: number;
}

export interface OsvRow { row: number; account: string; raw: Cell; name: string; sd: number; sc: number; td: number; tc: number; ed: number; ec: number }

const num = (value: Cell): number => {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value !== "string") return 0;
  const text = value.replace(/[\s ]/g, "").replace(",", ".");
  const n = Number(text);
  return Number.isFinite(n) ? n : 0;
};

/** Номер счёта из ячейки: «01», 1 → «01», «60.01», «10.1» → как есть. */
export function accountOf(value: Cell): string | null {
  if (typeof value === "number" && Number.isInteger(value) && value > 0 && value < 100) return String(value).padStart(2, "0");
  if (typeof value === "number" && value > 0 && value < 100) {
    // 60.1 из числа — субсчёт 60.01 в 1С пишется текстом; число — редкость, берём как есть.
    const [whole, part] = String(value).split(".");
    return `${whole.padStart(2, "0")}.${part}`;
  }
  const text = String(value ?? "").trim();
  const m = /^(\d{2})((?:\.\d{1,3})*)$/.exec(text);
  return m ? m[1] + m[2] : null;
}

/** Столбцы ОСВ по шапке: «Счёт», «Сальдо на начало / Обороты / Сальдо на конец», под ними «Дебет | Кредит». */
export function detectOsv(values: readonly (readonly Cell[])[]): OsvColumns | null {
  const text = (r: number, c: number) => String(values[r]?.[c] ?? "").toLowerCase().replace(/\s+/g, " ").trim();
  const width = Math.max(0, ...values.map((row) => row.length));
  for (let r = 0; r < Math.min(values.length, 30); r++) {
    const accountCol = Array.from({ length: width }, (_, c) => c).find((c) => /^сч[её]т$|^сч[её]т\b|^счет$/.test(text(r, c)));
    if (accountCol === undefined) continue;
    // Группы: слово в строке r (объединённая ячейка — в первом столбце пары), Дт/Кт — в строке r+1 (или в той же).
    const groups: Record<string, number> = {};
    for (let c = 0; c < width; c++) {
      const t = text(r, c);
      if (/сальдо на начало|начальн/.test(t)) groups.start ??= c;
      else if (/оборот/.test(t)) groups.turn ??= c;
      else if (/сальдо на конец|конечн/.test(t)) groups.end ??= c;
    }
    if (groups.start === undefined || groups.turn === undefined || groups.end === undefined) continue;
    const sub = r + 1;
    const pair = (c: number): [number, number] => {
      // Ищем «Дебет» и «Кредит» под группой; нет подписей — берём два столбца подряд.
      for (let k = c; k < Math.min(width, c + 3); k++) {
        if (/дебет|дт/.test(text(sub, k)) && /кредит|кт/.test(text(sub, k + 1))) return [k, k + 1];
      }
      return [c, c + 1];
    };
    const [sd, sc] = pair(groups.start);
    const [td, tc] = pair(groups.turn);
    const [ed, ec] = pair(groups.end);
    const hasSub = /дебет|дт/.test(text(sub, sd));
    const nameCol = Array.from({ length: width }, (_, c) => c).find((c) => c !== accountCol && /наименован|субконто|счет$/.test(text(r, c)));
    return { account: accountCol, name: nameCol ?? null, startDebit: sd, startCredit: sc, turnDebit: td, turnCredit: tc, endDebit: ed, endCredit: ec, firstRow: hasSub ? sub + 1 : r + 1 };
  }
  return null;
}

export function readOsv(values: readonly (readonly Cell[])[], columns: OsvColumns): OsvRow[] {
  const rows: OsvRow[] = [];
  for (let r = columns.firstRow; r < values.length; r++) {
    const line = values[r] ?? [];
    const account = accountOf(line[columns.account]);
    if (!account) continue;
    const nameCell = columns.name !== null ? line[columns.name] : line[columns.account + 1];
    rows.push({
      row: r, account, raw: line[columns.account], name: String(nameCell ?? "").trim(),
      sd: num(line[columns.startDebit]), sc: num(line[columns.startCredit]),
      td: num(line[columns.turnDebit]), tc: num(line[columns.turnCredit]),
      ed: num(line[columns.endDebit]), ec: num(line[columns.endCredit])
    });
  }
  return rows;
}

/**
 * Как взять счёт из ОСВ: строкой самого счёта («62»), а если её нет — суммой
 * субсчетов следующего уровня («62.01», «62.02»), без двойного счёта иерархии.
 */
export interface AccountPick { account: string; rows: OsvRow[]; criteria: Cell[] }

export function pickAccount(rows: readonly OsvRow[], account: string): AccountPick {
  const exact = rows.filter((row) => row.account === account);
  if (exact.length) return { account, rows: exact.slice(0, 1), criteria: [exact[0].raw] };
  const depth = account.split(".").length;
  const children = rows.filter((row) => row.account.startsWith(account + ".") && row.account.split(".").length === depth + 1);
  if (children.length) return { account, rows: children, criteria: children.map((row) => row.raw) };
  // Только более глубокие уровни (90.01.1) — по ближайшему уровню, что есть.
  const deeper = rows.filter((row) => row.account.startsWith(account + "."));
  const minDepth = Math.min(...deeper.map((row) => row.account.split(".").length));
  const level = deeper.filter((row) => row.account.split(".").length === minDepth);
  return { account, rows: level, criteria: level.map((row) => row.raw) };
}

export type Field = "sd" | "sc" | "td" | "tc" | "ed" | "ec";

/** Слагаемое строки: счёт, поле ведомости, знак. */
export interface Term { account: string; field: Field; sign: 1 | -1 }

export interface FormLine {
  code: string;
  name: string;
  /** Слагаемые для «на отчётную дату» (сальдо на конец / обороты). Пусто — итоговая строка. */
  terms: Term[];
  /** Итоговая строка: сумма других строк (со знаком). */
  sumOf?: { code: string; sign: 1 | -1 }[];
  bold?: boolean;
  section?: string;
  note?: string;
}

/** Слагаемые на начало периода: поля сальдо на начало вместо сальдо на конец. */
const START: Record<Field, Field> = { ed: "sd", ec: "sc", sd: "sd", sc: "sc", td: "td", tc: "tc" };
export const toStart = (terms: Term[]): Term[] => terms.map((term) => ({ ...term, field: START[term.field] }));

const D = (account: string, sign: 1 | -1 = 1): Term => ({ account, field: "ed", sign });
const K = (account: string, sign: 1 | -1 = 1): Term => ({ account, field: "ec", sign });
/** Сальдо по счёту, как оно стоит в балансе: дебет − кредит (актив) или кредит − дебет (пассив). */
const KD = (account: string): Term[] => [K(account), D(account, -1)];

const SETTLEMENT = ["60", "62", "68", "69", "70", "71", "73", "75", "76"];

/** Форма 1: типовой разнос по плану счетов (приказ Минфина № 94н, форма — приказ № 66н). */
export const BALANCE: FormLine[] = [
  { code: "", name: "АКТИВ", terms: [], section: "head" },
  { code: "", name: "I. ВНЕОБОРОТНЫЕ АКТИВЫ", terms: [], section: "head" },
  { code: "1110", name: "Нематериальные активы", terms: [D("04"), K("05", -1)] },
  { code: "1150", name: "Основные средства", terms: [D("01"), D("03"), K("02", -1)], note: "01 + 03 − 02" },
  { code: "1170", name: "Финансовые вложения", terms: [], note: "долгосрочные вложения из 58 по аналитике — вручную; здесь всё 58 в строке 1240" },
  { code: "1180", name: "Отложенные налоговые активы", terms: [D("09")] },
  { code: "1190", name: "Прочие внеоборотные активы", terms: [D("08")], note: "08 «Вложения во внеоборотные активы»" },
  { code: "1100", name: "Итого по разделу I", terms: [], sumOf: ["1110", "1150", "1170", "1180", "1190"].map((code) => ({ code, sign: 1 as const })), bold: true },
  { code: "", name: "II. ОБОРОТНЫЕ АКТИВЫ", terms: [], section: "head" },
  { code: "1210", name: "Запасы", terms: [...["10", "11", "15", "16", "20", "21", "23", "29", "41", "43", "44", "45", "97"].map((a) => D(a)), K("14", -1), K("42", -1)], note: "10, 11, 15, 16, 20, 21, 23, 29, 41, 43, 44, 45, 97 − 14 − 42" },
  { code: "1220", name: "Налог на добавленную стоимость по приобретённым ценностям", terms: [D("19")] },
  { code: "1230", name: "Дебиторская задолженность", terms: [...SETTLEMENT.map((a) => D(a)), K("63", -1)], note: "дебетовые сальдо 60, 62, 68, 69, 70, 71, 73, 75, 76 − 63" },
  { code: "1240", name: "Финансовые вложения (за исключением денежных эквивалентов)", terms: [D("58"), K("59", -1)] },
  { code: "1250", name: "Денежные средства и денежные эквиваленты", terms: ["50", "51", "52", "55", "57"].map((a) => D(a)) },
  { code: "1260", name: "Прочие оборотные активы", terms: [D("94")] },
  { code: "1200", name: "Итого по разделу II", terms: [], sumOf: ["1210", "1220", "1230", "1240", "1250", "1260"].map((code) => ({ code, sign: 1 as const })), bold: true },
  { code: "1600", name: "БАЛАНС", terms: [], sumOf: [{ code: "1100", sign: 1 }, { code: "1200", sign: 1 }], bold: true },
  { code: "", name: "ПАССИВ", terms: [], section: "head" },
  { code: "", name: "III. КАПИТАЛ И РЕЗЕРВЫ", terms: [], section: "head" },
  { code: "1310", name: "Уставный капитал (складочный капитал, уставный фонд, вклады товарищей)", terms: KD("80") },
  { code: "1320", name: "Собственные акции, выкупленные у акционеров", terms: [D("81", -1)], note: "со знаком минус" },
  { code: "1350", name: "Добавочный капитал (без переоценки)", terms: KD("83") },
  { code: "1360", name: "Резервный капитал", terms: KD("82") },
  { code: "1370", name: "Нераспределённая прибыль (непокрытый убыток)", terms: [...KD("84"), ...KD("99")], note: "84 + 99 (прибыль текущего года до реформации)" },
  { code: "1300", name: "Итого по разделу III", terms: [], sumOf: ["1310", "1320", "1350", "1360", "1370"].map((code) => ({ code, sign: 1 as const })), bold: true },
  { code: "", name: "IV. ДОЛГОСРОЧНЫЕ ОБЯЗАТЕЛЬСТВА", terms: [], section: "head" },
  { code: "1410", name: "Заёмные средства", terms: [K("67")] },
  { code: "1420", name: "Отложенные налоговые обязательства", terms: [K("77")] },
  { code: "1400", name: "Итого по разделу IV", terms: [], sumOf: ["1410", "1420"].map((code) => ({ code, sign: 1 as const })), bold: true },
  { code: "", name: "V. КРАТКОСРОЧНЫЕ ОБЯЗАТЕЛЬСТВА", terms: [], section: "head" },
  { code: "1510", name: "Заёмные средства", terms: [K("66")] },
  { code: "1520", name: "Кредиторская задолженность", terms: SETTLEMENT.map((a) => K(a)), note: "кредитовые сальдо 60, 62, 68, 69, 70, 71, 73, 75, 76" },
  { code: "1530", name: "Доходы будущих периодов", terms: [K("98")] },
  { code: "1540", name: "Оценочные обязательства", terms: [K("96")] },
  { code: "1500", name: "Итого по разделу V", terms: [], sumOf: ["1510", "1520", "1530", "1540"].map((code) => ({ code, sign: 1 as const })), bold: true },
  { code: "1700", name: "БАЛАНС", terms: [], sumOf: [{ code: "1300", sign: 1 }, { code: "1400", sign: 1 }, { code: "1500", sign: 1 }], bold: true }
];

const T = (account: string, field: "td" | "tc", sign: 1 | -1 = 1): Term => ({ account, field, sign });

/** Форма 2: по оборотам субсчетов 90, 91, 99 (план счетов 1С). */
export const RESULTS: FormLine[] = [
  { code: "2110", name: "Выручка", terms: [T("90.01", "tc"), T("90.03", "td", -1), T("90.04", "td", -1)], note: "Кт 90.01 − Дт 90.03 (НДС) − Дт 90.04 (акцизы)" },
  { code: "2120", name: "Себестоимость продаж", terms: [T("90.02", "td")], note: "Дт 90.02; в форме — в скобках" },
  { code: "2100", name: "Валовая прибыль (убыток)", terms: [], sumOf: [{ code: "2110", sign: 1 }, { code: "2120", sign: -1 }], bold: true },
  { code: "2210", name: "Коммерческие расходы", terms: [T("90.07", "td")] },
  { code: "2220", name: "Управленческие расходы", terms: [T("90.08", "td")] },
  { code: "2200", name: "Прибыль (убыток) от продаж", terms: [], sumOf: [{ code: "2100", sign: 1 }, { code: "2210", sign: -1 }, { code: "2220", sign: -1 }], bold: true },
  { code: "2340", name: "Прочие доходы", terms: [T("91.01", "tc")], note: "Кт 91.01; проценты и доходы от участия — по аналитике, вручную" },
  { code: "2350", name: "Прочие расходы", terms: [T("91.02", "td")], note: "Дт 91.02; проценты к уплате — по аналитике, вручную" },
  { code: "2300", name: "Прибыль (убыток) до налогообложения", terms: [], sumOf: [{ code: "2200", sign: 1 }, { code: "2340", sign: 1 }, { code: "2350", sign: -1 }], bold: true },
  { code: "2410", name: "Налог на прибыль", terms: [T("99.02", "td"), T("99.02", "tc", -1)], note: "Дт − Кт 99.02 (текущий и отложенный налог)" },
  { code: "2400", name: "Чистая прибыль (убыток)", terms: [], sumOf: [{ code: "2300", sign: 1 }, { code: "2410", sign: -1 }], bold: true }
];

export interface ComputedLine { line: FormLine; now: number; before: number | null; used: string[] }

/** Расчёт строк: значения и какие счета реально нашлись. */
export function computeForm(form: readonly FormLine[], rows: readonly OsvRow[], withStart: boolean): ComputedLine[] {
  const value = (terms: Term[]) => {
    let total = 0;
    const used: string[] = [];
    for (const term of terms) {
      const pick = pickAccount(rows, term.account);
      if (!pick.rows.length) continue;
      const amount = pick.rows.reduce((sum, row) => sum + row[term.field], 0);
      if (amount) used.push(term.account);
      total += term.sign * amount;
    }
    return { total: Math.round(total * 100) / 100, used: [...new Set(used)] };
  };
  const out: ComputedLine[] = [];
  const byCode = new Map<string, ComputedLine>();
  for (const line of form) {
    let now = 0; let before: number | null = withStart ? 0 : null; let used: string[] = [];
    if (line.sumOf) {
      for (const part of line.sumOf) {
        const item = byCode.get(part.code);
        if (!item) continue;
        now += part.sign * item.now;
        if (before !== null) before += part.sign * (item.before ?? 0);
      }
    } else if (line.terms.length) {
      const a = value(line.terms);
      now = a.total; used = a.used;
      if (withStart) before = value(toStart(line.terms)).total;
    }
    const item = { line, now: Math.round(now * 100) / 100, before: before === null ? null : Math.round(before * 100) / 100, used };
    out.push(item);
    if (line.code) byCode.set(line.code, item);
  }
  return out;
}

/** Счета с остатками или оборотами, которые не попали ни в одну строку. */
export function unmappedAccounts(rows: readonly OsvRow[], forms: readonly (readonly FormLine[])[]): string[] {
  const mapped = new Set(forms.flatMap((form) => form.flatMap((line) => line.terms.map((term) => term.account.split(".")[0]))));
  // Служебные: закрываются внутри периода и в баланс не идут.
  for (const service of ["90", "91", "99", "25", "26", "28", "40", "46", "00", "000"]) mapped.add(service);
  const out = new Set<string>();
  for (const row of rows) {
    const top = row.account.split(".")[0];
    if (mapped.has(top)) continue;
    if (row.ed || row.ec || row.sd || row.sc) out.add(top);
  }
  return [...out].sort();
}

/** Сводка ОСВ: итоги дебета и кредита по счетам верхнего уровня — ведомость сама сходится? */
export function osvTotals(rows: readonly OsvRow[]) {
  const top = new Map<string, OsvRow>();
  for (const row of rows) {
    const key = row.account.split(".")[0];
    const pick = pickAccount(rows, key);
    if (!top.has(key) && pick.rows.length) top.set(key, { ...pick.rows[0], ed: pick.rows.reduce((s, r) => s + r.ed, 0), ec: pick.rows.reduce((s, r) => s + r.ec, 0), sd: pick.rows.reduce((s, r) => s + r.sd, 0), sc: pick.rows.reduce((s, r) => s + r.sc, 0) });
  }
  const sum = (field: Field) => Math.round([...top.values()].reduce((s, r) => s + r[field], 0) * 100) / 100;
  return { endDebit: sum("ed"), endCredit: sum("ec"), startDebit: sum("sd"), startCredit: sum("sc") };
}
