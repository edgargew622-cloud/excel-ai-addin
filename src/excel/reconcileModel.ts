/**
 * Сверка двух таблиц — ядро (этап 11, PLAN_11_RECONCILE.md).
 *
 * Сопоставляет строки двух таблиц без Excel и без модели: одинаковые данные
 * дают одинаковый результат. Модель только выбирает столбцы и допуски.
 *
 * Поля строки: дата, сумма (или дебет/кредит), названия контрагентов, ИНН,
 * номера документов и тексты назначения. Правила: сумма — точно или с
 * допуском, дата — ± дни, название — похожесть, ИНН и номер счёта — точно.
 *
 * Разделы результата: «Совпало» (все правила выполнены), «Вероятно» (не всё,
 * с оценкой и причиной — не доказано, пока пользователь не подтвердит),
 * «Расхождение суммы», «Комиссии банка», «Только слева», «Только справа».
 * Каждая строка обеих таблиц попадает ровно в один раздел — это проверяется.
 */

/* ------------------------------------------------------------ разбор */

export type Cell = unknown;

/** Число из ячейки: 1234.5, «1 234,50», «-1 234,50 руб.», «(1 234)». */
export function toNumber(value: Cell): number | null {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value !== "string") return null;
  let text = value.trim().replace(/[  \s]/g, "").replace(/(руб\.?|р\.|₽|rub)$/i, "");
  if (!text) return null;
  let negative = false;
  if (/^\(.*\)$/.test(text)) { negative = true; text = text.slice(1, -1); }
  if (text.startsWith("-")) { negative = !negative; text = text.slice(1); }
  if (!/^[0-9.,]+$/.test(text)) return null;
  const lastComma = text.lastIndexOf(",");
  const lastDot = text.lastIndexOf(".");
  // Последний из «,» и «.» — десятичный, если после него 1–2 цифры; остальные — разделители групп.
  const decimalAt = Math.max(lastComma, lastDot);
  let integer = text;
  let fraction = "";
  if (decimalAt >= 0 && text.length - decimalAt - 1 <= 2 && text.length - decimalAt - 1 >= 1) {
    integer = text.slice(0, decimalAt);
    fraction = text.slice(decimalAt + 1);
  }
  integer = integer.replace(/[.,]/g, "");
  if (!/^\d+$/.test(integer || "0") || (fraction && !/^\d+$/.test(fraction))) return null;
  const number = Number(`${integer || "0"}.${fraction || "0"}`);
  return negative ? -number : number;
}

/** Дата → номер дня Excel: число Excel, «dd.mm.yyyy», «yyyy-mm-dd», «dd/mm/yy». */
export function toDay(value: Cell): number | null {
  if (typeof value === "number" && Number.isFinite(value) && value > 20000 && value < 80000) return Math.floor(value);
  if (typeof value !== "string") return null;
  const text = value.trim();
  let m = /^(\d{1,2})[./-](\d{1,2})[./-](\d{2}|\d{4})(?:\s.*)?$/.exec(text);
  let year: number; let month: number; let day: number;
  if (m) { day = +m[1]; month = +m[2]; year = +m[3]; if (year < 100) year += 2000; }
  else {
    m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/.exec(text);
    if (!m) return null;
    year = +m[1]; month = +m[2]; day = +m[3];
  }
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  const utc = Date.UTC(year, month - 1, day);
  const check = new Date(utc);
  if (check.getUTCMonth() !== month - 1) return null;
  return Math.round(utc / 86_400_000) + 25569;
}

/** Номер дня Excel → «dd.mm.yyyy». */
export function dayText(day: number): string {
  const d = new Date((day - 25569) * 86_400_000);
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${pad(d.getUTCDate())}.${pad(d.getUTCMonth() + 1)}.${d.getUTCFullYear()}`;
}

/** ОПФ, которые не отличают контрагента. Сравниваются целыми словами после очистки знаков. */
const LEGAL_FORMS = new Set(["ооо", "оао", "зао", "пао", "ао", "нао", "ип", "чп", "нко", "ано", "гуп", "муп", "фгуп", "тоо", "llc", "ltd", "inc", "gmbh"]);
// В JavaScript \b не видит кириллицу: прежнее правило с \b «ООО» не снимало
// («Книга11», 08.10.2026 — ключ «ооо полюс»). Длинные формы — заменой фразы.
const LONG_FORMS = /индивидуальный предприниматель|общество с ограниченной ответственностью|публичное акционерное общество|акционерное общество/g;

/** Название для сравнения: без ОПФ, кавычек, знаков, регистра; слова по алфавиту. */
export function normalizeName(value: Cell): string {
  return String(value ?? "")
    .toLowerCase()
    .replace(/ё/g, "е")
    .replace(LONG_FORMS, " ")
    .replace(/[^a-zа-я0-9]+/gi, " ")
    .trim()
    .split(/\s+/)
    .filter((word) => word && !LEGAL_FORMS.has(word))
    .sort()
    .join(" ");
}

function bigrams(text: string): Map<string, number> {
  const out = new Map<string, number>();
  const padded = ` ${text} `;
  for (let i = 0; i < padded.length - 1; i++) {
    const gram = padded.slice(i, i + 2);
    out.set(gram, (out.get(gram) ?? 0) + 1);
  }
  return out;
}

/** Похожесть названий 0–1 (коэффициент Дайса по парам букв после нормализации). */
export function nameSimilarity(a: Cell, b: Cell): number {
  const x = normalizeName(a);
  const y = normalizeName(b);
  if (!x || !y) return 0;
  if (x === y) return 1;
  // Одно название целиком внутри другого («ромашка» и «ромашка торг»).
  if ((x.length >= 4 && y.includes(x)) || (y.length >= 4 && x.includes(y))) return 0.9;
  const gx = bigrams(x);
  const gy = bigrams(y);
  let common = 0;
  for (const [gram, count] of gx) common += Math.min(count, gy.get(gram) ?? 0);
  const total = [...gx.values()].reduce((s, n) => s + n, 0) + [...gy.values()].reduce((s, n) => s + n, 0);
  return total ? (2 * common) / total : 0;
}

/** ИНН из ячейки: 10 или 12 цифр. */
export function toInn(value: Cell): string | null {
  const digits = String(value ?? "").replace(/\D/g, "");
  return digits.length === 10 || digits.length === 12 ? digits : null;
}

/** Номера счетов и документов из текста: «сч. № 1234», «счет 12/А», «по счёту №00045 от». */
export function invoiceNumbers(value: Cell): string[] {
  const text = String(value ?? "");
  const out = new Set<string>();
  const re = /(?:сч(?:[её]т(?:у|а)?|\.|-ф(?:актур[ае])?)?|сф|сч\.?-?ф\.?|сч\/ф|сф\.|инвойс|invoice|сч\.\s*-?\s*оферт[ае]?|договор(?:у|а)?|дог\.?|заказ(?:у|а)?|накладн(?:ой|ая)|упд|тн)\s*(?:№|n|no\.?|#)?\s*([0-9a-zа-я][0-9a-zа-я/_-]{0,19})/gi;
  for (const m of text.matchAll(re)) {
    const number = m[1].replace(/^0+(?=\d)/, "").toLowerCase();
    if (/\d/.test(number)) out.add(number);
  }
  // Голый «№ 1234» тоже номер документа.
  for (const m of text.matchAll(/№\s*([0-9][0-9a-zа-я/_-]{0,19})/gi)) out.add(m[1].replace(/^0+(?=\d)/, "").toLowerCase());
  return [...out];
}

/** Номер документа из ячейки номера (без слов): «00123» → «123». */
export function docNumber(value: Cell): string | null {
  const text = String(value ?? "").trim().toLowerCase().replace(/^№\s*/, "");
  if (!text || !/\d/.test(text) || text.length > 20) return null;
  return text.replace(/^0+(?=\d)/, "");
}

/* ----------------------------------------------------------- столбцы */

/** Какие столбцы стороны что значат — номера столбцов в таблице (0 — первый). */
export interface SideColumns {
  date?: number;
  amount?: number;
  debit?: number;
  credit?: number;
  names: number[];
  inns: number[];
  /** Тексты, из которых берутся номера счетов: назначение, комментарий. */
  texts: number[];
  /** Столбцы с номером документа. */
  docs: number[];
}

const HEADER_RULES: { key: keyof SideColumns; re: RegExp }[] = [
  { key: "inns", re: /инн|inn|tax ?id/i },
  { key: "debit", re: /дебет|списан|расход|debit|outflow|withdraw/i },
  { key: "credit", re: /кредит|поступ|приход|зачисл|credit|inflow|deposit/i },
  { key: "texts", re: /назначени|описани|комментари|основани|purpose|description|memo|details/i },
  { key: "docs", re: /^(№|номер|n|no\.?)(\s|$)|номер док|№ ?док|номер сч|№ ?сч|сч[её]т ?№|^сч[её]т$|invoice|док(умент)?\s*№|^документ$/i },
  { key: "names", re: /контрагент|получател|плательщик|поставщик|покупател|клиент|наименовани|организаци|партн[её]р|counterparty|payee|payer|vendor|customer|supplier|name/i },
  { key: "amount", re: /сумм|amount|итог|total|к оплате|оплачено|платеж|платёж/i },
  { key: "date", re: /дата|date|число|период/i }
];

/** Подбор столбцов по заголовкам и содержимому. */
export function detectColumns(header: readonly Cell[], rows: readonly (readonly Cell[])[]): SideColumns {
  const result: SideColumns = { names: [], inns: [], texts: [], docs: [] };
  const taken = new Set<number>();
  // Все строки (до 5000): столбец «Поступление» в выписке бывает заполнен
  // одной строкой в самом конце.
  const sample = rows.slice(0, 5000);
  const share = (column: number, test: (value: Cell) => boolean, emptyAs = 0) => {
    const filled = sample.map((row) => row[column]).filter((value) => value !== "" && value !== null && value !== undefined);
    return filled.length ? filled.filter(test).length / filled.length : emptyAs;
  };
  header.forEach((raw, column) => {
    const name = String(raw ?? "").trim();
    for (const rule of HEADER_RULES) {
      if (!rule.re.test(name)) continue;
      // Заголовок «ИНН получателя» — это ИНН, а не название.
      if (rule.key === "names" && /инн|счет|счёт|банк|бик|кпп/i.test(name)) continue;
      if (rule.key === "amount" && /ндс|налог/i.test(name)) continue;
      if (rule.key === "date" && share(column, (value) => toDay(value) !== null) < 0.6) continue;
      if ((rule.key === "amount" || rule.key === "debit" || rule.key === "credit") && share(column, (value) => toNumber(value) !== null, rule.key === "amount" ? 0 : 1) < 0.6) continue;
      if (Array.isArray(result[rule.key])) (result[rule.key] as number[]).push(column);
      else if (result[rule.key] === undefined) (result as any)[rule.key] = column;
      else continue;
      taken.add(column);
      break;
    }
  });
  // Без заголовков: дата — первый столбец из дат, сумма — первый числовой не-ИНН.
  if (result.date === undefined) {
    const column = header.findIndex((_, index) => !taken.has(index) && share(index, (value) => toDay(value) !== null) >= 0.8);
    if (column >= 0) { result.date = column; taken.add(column); }
  }
  if (result.amount === undefined && result.debit === undefined && result.credit === undefined) {
    const column = header.findIndex((_, index) => !taken.has(index) && share(index, (value) => toNumber(value) !== null && toInn(value) === null) >= 0.8);
    if (column >= 0) result.amount = column;
  }
  return result;
}

/* ------------------------------------------------------------ строки */

export interface SideRow {
  /** Номер строки в таблице (0 — первая строка данных). */
  index: number;
  day: number | null;
  /** Сумма по модулю. */
  amount: number | null;
  /** Направление, если есть дебет/кредит: «out» — списание, «in» — поступление. */
  direction?: "in" | "out";
  names: string[];
  inns: string[];
  invoices: string[];
  docs: string[];
  /** Похоже на комиссию банка (по назначению или названию). */
  fee: boolean;
}

export function readSide(rows: readonly (readonly Cell[])[], columns: SideColumns): SideRow[] {
  return rows.map((row, index) => {
    let amount: number | null = null;
    let direction: "in" | "out" | undefined;
    if (columns.amount !== undefined) amount = toNumber(row[columns.amount]);
    if (amount === null && (columns.debit !== undefined || columns.credit !== undefined)) {
      const debit = columns.debit !== undefined ? toNumber(row[columns.debit]) : null;
      const credit = columns.credit !== undefined ? toNumber(row[columns.credit]) : null;
      if (debit) { amount = debit; direction = "out"; } else if (credit) { amount = credit; direction = "in"; }
    }
    const texts = columns.texts.map((column) => String(row[column] ?? ""));
    const names = columns.names.map((column) => String(row[column] ?? "")).filter((value) => value.trim());
    return {
      index,
      day: columns.date !== undefined ? toDay(row[columns.date]) : null,
      amount: amount === null ? null : Math.abs(Math.round(amount * 100) / 100),
      ...(direction ? { direction } : {}),
      names,
      inns: columns.inns.map((column) => toInn(row[column])).filter((value): value is string => Boolean(value)),
      invoices: [...new Set(texts.flatMap((text) => invoiceNumbers(text)))],
      docs: columns.docs.map((column) => docNumber(row[column])).filter((value): value is string => Boolean(value)),
      fee: [...texts, ...names].some((text) => /комисси|обслуживани[ея] сч|плата за|fee|bank charge/i.test(text))
    };
  });
}

/* -------------------------------------------------------- сопоставление */

export interface Tolerances {
  /** Допустимая разница суммы в рублях (0 — точно). */
  amount: number;
  /** Допустимая разница даты в днях. */
  days: number;
  /** Похожесть названия 0–1, с которой оно считается совпавшим. */
  name: number;
  /** Разница суммы, похожая на комиссию банка. */
  fee: number;
  /** Искать сочетания строк (несколько платежей на один счёт). */
  groups: boolean;
}

export const DEFAULT_TOLERANCES: Tolerances = { amount: 0, days: 3, name: 0.8, fee: 500, groups: true };

export type Section = "matched" | "confirmed" | "probable" | "amountDiff" | "fees" | "leftOnly" | "rightOnly";

export interface MatchItem {
  section: Section;
  left: number[];
  right: number[];
  /** 0–100. */
  score: number;
  /** Что совпало и что нет — словами. */
  reason: string;
  /** Разница сумм: слева минус справа. */
  diff: number;
  /** Ключ пары для решений пользователя — устойчив к перестановке строк. */
  key: string;
  /** Ближайший кандидат для строк без пары. */
  nearest?: string;
}

interface Check { amount: "equal" | "near" | "off" | "none"; date: "ok" | "off" | "none"; name: "ok" | "near" | "off" | "none"; inn: "ok" | "off" | "none"; invoice: "ok" | "none"; similarity: number; dayDiff: number | null; amountDiff: number }

function compare(a: SideRow, b: SideRow, t: Tolerances): Check {
  const amountDiff = a.amount !== null && b.amount !== null ? Math.round((a.amount - b.amount) * 100) / 100 : NaN;
  const amount = Number.isNaN(amountDiff) ? "none" : Math.abs(amountDiff) <= 0.005 ? "equal" : Math.abs(amountDiff) <= t.amount + 0.005 ? "near" : "off";
  const dayDiff = a.day !== null && b.day !== null ? b.day - a.day : null;
  const date = dayDiff === null ? "none" : Math.abs(dayDiff) <= t.days ? "ok" : "off";
  let similarity = 0;
  for (const x of a.names) for (const y of b.names) similarity = Math.max(similarity, nameSimilarity(x, y));
  const name = !a.names.length || !b.names.length ? "none" : similarity >= t.name ? "ok" : similarity >= t.name - 0.2 ? "near" : "off";
  const inn = !a.inns.length || !b.inns.length ? "none" : a.inns.some((x) => b.inns.includes(x)) ? "ok" : "off";
  const leftKeys = [...a.invoices, ...a.docs];
  const rightKeys = [...b.invoices, ...b.docs];
  const invoice = leftKeys.some((x) => rightKeys.includes(x)) ? "ok" : "none";
  return { amount, date, name, inn, invoice, similarity, dayDiff, amountDiff };
}

/** Оценка пары 0–100: сумма весит больше всего, ключи (ИНН, счёт) — сильные подтверждения. */
function score(c: Check): number {
  let s = 0;
  s += c.amount === "equal" ? 40 : c.amount === "near" ? 30 : 0;
  s += c.date === "ok" ? 15 : c.date === "none" ? 7 : 0;
  s += c.name === "ok" ? 20 : c.name === "near" ? 10 : c.name === "none" ? 8 : 0;
  s += c.inn === "ok" ? 15 : c.inn === "none" ? 6 : -30;
  s += c.invoice === "ok" ? 15 : 0;
  return Math.max(0, Math.min(100, s));
}

function reasonText(c: Check): string {
  const parts: string[] = [];
  if (c.amount === "equal") parts.push("сумма =");
  else if (c.amount === "near") parts.push(`сумма ±${Math.abs(c.amountDiff)}`);
  else if (c.amount === "off") parts.push(`сумма отличается на ${Math.abs(c.amountDiff)}`);
  if (c.date === "ok") parts.push(c.dayDiff ? `дата ${c.dayDiff > 0 ? "+" : ""}${c.dayDiff} дн.` : "дата =");
  else if (c.date === "off") parts.push(`дата ${c.dayDiff! > 0 ? "+" : ""}${c.dayDiff} дн. (вне допуска)`);
  if (c.inn === "ok") parts.push("ИНН =");
  else if (c.inn === "off") parts.push("ИНН разный");
  if (c.name === "ok") parts.push(c.similarity === 1 ? "название =" : `название ${Math.round(c.similarity * 100)}%`);
  else if (c.name === "near" || c.name === "off") parts.push(`название похоже на ${Math.round(c.similarity * 100)}%`);
  if (c.invoice === "ok") parts.push("номер счёта =");
  return parts.join(", ");
}

/** Все правила выполнены: сумма точно (или в допуске), остальные поля, что есть, совпали. */
function isExact(c: Check): boolean {
  return (c.amount === "equal" || c.amount === "near") && c.date !== "off" && c.inn !== "off" && (c.name === "ok" || c.name === "none" || c.invoice === "ok" || c.inn === "ok");
}

/** Пара с расхождением суммы: всё остальное сильно совпало. */
function isAmountDiff(c: Check): boolean {
  return c.amount === "off" && c.date !== "off" && c.inn !== "off" && (c.invoice === "ok" || c.inn === "ok" || c.name === "ok");
}

const rowSignature = (r: SideRow) => `${r.day ?? ""}:${r.amount ?? ""}:${r.names[0] ? normalizeName(r.names[0]) : ""}`;
/** Номер вхождения среди одинаковых строк своей стороны. Без него у двух
 * одинаковых аренд был один ключ, и «да» у одной ложилось на обе («Книга11»). */
const occurrence = new WeakMap<SideRow, number>();
function numberDuplicates(rows: readonly SideRow[]) {
  const seen = new Map<string, number>();
  for (const row of rows) {
    const signature = rowSignature(row);
    const n = (seen.get(signature) ?? 0) + 1;
    seen.set(signature, n);
    occurrence.set(row, n);
  }
}
const rowKey = (side: "L" | "R", r: SideRow) => `${side}${rowSignature(r)}${(occurrence.get(r) ?? 1) > 1 ? `#${occurrence.get(r)}` : ""}`;
const pairKey = (leftRows: readonly SideRow[], rightRows: readonly SideRow[]) =>
  [...leftRows.map((r) => rowKey("L", r)), ...rightRows.map((r) => rowKey("R", r))].join("|");

export interface ReconcileInput {
  left: SideRow[];
  right: SideRow[];
  tolerances: Tolerances;
  /** Решения пользователя по ключам пар: true — подтверждено, false — отвергнуто. */
  decisions?: Record<string, boolean>;
  /** Как называть строку в «ближайшем кандидате»: по умолчанию «строка N справа». */
  label?: (side: "left" | "right", index: number) => string;
}

export interface ReconcileResult {
  items: MatchItem[];
  /** Проверка: каждая строка в одном разделе. */
  coverage: { left: number; right: number; leftTotal: number; rightTotal: number; ok: boolean };
  tolerances: Tolerances;
}

/** Поиск сочетаний: одна строка против 2–3 строк другой стороны с той же суммой. */
function findGroup(one: SideRow, pool: SideRow[], t: Tolerances, oneIsLeft: boolean): SideRow[] | null {
  if (one.amount === null) return null;
  // Кандидаты: тот же контрагент (ИНН или похожее название) и дата в окне ±(допуск+30 дн.).
  const near = pool.filter((row) => {
    if (row.amount === null || row.amount > one.amount! + t.amount + 0.005) return false;
    const c = oneIsLeft ? compare(one, row, { ...t, days: t.days + 30 }) : compare(row, one, { ...t, days: t.days + 30 });
    return c.date !== "off" && c.inn !== "off" && (c.inn === "ok" || c.name === "ok" || c.invoice === "ok");
  }).slice(0, 25);
  const target = one.amount;
  const fits = (sum: number) => Math.abs(sum - target) <= t.amount + 0.005;
  for (let i = 0; i < near.length; i++) {
    for (let j = i + 1; j < near.length; j++) {
      const two = near[i].amount! + near[j].amount!;
      if (fits(two)) return [near[i], near[j]];
      for (let k = j + 1; k < near.length; k++) {
        if (fits(two + near[k].amount!)) return [near[i], near[j], near[k]];
      }
    }
  }
  return null;
}

export function reconcile(input: ReconcileInput): ReconcileResult {
  numberDuplicates(input.left);
  numberDuplicates(input.right);
  const t = input.tolerances;
  const decisions = input.decisions ?? {};
  const usedLeft = new Set<number>();
  const usedRight = new Set<number>();
  const items: MatchItem[] = [];

  // Кандидаты: все пары с оценкой; по суммам — через индекс, чтобы не сравнивать всё со всем.
  const byCents = new Map<number, SideRow[]>();
  const rightSorted = input.right.filter((row) => row.amount !== null).sort((x, y) => x.amount! - y.amount!);
  for (const row of input.right) if (row.amount !== null) {
    const key = Math.round(row.amount * 100);
    byCents.set(key, [...(byCents.get(key) ?? []), row]);
  }
  type Candidate = { a: SideRow; b: SideRow; c: Check; s: number };
  const candidates: Candidate[] = [];
  const window = Math.max(t.amount, t.fee);
  for (const a of input.left) {
    if (a.amount === null) continue;
    // Суммы в окне допуска и комиссии: бинарный поиск по отсортированным.
    let lo = 0; let hi = rightSorted.length;
    while (lo < hi) { const mid = (lo + hi) >> 1; if (rightSorted[mid].amount! < a.amount - window - 0.005) lo = mid + 1; else hi = mid; }
    for (let i = lo; i < rightSorted.length && rightSorted[i].amount! <= a.amount + window + 0.005; i++) {
      const b = rightSorted[i];
      const c = compare(a, b, t);
      if (c.amount === "off" && !isAmountDiff(c)) continue;
      candidates.push({ a, b, c, s: score(c) });
    }
    // Расхождение суммы при совпавшем номере счёта или ИНН — и далеко от суммы.
    if (a.invoices.length || a.docs.length) {
      for (const b of input.right) {
        if (b.amount === null || Math.abs(b.amount - a.amount) <= window + 0.005) continue;
        const c = compare(a, b, t);
        if (c.invoice === "ok" && isAmountDiff(c)) candidates.push({ a, b, c, s: score(c) });
      }
    }
  }
  void byCents;
  candidates.sort((x, y) => y.s - x.s || Math.abs(x.c.dayDiff ?? 99) - Math.abs(y.c.dayDiff ?? 99) || x.a.index - y.a.index || x.b.index - y.b.index);

  // Сколько равных по оценке кандидатов у строки — спорные не решаются молча.
  const rivals = (side: "a" | "b", row: SideRow, s: number) => candidates.filter((item) => item[side] === row && item.s === s).length;

  const take = (cand: Candidate, allowDiff: boolean) => {
    if (usedLeft.has(cand.a.index) || usedRight.has(cand.b.index)) return;
    const key = pairKey([cand.a], [cand.b]);
    if (decisions[key] === false) return;
    const exact = isExact(cand.c);
    const amountDiff = !exact && isAmountDiff(cand.c);
    if (amountDiff && !allowDiff) return;
    if (!exact && !amountDiff && cand.s < 45) return;
    let section: Section = exact ? "matched" : amountDiff ? "amountDiff" : "probable";
    let reason = reasonText(cand.c);
    const contested = rivals("a", cand.a, cand.s) > 1 || rivals("b", cand.b, cand.s) > 1;
    if (section === "matched" && contested) { section = "probable"; reason += "; есть другой кандидат с той же оценкой"; }
    if (section === "matched" && cand.c.name === "ok" && cand.c.similarity < 1 && cand.c.inn !== "ok" && cand.c.invoice !== "ok") {
      // Название совпало лишь похоже, без ключа — не доказано.
      section = "probable";
    }
    if (section === "amountDiff" && Math.abs(cand.c.amountDiff) <= t.fee) reason += "; похоже на комиссию банка";
    if (decisions[key] === true) { section = "confirmed"; reason += "; подтверждено пользователем"; }
    usedLeft.add(cand.a.index);
    usedRight.add(cand.b.index);
    items.push({ section, left: [cand.a.index], right: [cand.b.index], score: cand.s, reason, diff: cand.c.amountDiff, key });
  };
  // Сначала пары с той же суммой (или в допуске); расхождения сумм — после
  // сочетаний: «50 000 против 30 000» иначе съедало частичную оплату 30 000 + 20 000.
  for (const cand of candidates) take(cand, false);

  // Сочетания: одна строка против нескольких (сборный и частичный платёж).
  if (t.groups) {
    for (const [oneSide, used, otherUsed, pool, isLeft] of [
      [input.left, usedLeft, usedRight, input.right, true],
      [input.right, usedRight, usedLeft, input.left, false]
    ] as const) {
      for (const one of oneSide) {
        if (used.has(one.index) || one.amount === null) continue;
        const free = pool.filter((row) => !otherUsed.has(row.index));
        const group = findGroup(one, free, t, isLeft);
        if (!group) continue;
        const leftRows = isLeft ? [one] : group;
        const rightRows = isLeft ? group : [one];
        const key = pairKey(leftRows, rightRows);
        if (decisions[key] === false) continue;
        used.add(one.index);
        for (const row of group) otherUsed.add(row.index);
        const confirmed = decisions[key] === true;
        items.push({
          section: confirmed ? "confirmed" : "probable",
          left: leftRows.map((row) => row.index),
          right: rightRows.map((row) => row.index),
          score: 60,
          reason: `${isLeft ? "одна строка слева" : "одна строка справа"} = ${group.length} строки ${isLeft ? "справа" : "слева"} (частичная или сборная оплата), сумма сошлась${confirmed ? "; подтверждено пользователем" : ""}`,
          diff: 0,
          key
        });
      }
    }
  }

  for (const cand of candidates) take(cand, true);

  // Комиссии банка без пары — свой раздел, а не «только справа/слева».
  for (const [rows, used, side] of [[input.left, usedLeft, "left"], [input.right, usedRight, "right"]] as const) {
    for (const row of rows) {
      if (used.has(row.index) || !row.fee) continue;
      used.add(row.index);
      items.push({ section: "fees", left: side === "left" ? [row.index] : [], right: side === "right" ? [row.index] : [], score: 0, reason: "комиссия или плата банка", diff: 0, key: pairKey(side === "left" ? [row] : [], side === "right" ? [row] : []) });
    }
  }

  // Без пары — с ближайшим кандидатом, чтобы пользователь видел, почему нет.
  const nearestFor = (row: SideRow, pool: SideRow[], isLeft: boolean) => {
    let best: { s: number; text: string } | null = null;
    for (const other of pool) {
      if (other.amount === null || row.amount === null) continue;
      const c = isLeft ? compare(row, other, t) : compare(other, row, t);
      const s = score(c);
      const name = input.label ? input.label(isLeft ? "right" : "left", other.index) : `строка ${other.index + 1} ${isLeft ? "справа" : "слева"}`;
      if (s >= 30 && (!best || s > best.s)) best = { s, text: `${name}: ${reasonText(c)}` };
    }
    return best?.text;
  };
  for (const row of input.left) if (!usedLeft.has(row.index)) {
    const nearest = nearestFor(row, input.right, true);
    items.push({ section: "leftOnly", left: [row.index], right: [], score: 0, reason: row.amount === null ? "нет суммы" : "пары не найдено", diff: 0, key: pairKey([row], []), ...(nearest ? { nearest } : {}) });
  }
  for (const row of input.right) if (!usedRight.has(row.index)) {
    const nearest = nearestFor(row, input.left, false);
    items.push({ section: "rightOnly", left: [], right: [row.index], score: 0, reason: row.amount === null ? "нет суммы" : "пары не найдено", diff: 0, key: pairKey([], [row]), ...(nearest ? { nearest } : {}) });
  }

  const seenLeft = items.flatMap((item) => item.left);
  const seenRight = items.flatMap((item) => item.right);
  const coverage = {
    left: new Set(seenLeft).size,
    right: new Set(seenRight).size,
    leftTotal: input.left.length,
    rightTotal: input.right.length,
    ok: seenLeft.length === input.left.length && seenRight.length === input.right.length &&
      new Set(seenLeft).size === input.left.length && new Set(seenRight).size === input.right.length
  };
  return { items, coverage, tolerances: t };
}

export const SECTION_TITLE: Record<Section, string> = {
  confirmed: "Подтверждено пользователем",
  matched: "Совпало",
  probable: "Вероятно — проверьте и поставьте «да» или «нет» в столбце «Решение»",
  amountDiff: "Расхождение суммы",
  fees: "Комиссии и платы банка",
  leftOnly: "Только в первой таблице",
  rightOnly: "Только во второй таблице"
};

export const SECTION_ORDER: readonly Section[] = ["matched", "confirmed", "probable", "amountDiff", "fees", "leftOnly", "rightOnly"];
