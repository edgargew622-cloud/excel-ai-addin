/**
 * Разметка страницы PDF в сетку листа Excel (перенос «как есть», 01.10.2026).
 *
 * Пользователю нужна точная копия таблицы или бланка из PDF: те же строки и
 * столбцы, объединённые ячейки, рамки, жирный текст, похожие размеры. pdf.js
 * даёт положение каждого куска текста и линии бланка; здесь из них строится
 * сетка:
 * - столбцы — по вертикальным линиям, а где линий нет — по просветам между
 *   колонками текста, одинаковым во всех строках;
 * - строки — по горизонтальным линиям и по строкам текста;
 * - соседние клетки внутри одной нарисованной рамки без разделяющей линии
 *   объединяются; строки записей таблицы без горизонтальных линий — нет
 *   (у них текст стоит в нескольких столбцах на одной высоте);
 * - рамки — по линиям, лежащим на краях клеток.
 *
 * Координаты — пункты, начало в левом верхнем углу страницы (y вниз).
 * Модуль не знает ни pdf.js, ни Excel: это чистая геометрия, её проверяют тесты.
 */

export interface PdfText {
  str: string;
  x0: number;
  x1: number;
  top: number;
  bottom: number;
  size: number;
  bold: boolean;
}

/** Отрезок, параллельный оси: x1 ≤ x2, y1 ≤ y2; width — толщина линии. */
export interface PdfSegment {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  width: number;
}

export interface PdfPageRaw {
  width: number;
  height: number;
  texts: PdfText[];
  segments: PdfSegment[];
}

export type Align = "Left" | "Center" | "Right";
export type BorderWeight = "Thin" | "Medium";

export interface LayoutCell {
  r: number;
  c: number;
  rowSpan: number;
  colSpan: number;
  /** Текст как в файле. */
  text: string;
  /** Что писать в ячейку: число, дата (серийный номер Excel) или текст. */
  value: string | number;
  /** «@» — текст; для чисел и дат — их формат. */
  numberFormat: string;
  bold: boolean;
  size: number;
  align: Align;
  wrap: boolean;
}

export interface SheetLayout {
  columnWidths: number[];
  rowHeights: number[];
  cells: LayoutCell[];
  /** Горизонтальная граница над строкой r (r = rows — под последней), столбец c. */
  hEdges: Array<[number, number, BorderWeight]>;
  /** Вертикальная граница слева от столбца c (c = columns — справа от последнего), строка r. */
  vEdges: Array<[number, number, BorderWeight]>;
  pages: number;
  /** С какой строки листа (с 0) начинается каждая страница. */
  pageStarts: number[];
  warnings: string[];
}

const LINE_TOL = 1.6;
/** Линия относится к границе сетки, если ближе этого: границы из близких линий усредняются. */
const SNAP = 2.6;
const MIN_SEGMENT = 4;
const MERGE_BOUNDARY = 4;
/** Рамка выше этого — раздел бланка, а не ячейка: строки текста в ней не сливаются. */
const MAX_CELL_BOX = 48;
export const MAX_LAYOUT_CELLS = 60_000;

/* ---------------------------------------------------------------- отрезки -- */

interface HLine { y: number; x1: number; x2: number; width: number }
interface VLine { x: number; y1: number; y2: number; width: number }

function joinRuns<T extends { a: number; b: number; width: number }>(runs: T[], gap: number): T[] {
  const sorted = [...runs].sort((p, q) => p.a - q.a);
  const out: T[] = [];
  for (const run of sorted) {
    const last = out[out.length - 1];
    if (last && run.a <= last.b + gap) {
      last.b = Math.max(last.b, run.b);
      last.width = Math.max(last.width, run.width);
    } else out.push({ ...run });
  }
  return out;
}

/** Соседние и перекрывающиеся отрезки одной прямой — один отрезок. */
export function normalizeSegments(segments: readonly PdfSegment[]): { h: HLine[]; v: VLine[] } {
  const hs = segments.filter((s) => s.y2 - s.y1 <= LINE_TOL && s.x2 - s.x1 >= MIN_SEGMENT)
    .map((s) => ({ key: (s.y1 + s.y2) / 2, a: s.x1, b: s.x2, width: s.width }));
  const vs = segments.filter((s) => s.x2 - s.x1 <= LINE_TOL && s.y2 - s.y1 >= MIN_SEGMENT)
    .map((s) => ({ key: (s.x1 + s.x2) / 2, a: s.y1, b: s.y2, width: s.width }));
  const group = (items: typeof hs) => {
    const sorted = [...items].sort((p, q) => p.key - q.key);
    const lines: Array<{ key: number; runs: typeof hs }> = [];
    for (const item of sorted) {
      const last = lines[lines.length - 1];
      if (last && item.key - last.key <= LINE_TOL) last.runs.push(item);
      else lines.push({ key: item.key, runs: [item] });
    }
    return lines.flatMap((line) => {
      const key = line.runs.reduce((sum, run) => sum + run.key, 0) / line.runs.length;
      return joinRuns(line.runs, 3).map((run) => ({ key, a: run.a, b: run.b, width: run.width }));
    });
  };
  return {
    h: group(hs).map((s) => ({ y: s.key, x1: s.a, x2: s.b, width: s.width })),
    v: group(vs).map((s) => ({ x: s.key, y1: s.a, y2: s.b, width: s.width }))
  };
}

/** Доля отрезка [a, b], покрытая линиями. */
function covered(runs: ReadonlyArray<{ a: number; b: number }>, a: number, b: number): number {
  if (b <= a) return 0;
  let total = 0;
  for (const run of joinRuns(runs.map((run) => ({ ...run, width: 0 })), 0)) {
    total += Math.max(0, Math.min(b, run.b) - Math.max(a, run.a));
  }
  return total / (b - a);
}

/* ------------------------------------------------------------------ текст -- */

/** Кусок текста с большими пробелами внутри — несколько кусков: так бывает, когда
 * строку таблицы PDF собрал одной строкой с пробелами вместо столбцов. */
export function splitWideGaps(texts: readonly PdfText[]): PdfText[] {
  const out: PdfText[] = [];
  for (const text of texts) {
    const str = text.str.replace(/\s+$/, "");
    if (!str.trim()) continue;
    if (!/\S\s{3,}\S/.test(str)) { out.push({ ...text, str: str.trimStart(), x0: text.x0 + charShift(text, str) }); continue; }
    const perChar = (text.x1 - text.x0) / Math.max(1, text.str.length);
    const pattern = /\S+(?:\s{1,2}\S+)*/g;
    let match: RegExpExecArray | null;
    while ((match = pattern.exec(str))) {
      out.push({ ...text, str: match[0], x0: text.x0 + match.index * perChar, x1: text.x0 + (match.index + match[0].length) * perChar });
    }
  }
  return out;
}

function charShift(text: PdfText, str: string): number {
  const lead = str.length - str.trimStart().length;
  return lead ? lead * (text.x1 - text.x0) / Math.max(1, text.str.length) : 0;
}

interface TextLine { top: number; bottom: number; items: PdfText[] }

/** Строки текста: куски, перекрывающиеся по высоте больше чем наполовину. */
function textLines(texts: readonly PdfText[]): TextLine[] {
  const sorted = [...texts].sort((p, q) => (p.top + p.bottom) - (q.top + q.bottom));
  const lines: TextLine[] = [];
  for (const text of sorted) {
    const line = lines.find((candidate) => {
      const overlap = Math.min(candidate.bottom, text.bottom) - Math.max(candidate.top, text.top);
      return overlap > 0.5 * Math.min(candidate.bottom - candidate.top, text.bottom - text.top);
    });
    if (line) { line.items.push(text); line.top = Math.min(line.top, text.top); line.bottom = Math.max(line.bottom, text.bottom); }
    else lines.push({ top: text.top, bottom: text.bottom, items: [text] });
  }
  for (const line of lines) line.items.sort((p, q) => p.x0 - q.x0);
  return lines.sort((p, q) => p.top - q.top);
}

/* ---------------------------------------------------------------- значения -- */

const GROUP = "[ \\u00a0\\u202f]";

export function cellValue(text: string): { value: string | number; numberFormat: string } {
  const s = text.trim();
  const date = /^(\d{2})\.(\d{2})\.(\d{4})$/.exec(s);
  if (date) {
    const [day, month, year] = [Number(date[1]), Number(date[2]), Number(date[3])];
    const utc = Date.UTC(year, month - 1, day);
    const check = new Date(utc);
    if (year >= 1900 && check.getUTCDate() === day && check.getUTCMonth() === month - 1) {
      return { value: utc / 86_400_000 + 25_569, numberFormat: "dd.mm.yyyy" };
    }
  }
  const grouped = new RegExp(`^(-?)(\\d{1,3}(?:${GROUP}\\d{3})+)(?:,(\\d+))?$`).exec(s);
  const plain = /^(-?)(0|[1-9]\d{0,6})(?:,(\d+))?$/.exec(s);
  const match = grouped ?? plain;
  if (match) {
    const whole = match[2].replace(new RegExp(GROUP, "g"), "");
    const decimals = match[3] ?? "";
    const value = Number(`${match[1]}${whole}${decimals ? `.${decimals}` : ""}`);
    if (Number.isFinite(value)) {
      const format = grouped || decimals
        ? `#,##0${decimals ? `.${"0".repeat(decimals.length)}` : ""}`
        : "General";
      return { value, numberFormat: format };
    }
  }
  // Коды, номера, ИНН, время, телефоны, проценты — текстом, как в файле.
  return { value: s, numberFormat: "@" };
}

/* ------------------------------------------------------------------ сетка -- */

interface Boundary { at: number; line: boolean }

function clusterBoundaries(items: readonly Boundary[], tol: number): number[] {
  const sorted = [...items].sort((p, q) => p.at - q.at);
  const groups: Boundary[][] = [];
  for (const item of sorted) {
    const last = groups[groups.length - 1];
    if (last && item.at - last[last.length - 1].at <= tol) last.push(item);
    else groups.push([item]);
  }
  return groups.map((group) => {
    const lines = group.filter((item) => item.line);
    const pick = lines.length ? lines : group;
    return pick.reduce((sum, item) => sum + item.at, 0) / pick.length;
  });
}

function locate(bounds: readonly number[], at: number): number {
  if (at < bounds[0]) return 0;
  for (let i = 0; i < bounds.length - 1; i++) if (at < bounds[i + 1]) return i;
  return bounds.length - 2;
}

interface PagePlan {
  raw: PdfPageRaw;
  texts: PdfText[];
  h: HLine[];
  v: VLine[];
  enclosed: (x: number, y: number) => boolean;
  lines: TextLine[];
}

function planPage(raw: PdfPageRaw): PagePlan {
  const texts = splitWideGaps(raw.texts);
  const { h, v } = normalizeSegments(raw.segments);
  const enclosed = (x: number, y: number) =>
    v.some((line) => line.x <= x && line.y1 - LINE_TOL <= y && line.y2 + LINE_TOL >= y) &&
    v.some((line) => line.x >= x && line.y1 - LINE_TOL <= y && line.y2 + LINE_TOL >= y) &&
    h.some((line) => line.y <= y && line.x1 - LINE_TOL <= x && line.x2 + LINE_TOL >= x) &&
    h.some((line) => line.y >= y && line.x1 - LINE_TOL <= x && line.x2 + LINE_TOL >= x);
  return { raw, texts, h, v, enclosed, lines: textLines(texts) };
}

/** Просветы между колонками текста вне рамок — одинаковые во всех строках. */
function textColumnGaps(page: PagePlan): number[] {
  const free = page.lines
    .map((line) => line.items.filter((item) => !page.enclosed((item.x0 + item.x1) / 2, (item.top + item.bottom) / 2)))
    .filter((items) => items.length >= 2);
  if (!free.length) return [];
  const merged = joinRuns(free.flat().map((item) => ({ a: item.x0, b: item.x1, width: 0 })), 0.5);
  // Просвет между колонками — шире пробела между словами; пусто он должен
  // быть во всех строках сразу, поэтому случайный пробел колонку не разрежет.
  const sizes = free.flat().map((item) => item.size).sort((p, q) => p - q);
  const minGap = Math.max(2, 0.3 * sizes[Math.floor(sizes.length / 2)]);
  const gaps: number[] = [];
  for (let i = 0; i + 1 < merged.length; i++) {
    const gap = merged[i + 1].a - merged[i].b;
    if (gap >= minGap) gaps.push(merged[i].b + gap / 2);
  }
  return [merged[0].a - 1, ...gaps, merged[merged.length - 1].b + 1];
}

/** Строки «метка — значение» вне рамок (два куска с заметным просветом):
 * значение начинает свой столбец, иначе метка и значение слились бы в одну
 * ячейку и потеряли своё оформление (путевой лист 01.10.2026: «Организация»). */
function labelValueStarts(page: PagePlan): number[] {
  return page.lines
    .map((line) => line.items.filter((item) => !page.enclosed((item.x0 + item.x1) / 2, (item.top + item.bottom) / 2)))
    .filter((items) => items.length === 2 && items[1].x0 - items[0].x1 > 1.2 * Math.max(items[0].size, items[1].size))
    .map((items) => items[1].x0 - 1);
}

export function buildSheetLayout(rawPages: readonly PdfPageRaw[]): SheetLayout {
  const warnings: string[] = [];
  const pages = rawPages.map(planPage);
  const allTexts = pages.flatMap((page) => page.texts);
  if (!allTexts.length) throw new Error("В PDF нет текстового слоя — похоже на скан.");

  // Столбцы — общие для всех страниц: страницы одного документа обычно
  // повторяют одну сетку, а ширина столбца на листе одна.
  const xs: Boundary[] = [];
  for (const page of pages) {
    for (const line of page.v) xs.push({ at: line.x, line: true });
    for (const at of textColumnGaps(page)) xs.push({ at, line: false });
    for (const at of labelValueStarts(page)) xs.push({ at, line: false });
  }
  xs.push({ at: Math.min(...allTexts.map((text) => text.x0)) - 1, line: false });
  let X = clusterBoundaries(xs, MERGE_BOUNDARY);
  if (X.length < 2) X = [X[0] ?? 0, Math.max(...allTexts.map((text) => text.x1)) + 1];
  const columns = X.length - 1;

  const columnWidths = X.slice(1).map((x, i) => Math.max(1, Math.min(1500, x - X[i])));
  const rowHeights: number[] = [];
  const cells: LayoutCell[] = [];
  const hEdges: Array<[number, number, BorderWeight]> = [];
  const vEdges: Array<[number, number, BorderWeight]> = [];
  const pageStarts: number[] = [];

  pages.forEach((page, pageIndex) => {
    if (pageIndex > 0) rowHeights.push(12);
    const base = rowHeights.length;
    pageStarts.push(base);

    const ys: Boundary[] = page.h.map((line) => ({ at: line.y, line: true }));
    // Соседние строки текста делит одна граница посередине; пустая строка
    // листа появляется только там, где в PDF заметный отступ.
    page.lines.forEach((line, i) => {
      const size = Math.max(...line.items.map((item) => item.size));
      const previous = page.lines[i - 1];
      if (previous && line.top - previous.bottom <= 0.8 * size) {
        ys.push({ at: (previous.bottom + line.top) / 2, line: false });
      } else {
        if (previous) ys.push({ at: previous.bottom + 1, line: false });
        ys.push({ at: line.top - 1, line: false });
      }
      if (i === page.lines.length - 1) ys.push({ at: line.bottom + 1, line: false });
    });
    let Y = clusterBoundaries(ys, 3);
    if (Y.length < 2) return;
    // Разрез по строке текста, который во всех столбцах проходит внутри одной
    // объединённой ячейки, лишний: строка таблицы с рамками — одна строка листа.
    for (let pass = 0; pass < 3; pass++) {
      const keep = removableRowCuts(page, X, Y);
      if (keep.length === Y.length) break;
      Y = keep;
    }
    const rows = Y.length - 1;
    for (let r = 0; r < rows; r++) rowHeights.push(Math.max(2, Math.min(409, Y[r + 1] - Y[r])));

    const { byCell, find, cellEnclosed } = unionCells(page, X, Y);
    const groups = new Map<number, number[]>();
    for (let i = 0; i < rows * columns; i++) groups.set(find(i), [...(groups.get(find(i)) ?? []), i]);

    const done = new Set<number>();
    for (const members of groups.values()) {
      const rs = members.map((i) => Math.floor(i / columns));
      const cs = members.map((i) => i % columns);
      const [r0, r1, c0, c1] = [Math.min(...rs), Math.max(...rs), Math.min(...cs), Math.max(...cs)];
      const rectangle = members.length === (r1 - r0 + 1) * (c1 - c0 + 1);
      const spans = rectangle ? [{ r0, r1, c0, c1, members }] : members.map((i) => ({ r0: Math.floor(i / columns), r1: Math.floor(i / columns), c0: i % columns, c1: i % columns, members: [i] }));
      for (const span of spans) {
        const items = span.members.flatMap((i) => byCell.get(`${Math.floor(i / columns)},${i % columns}`) ?? []);
        span.members.forEach((i) => done.add(i));
        if (!items.length && span.r0 === span.r1 && span.c0 === span.c1) continue;
        cells.push(makeCell(items, span, X, Y, base, cellEnclosed(span.r0, span.c0)));
      }
    }

    // Заголовок по центру страницы вне рамок — объединение на всю ширину.
    for (const cell of cells.filter((item) => item.r >= base && item.rowSpan === 1 && item.colSpan === 1)) {
      const r = cell.r - base;
      const sameRow = cells.filter((item) => item.r === cell.r);
      const items = byCell.get(`${r},${cell.c}`) ?? [];
      if (sameRow.length !== 1 || !items.length || cellEnclosed(r, cell.c)) continue;
      const x0 = Math.min(...items.map((item) => item.x0));
      const x1 = Math.max(...items.map((item) => item.x1));
      const middle = (X[0] + X[columns]) / 2;
      if (Math.abs((x0 + x1) / 2 - middle) <= 12 && x1 - x0 < 0.9 * (X[columns] - X[0]) && x0 - X[0] > 20) {
        cell.c = 0;
        cell.colSpan = columns;
        cell.align = "Center";
      }
    }

    // Рамки: линия, лежащая на краю клетки.
    const weight = (width: number): BorderWeight => (width >= 1.3 ? "Medium" : "Thin");
    for (let r = 0; r <= rows; r++) {
      for (let c = 0; c < columns; c++) {
        const lines = page.h.filter((line) => Math.abs(line.y - Y[r]) <= SNAP);
        if (covered(lines.map((line) => ({ a: line.x1, b: line.x2 })), X[c], X[c + 1]) >= 0.6) {
          hEdges.push([base + r, c, weight(Math.max(...lines.map((line) => line.width)))]);
        }
      }
    }
    for (let c = 0; c <= columns; c++) {
      for (let r = 0; r < rows; r++) {
        const lines = page.v.filter((line) => Math.abs(line.x - X[c]) <= SNAP);
        if (covered(lines.map((line) => ({ a: line.y1, b: line.y2 })), Y[r], Y[r + 1]) >= 0.6) {
          vEdges.push([base + r, c, weight(Math.max(...lines.map((line) => line.width)))]);
        }
      }
    }
    void done;
  });

  if (rowHeights.length * columns > MAX_LAYOUT_CELLS) {
    throw new Error(`Сетка копии — ${rowHeights.length} × ${columns} клеток, больше ${MAX_LAYOUT_CELLS}: перенесите меньше страниц.`);
  }
  return { columnWidths, rowHeights, cells, hEdges, vEdges, pages: pages.length, pageStarts, warnings };
}

/** Границы строк без лишних разрезов (см. buildSheetLayout). */
function removableRowCuts(page: PagePlan, X: readonly number[], Y: readonly number[]): number[] {
  const rows = Y.length - 1;
  const columns = X.length - 1;
  const { find } = unionCells(page, X, Y);
  const isLine = (y: number) => page.h.some((line) => Math.abs(line.y - y) <= SNAP);
  return Y.filter((y, i) => {
    if (i === 0 || i === rows || isLine(y)) return true;
    for (let c = 0; c < columns; c++) if (find((i - 1) * columns + c) !== find(i * columns + c)) return true;
    return false;
  });
}

/** Клетки внутри одной нарисованной рамки без разделяющей линии — одна ячейка. */
function unionCells(page: PagePlan, X: readonly number[], Y: readonly number[]) {
  const rows = Y.length - 1;
  const columns = X.length - 1;
  const byCell = new Map<string, PdfText[]>();
  for (const text of page.texts) {
    const key = `${locate(Y, (text.top + text.bottom) / 2)},${locate(X, text.x0 + 0.5)}`;
    byCell.set(key, [...(byCell.get(key) ?? []), text]);
  }
  const occupied = (r: number) => new Set([...byCell.keys()].filter((key) => key.startsWith(`${r},`)).map((key) => key.split(",")[1])).size;
  const cellEnclosed = (r: number, c: number) => page.enclosed((X[c] + X[c + 1]) / 2, (Y[r] + Y[r + 1]) / 2);
  const vRuns = (x: number) => page.v.filter((line) => Math.abs(line.x - x) <= SNAP).map((line) => ({ a: line.y1, b: line.y2 }));
  const hRuns = (y: number) => page.h.filter((line) => Math.abs(line.y - y) <= SNAP).map((line) => ({ a: line.x1, b: line.x2 }));
  /** Можно ли слить строки внутри рамки вокруг точки (середина столбца c, высота y):
   * низкая рамка — ячейка; высокая — тоже, если в этом столбце внутри неё не
   * больше двух строк текста (крупная надпись по центру), иначе это раздел бланка. */
  const cellLikeBox = (c: number, y: number) => {
    const x = (X[c] + X[c + 1]) / 2;
    const across = page.h.filter((line) => line.x1 - LINE_TOL <= x && line.x2 + LINE_TOL >= x);
    const above = Math.max(-Infinity, ...across.filter((line) => line.y <= y).map((line) => line.y));
    const below = Math.min(Infinity, ...across.filter((line) => line.y >= y).map((line) => line.y));
    if (below - above <= MAX_CELL_BOX) return true;
    const inside = page.texts.filter((text) => {
      const cx = (text.x0 + text.x1) / 2;
      const cy = (text.top + text.bottom) / 2;
      return cx >= X[c] && cx < X[c + 1] && cy > above && cy < below;
    });
    return new Set(inside.map((text) => Math.round((text.top + text.bottom) / 4))).size <= 2;
  };
  const parent = Array.from({ length: rows * columns }, (_, i) => i);
  const find = (i: number): number => (parent[i] === i ? i : (parent[i] = find(parent[i])));
  const unite = (i: number, j: number) => { parent[find(i)] = find(j); };
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < columns; c++) {
      if (!cellEnclosed(r, c)) continue;
      if (c + 1 < columns && cellEnclosed(r, c + 1) && covered(vRuns(X[c + 1]), Y[r], Y[r + 1]) < 0.5) unite(r * columns + c, r * columns + c + 1);
      if (r + 1 < rows && cellEnclosed(r + 1, c) && covered(hRuns(Y[r + 1]), X[c], X[c + 1]) < 0.5) {
        // Строки записей таблицы без горизонтальных линий не сливаются:
        // у них текст стоит в нескольких столбцах на одной высоте. Не сливаются
        // и строки внутри высокой рамки — это раздел бланка (путевой лист
        // 01.10.2026: весь верхний блок вышел одной ячейкой).
        // Если на этой высоте где-то есть линия, это граница настоящих строк
        // таблицы, а клетка без линии здесь — высокая ячейка: её сливаем.
        const lineHere = page.h.some((line) => Math.abs(line.y - Y[r + 1]) <= SNAP);
        const record = !lineHere && occupied(r) >= 2 && occupied(r + 1) >= 2;
        if (!record && cellLikeBox(c, Y[r + 1])) unite(r * columns + c, (r + 1) * columns + c);
      }
    }
  }
  return { byCell, find, cellEnclosed };
}

function makeCell(
  items: PdfText[],
  span: { r0: number; r1: number; c0: number; c1: number },
  X: readonly number[],
  Y: readonly number[],
  base: number,
  boxed: boolean
): LayoutCell {
  const lines = textLines(items);
  const text = lines.map((line) => line.items.reduce((out, item, i) => {
    if (i === 0) return item.str;
    const gap = item.x0 - line.items[i - 1].x1;
    return out + (gap > 0.2 * item.size || /\s$/.test(out) ? " " : "") + item.str;
  }, "").trim()).join("\n");
  const chars = (predicate: (item: PdfText) => boolean) => items.filter(predicate).reduce((sum, item) => sum + item.str.length, 0);
  const bold = items.length > 0 && chars((item) => item.bold) * 2 >= chars(() => true);
  const size = items.length ? Math.max(6, Math.min(36, Math.round(Math.max(...items.map((item) => item.size)) * 2) / 2)) : 10;
  let align: Align = "Left";
  if (items.length) {
    const left = Math.min(...items.map((item) => item.x0)) - X[span.c0];
    const right = X[span.c1 + 1] - Math.max(...items.map((item) => item.x1));
    // Текст длиннее своей ячейки — по левому краю: тогда Excel продолжит его
    // вправо, как в PDF (иначе видно только окончание — «000001» вместо строки).
    if (right < -1) align = "Left";
    else if (left > 3 && Math.abs(left - right) <= Math.max(2.5, 0.12 * (left + right))) align = "Center";
    else if (right < left && right < 6) align = "Right";
    else if (!boxed && right < left) align = "Right";
  }
  const { value, numberFormat } = text.includes("\n") || !text ? { value: text, numberFormat: "@" } : cellValue(text);
  void Y;
  return {
    r: base + span.r0,
    c: span.c0,
    rowSpan: span.r1 - span.r0 + 1,
    colSpan: span.c1 - span.c0 + 1,
    text,
    value,
    numberFormat,
    bold,
    size,
    align,
    wrap: text.includes("\n")
  };
}
