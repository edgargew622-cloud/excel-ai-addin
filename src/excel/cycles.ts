/**
 * Циклические ссылки (этап 8, 8.4.2).
 *
 * Замер 28 сентября 2026 года (Office 2021): ячейки в цикле Excel ошибкой не
 * помечает — в них 0 или последнее посчитанное значение, тип «число». По
 * значениям цикл не виден, поэтому он ищется по формулам: граф зависимостей
 * между ячейками с формулами, в том числе через другие листы и имена
 * диапазонов (имя Excel отдаёт как «=Лист!$D$1»). Замкнутая цепочка —
 * доказанная находка с адресами.
 *
 * Не видны: адреса, которые формула вычисляет (INDIRECT/OFFSET), имена,
 * ссылающиеся не на диапазон, и листы, которые не разбирались. Об этом
 * аудит говорит отдельно.
 */

import { formulaReferences, type FormulaReference } from "./rowOps";
import { columnLetters } from "./formulaFill";

export interface CycleSheet {
  name: string;
  rowIndex: number;
  columnIndex: number;
  formulas: readonly (readonly unknown[])[];
}

/** Имя диапазона: как его отдаёт Excel, и лист, если имя задано только для листа. */
export interface WorkbookName {
  name: string;
  formula: string;
  scope?: string;
}

const isFormula = (value: unknown): value is string => typeof value === "string" && value.startsWith("=");
const lower = (text: string) => text.trim().toLowerCase();
const key = (sheet: string, row: number, column: number) => `${lower(sheet)}!${row},${column}`;

/** Имена, упомянутые в формуле: слово вне кавычек, не функция, не часть адреса. */
function namesIn(formula: string, known: ReadonlySet<string>): string[] {
  const bare = formula.replace(/"(?:[^"]|"")*"/g, '""').replace(/'[^']*'!/g, "");
  const found: string[] = [];
  const pattern = /(?<![\p{L}\p{N}_.!$])([\p{L}_][\p{L}\p{N}_.]*)(?![\p{L}\p{N}_.(!])/gu;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(bare))) {
    const word = lower(match[1]);
    if (known.has(word)) found.push(word);
  }
  return found;
}

/**
 * Все замкнутые цепочки: по одной на каждую группу ячеек, зависящих друг от
 * друга по кругу (сильно связные компоненты), с путём для показа.
 */
export function findCycles(sheets: readonly CycleSheet[], names: readonly WorkbookName[] = []): { cells: string[]; size: number }[] {
  const display = new Map<string, string>();
  const formulaOf = new Map<string, string>();
  const sheetOf = new Map<string, string>();
  // Ячейки с формулами по листу и столбцу — чтобы быстро находить их внутри диапазона.
  const byColumn = new Map<string, Map<number, number[]>>();
  for (const sheet of sheets) {
    const columns = new Map<number, number[]>();
    sheet.formulas.forEach((row, r) => row.forEach((formula, c) => {
      if (!isFormula(formula)) return;
      const rowNumber = sheet.rowIndex + r + 1;
      const columnNumber = sheet.columnIndex + c + 1;
      const id = key(sheet.name, rowNumber, columnNumber);
      display.set(id, `${sheet.name}!${columnLetters(columnNumber)}${rowNumber}`);
      formulaOf.set(id, formula);
      sheetOf.set(id, sheet.name);
      if (!columns.has(columnNumber)) columns.set(columnNumber, []);
      columns.get(columnNumber)!.push(rowNumber);
    }));
    byColumn.set(lower(sheet.name), columns);
  }

  const nameRefs = new Map<string, { scope?: string; refs: FormulaReference[] }[]>();
  for (const item of names) {
    const refs = formulaReferences(item.formula).filter((ref) => ref.sheet);
    if (!refs.length) continue;
    const list = nameRefs.get(lower(item.name)) ?? [];
    list.push({ scope: item.scope, refs });
    nameRefs.set(lower(item.name), list);
  }
  const knownNames = new Set(nameRefs.keys());

  const targets = (sheetName: string, ref: FormulaReference): string[] => {
    const target = ref.sheet ?? sheetName;
    const columns = byColumn.get(lower(target));
    if (!columns) return [];
    const found: string[] = [];
    for (const [column, rows] of columns) {
      if (column < ref.columnStart || column > ref.columnEnd) continue;
      for (const row of rows) if (row >= ref.rowStart && row <= ref.rowEnd) found.push(key(target, row, column));
    }
    return found;
  };

  const edges = new Map<string, string[]>();
  for (const [id, formula] of formulaOf) {
    const sheetName = sheetOf.get(id)!;
    const refs = [...formulaReferences(formula)];
    for (const name of namesIn(formula, knownNames)) {
      const variants = nameRefs.get(name)!;
      // Имя листа важнее имени книги с тем же названием, как в Excel.
      const chosen = variants.find((item) => item.scope && lower(item.scope) === lower(sheetName)) ?? variants.find((item) => !item.scope);
      if (chosen) refs.push(...chosen.refs);
    }
    edges.set(id, [...new Set(refs.flatMap((ref) => targets(sheetName, ref)))]);
  }

  // Тарьян без рекурсии: в больших моделях цепочки длинные.
  let counter = 0;
  const index = new Map<string, number>();
  const low = new Map<string, number>();
  const onStack = new Set<string>();
  const stack: string[] = [];
  const components: string[][] = [];
  for (const start of edges.keys()) {
    if (index.has(start)) continue;
    const work: { node: string; next: number }[] = [{ node: start, next: 0 }];
    index.set(start, counter); low.set(start, counter); counter++;
    stack.push(start); onStack.add(start);
    while (work.length) {
      const frame = work[work.length - 1];
      const out = edges.get(frame.node) ?? [];
      if (frame.next < out.length) {
        const to = out[frame.next++];
        if (!index.has(to)) {
          index.set(to, counter); low.set(to, counter); counter++;
          stack.push(to); onStack.add(to);
          work.push({ node: to, next: 0 });
        } else if (onStack.has(to)) {
          low.set(frame.node, Math.min(low.get(frame.node)!, index.get(to)!));
        }
        continue;
      }
      work.pop();
      if (work.length) {
        const parent = work[work.length - 1].node;
        low.set(parent, Math.min(low.get(parent)!, low.get(frame.node)!));
      }
      if (low.get(frame.node) === index.get(frame.node)) {
        const component: string[] = [];
        let node: string;
        do { node = stack.pop()!; onStack.delete(node); component.push(node); } while (node !== frame.node);
        const selfLoop = component.length === 1 && (edges.get(component[0]) ?? []).includes(component[0]);
        if (component.length > 1 || selfLoop) components.push(component);
      }
    }
  }

  // Путь для показа: от первой ячейки группы по связям внутри группы назад к ней.
  return components.map((component) => {
    const inside = new Set(component);
    const first = [...component].sort()[0];
    const previous = new Map<string, string>();
    const queue = [first];
    const seen = new Set([first]);
    let closing: string | null = (edges.get(first) ?? []).includes(first) ? first : null;
    while (queue.length && !closing) {
      const node = queue.shift()!;
      for (const to of edges.get(node) ?? []) {
        if (!inside.has(to)) continue;
        if (to === first) { closing = node; break; }
        if (!seen.has(to)) { seen.add(to); previous.set(to, node); queue.push(to); }
      }
    }
    const path: string[] = [];
    for (let node: string | undefined = closing ?? first; node && node !== first; node = previous.get(node)) path.unshift(node);
    const cells = [first, ...path, first].map((id) => display.get(id)!);
    return { cells: cells.length > 12 ? [...cells.slice(0, 11), "…"] : cells, size: component.length };
  });
}
