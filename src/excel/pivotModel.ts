/**
 * Сводная таблица, посчитанная панелью до того, как её построит Excel.
 *
 * Прежний `create_pivot_table` ничего не знал о результате: сводная могла
 * лечь поверх данных (её размер заранее неизвестен), а текстовый столбец
 * в значениях тихо превращал «сумму» в «количество». Здесь по исходным
 * данным считается то же, что посчитает Excel: группы, итоги по ним, общий
 * итог и размер сводной. Размер нужен, чтобы проверить место до операции,
 * итоги — чтобы сверить построенное после.
 *
 * Модуль не обращается к Excel и проверяется без него.
 */

export type Aggregation = "sum" | "count" | "average" | "max" | "min";

export const AGGREGATIONS: readonly Aggregation[] = ["sum", "count", "average", "max", "min"];

/** Имена функций Office.js. */
export const OFFICE_AGGREGATION: Record<Aggregation, string> = {
  sum: "Sum",
  count: "Count",
  average: "Average",
  max: "Max",
  min: "Min"
};

export const AGGREGATION_TEXT: Record<Aggregation, string> = {
  sum: "сумма",
  count: "количество",
  average: "среднее",
  max: "максимум",
  min: "минимум"
};

export interface PivotValueField {
  field: string;
  aggregation: Aggregation;
}

/** Группа сводной на любом уровне: путь подписей от первого поля строк. */
export interface PivotNode {
  /** Подписи, как их покажет Excel; пустое значение — «(пусто)». */
  path: string[];
  /** Ключи сравнения: по ним Excel сводит элементы в один. */
  keys: string[];
  /** Итог группы по каждому полю значений (по всем столбцам). */
  totals: number[];
  /** С полем в столбцах: итоги группы по каждому элементу столбцов. Нет ключа — нет данных, Excel оставит ячейку пустой. */
  cells?: Record<string, number[]>;
}

/**
 * Фильтр поля сводной (этап 8, 8.2). include — оставить только эти элементы;
 * top/bottom — первые или последние N элементов по полю значений `by`.
 */
export type PivotFilter =
  | { field: string; include: string[] }
  | { field: string; top: number; by: number }
  | { field: string; bottom: number; by: number };

/** Сортировка элементов поля строк по полю значений `by` (индекс в values). */
export interface PivotSort {
  field: string;
  by: number;
  order: "asc" | "desc";
}

export interface PivotOptions {
  /** Поле в столбцах — одно (8.2). */
  columnField?: string;
  filters?: readonly PivotFilter[];
  sort?: PivotSort;
}

export interface PivotExpectation {
  /** Индексы столбцов источника для строк и значений. */
  rowColumns: number[];
  valueColumns: number[];
  /** Индекс столбца источника для поля в столбцах; -1 — поля нет. */
  columnColumn: number;
  /** Элементы поля в столбцах: ключи и подписи. */
  columns: { key: string; label: string }[];
  /** Строк шапки: 1 без поля в столбцах, 2 с ним, 3 — с ним и несколькими полями значений (замер 27.09.2026). */
  headerRows: number;
  /** Итоги по элементам первого поля строк — для предпросмотра. */
  groups: { label: string; totals: number[] }[];
  /** Итоги всех групп всех уровней: по ним сверяется построенная сводная. */
  nodes: PivotNode[];
  /** Общий итог по каждому полю значений. */
  grandTotals: number[];
  /** С полем в столбцах: общий итог по каждому элементу столбцов. */
  grandCells?: Record<string, number[]>;
  /** Сколько строк и столбцов займёт сводная. */
  height: number;
  width: number;
  /**
   * Размер до фильтров. Пока поля и фильтры добавляются по очереди, сводная
   * бывает больше итоговой, поэтому пустым должно быть и это место.
   */
  fullHeight: number;
  fullWidth: number;
  /** Сколько строк источника с пустой подписью первого поля: Excel соберёт их в «(пусто)». */
  blankLabels: number;
  /** Сортировка, которую надо проверить в построенной сводной: уровень поля строк. */
  sort?: { level: number; by: number; order: "asc" | "desc" };
  /** Элементы, отброшенные фильтрами, — для предпросмотра. */
  filteredOut: { field: string; items: string[] }[];
  warnings: string[];
}

const isNumber = (value: unknown): value is number => typeof value === "number" && Number.isFinite(value);
const isBlank = (value: unknown) => value === "" || value === null || value === undefined;

/**
 * Проблемы шапки, при которых сводную не построить или построить неверно.
 *
 * Excel требует у каждого столбца источника непустое имя поля. Повтор имён
 * Excel примет, переименовав второе, но тогда поле по имени найдётся не то.
 */
export function pivotHeaderProblems(headers: readonly unknown[]): string[] {
  const problems: string[] = [];
  const seen = new Map<string, number>();
  headers.forEach((value, index) => {
    const text = isBlank(value) ? "" : String(value).trim();
    if (!text) {
      problems.push(`столбец ${index + 1} без заголовка — Excel не построит сводную по такой области`);
      return;
    }
    const key = text.toLowerCase();
    const first = seen.get(key);
    if (first !== undefined) problems.push(`заголовок «${text}» повторяется в столбцах ${first + 1} и ${index + 1}`);
    else seen.set(key, index);
  });
  return problems;
}

/** Поле по имени — без учёта регистра и пробелов по краям, как ищет Excel. */
export function fieldIndex(headers: readonly unknown[], name: string): number {
  const wanted = name.trim().toLowerCase();
  return headers.findIndex((value) => String(value ?? "").trim().toLowerCase() === wanted);
}

function aggregate(values: readonly unknown[], aggregation: Aggregation): number {
  if (aggregation === "count") return values.filter((value) => !isBlank(value)).length;
  const numbers = values.filter(isNumber);
  if (!numbers.length) return 0;
  switch (aggregation) {
    case "sum": return numbers.reduce((total, value) => total + value, 0);
    case "average": return numbers.reduce((total, value) => total + value, 0) / numbers.length;
    case "max": return Math.max(...numbers);
    case "min": return Math.min(...numbers);
  }
  return 0;
}

/** Подпись элемента так, как Excel покажет её в сводной.
 *
 * Пробелы по краям не срезаются: замер в Excel 24 сентября 2026 года —
 * «Москва » с пробелом стала отдельным элементом рядом с «Москва». */
export function pivotLabel(value: unknown): string {
  if (typeof value === "boolean") return value ? "ИСТИНА" : "ЛОЖЬ";
  return String(value);
}

const BLANK_KEY = "\u0000пусто";
const BLANK_TEXT = "(пусто)";

/**
 * Ключ, по которому Excel сводит значения в один элемент.
 *
 * Замер в Excel 24 сентября 2026 года: регистр не различается («москва»
 * и «Москва» — один элемент), число 1 и текст «1» — один элемент, пробел
 * в конце — различается. Поэтому ключ — текст подписи в нижнем регистре,
 * без учёта типа и без обрезки.
 */
export function pivotKey(value: unknown): string {
  return isBlank(value) ? BLANK_KEY : pivotLabel(value).toLowerCase();
}

/** Ключ подписи, прочитанной из сводной: пустой элемент Excel называет
 * на языке интерфейса. Значение «(пусто)» в самих данных с ним совпадёт —
 * такой случай не различается. */
function layoutKey(value: unknown): string {
  const text = pivotLabel(value);
  return /^\((пусто|blank)\)$/i.test(text.trim()) ? BLANK_KEY : text.toLowerCase();
}

const pathKey = (keys: readonly string[]) => keys.join("\u0001");
const pathText = (path: readonly string[]) => path.join(" › ");

/** Элементы поля в строках источника: ключ → подпись и строки. */
function itemsOf(rows: readonly (readonly unknown[])[], column: number) {
  const items = new Map<string, { label: string; rows: (readonly unknown[])[] }>();
  for (const row of rows) {
    const key = pivotKey(row[column]);
    if (!items.has(key)) items.set(key, { label: isBlank(row[column]) ? BLANK_TEXT : pivotLabel(row[column]), rows: [] });
    items.get(key)!.rows.push(row);
  }
  return items;
}

/** Сколько разных групп всех уровней в строках — для размера сводной. */
function nodeCount(rows: readonly (readonly unknown[])[], rowColumns: readonly number[]): number {
  let count = 0;
  for (let level = 1; level <= rowColumns.length; level++) {
    count += new Set(rows.map((row) => pathKey(rowColumns.slice(0, level).map((column) => pivotKey(row[column]))))).size;
  }
  return count;
}

/**
 * Что построит Excel в табличном макете с итогами внизу групп: шапка,
 * по строке на каждую группу последнего уровня, строка итога под каждой
 * группой остальных уровней, строка общего итога. По столбцу на каждое
 * поле строк и на каждое поле значений.
 *
 * Табличный, а не компактный: в компактном все уровни пишутся в один
 * столбец, и одинаковая подпись под разными родителями неотличима —
 * проверить вложенные итоги по нему нельзя (план стабилизации, S3.1).
 *
 * С полем в столбцах (8.2, замер 27.09.2026): под каждым его элементом —
 * по столбцу на каждое поле значений, в конце — итоговые столбцы по каждому
 * полю значений; шапка в две строки, а с несколькими полями значений — в три.
 * Сочетания, которого нет в данных, Excel оставляет пустой ячейкой, не нулём.
 *
 * Фильтры (8.2): сначала выбор элементов, затем первые/последние N — среди
 * уже выбранного. Общий итог Excel считает только по оставшемуся: замер
 * 27.09.2026 — после «первых 2» из трёх городов итог 438 вместо 473.
 */
export function expectPivot(
  values: readonly (readonly unknown[])[],
  rowFields: readonly string[],
  valueFields: readonly PivotValueField[],
  options: PivotOptions = {}
): PivotExpectation {
  const headers = values[0] ?? [];
  const allRows = values.slice(1).filter((row) => row.some((cell) => !isBlank(cell)));
  const warnings: string[] = [];

  const rowColumns = rowFields.map((name) => fieldIndex(headers, name));
  const valueColumns = valueFields.map((item) => fieldIndex(headers, item.field));
  const columnColumn = options.columnField ? fieldIndex(headers, options.columnField) : -1;
  const hasColumns = columnColumn >= 0;
  const valueCount = valueFields.length;
  const totalsOf = (rows: readonly (readonly unknown[])[]) =>
    valueFields.map((item, index) => aggregate(rows.map((row) => row[valueColumns[index]]), item.aggregation));
  const cellsOf = (rows: readonly (readonly unknown[])[]) => {
    const cells: Record<string, number[]> = {};
    for (const [key, item] of itemsOf(rows, columnColumn)) cells[key] = totalsOf(item.rows);
    return cells;
  };

  let body = allRows;
  const filteredOut: { field: string; items: string[] }[] = [];
  const filters = options.filters ?? [];
  for (const filter of filters) {
    if (!("include" in filter)) continue;
    const column = fieldIndex(headers, filter.field);
    const wanted = new Set(filter.include.map(pivotKey));
    const dropped = [...itemsOf(body, column).entries()].filter(([key]) => !wanted.has(key)).map(([, item]) => item.label);
    body = body.filter((row) => wanted.has(pivotKey(row[column])));
    if (dropped.length) filteredOut.push({ field: filter.field, items: dropped });
  }
  for (const filter of filters) {
    if ("include" in filter) continue;
    const column = fieldIndex(headers, filter.field);
    const top = "top" in filter;
    const count = top ? filter.top : filter.bottom;
    const ranked = [...itemsOf(body, column).entries()].map(([key, item]) => ({
      key,
      label: item.label,
      value: aggregate(item.rows.map((row) => row[valueColumns[filter.by]]), valueFields[filter.by].aggregation)
    }));
    ranked.sort((x, y) => (top ? y.value - x.value : x.value - y.value));
    const kept = ranked.slice(0, count);
    if (ranked.length > count && kept.length && ranked[count].value === kept[kept.length - 1].value) {
      warnings.push(
        `В поле «${filter.field}» на границе ${top ? "первых" : "последних"} ${count} равные значения (${kept[kept.length - 1].value}): ` +
        "Excel может оставить больше элементов — сверка это покажет."
      );
    }
    const keep = new Set(kept.map((item) => item.key));
    const dropped = ranked.filter((item) => !keep.has(item.key)).map((item) => item.label);
    body = body.filter((row) => keep.has(pivotKey(row[column])));
    if (dropped.length) filteredOut.push({ field: filter.field, items: dropped });
  }

  // Группы всех уровней: сочетания ключей от первого поля до этого уровня.
  // По строке на каждую: листья последнего уровня и итоги остальных.
  const nodes: PivotNode[] = [];
  for (let level = 1; level <= rowColumns.length; level++) {
    const byPath = new Map<string, { path: string[]; keys: string[]; rows: (readonly unknown[])[] }>();
    for (const row of body) {
      const cells = rowColumns.slice(0, level).map((column) => row[column]);
      const keys = cells.map(pivotKey);
      const key = pathKey(keys);
      if (!byPath.has(key)) byPath.set(key, { path: cells.map((cell) => (isBlank(cell) ? BLANK_TEXT : pivotLabel(cell))), keys, rows: [] });
      byPath.get(key)!.rows.push(row);
    }
    for (const group of byPath.values()) {
      nodes.push({
        path: group.path,
        keys: group.keys,
        totals: totalsOf(group.rows),
        ...(hasColumns ? { cells: cellsOf(group.rows) } : {})
      });
    }
  }
  const columns = hasColumns ? [...itemsOf(body, columnColumn).entries()].map(([key, item]) => ({ key, label: item.label })) : [];
  const headerRows = hasColumns ? (valueCount > 1 ? 3 : 2) : 1;
  const widthFor = (columnItems: number) => rowColumns.length + (hasColumns ? (columnItems + 1) * valueCount : valueCount);
  const height = headerRows + nodes.length + 1;
  const width = widthFor(columns.length);
  const fullHeight = headerRows + nodeCount(allRows, rowColumns) + 1;
  const fullWidth = widthFor(hasColumns ? itemsOf(allRows, columnColumn).size : 0);

  const first = rowColumns[0];
  const blankLabels = body.filter((row) => isBlank(row[first])).length;
  const groups = nodes
    .filter((node) => node.keys.length === 1 && node.keys[0] !== BLANK_KEY)
    .map((node) => ({ label: node.path[0], totals: node.totals }));
  const grandTotals = totalsOf(body);

  valueFields.forEach((item, index) => {
    const cells = body.map((row) => row[valueColumns[index]]).filter((value) => !isBlank(value));
    const text = cells.filter((value) => !isNumber(value)).length;
    if (item.aggregation !== "count" && text > 0) {
      warnings.push(
        cells.length === text
          ? `В поле «${item.field}» нет чисел: ${AGGREGATION_TEXT[item.aggregation]} будет нулём. Для текста подходит только количество.`
          : `В поле «${item.field}» ${text} нечисловых значений из ${cells.length}: в ${AGGREGATION_TEXT[item.aggregation]} они не войдут.`
      );
    }
  });
  if (blankLabels > 0) {
    warnings.push(`В поле «${rowFields[0]}» ${blankLabels} пустых значений: Excel соберёт их в отдельную строку «(пусто)».`);
  }
  const firstLevel = nodes.filter((node) => node.keys.length === 1).length;
  if (firstLevel > 200) {
    warnings.push(`В поле «${rowFields[0]}» ${firstLevel} разных значений — сводная выйдет длиной в ${height} строк и почти ничего не сведёт.`);
  }
  if (columns.length > 30) {
    warnings.push(`В поле «${options.columnField}» ${columns.length} разных значений — сводная выйдет шириной в ${width} столбцов.`);
  }

  const sortLevel = options.sort ? rowColumns.indexOf(fieldIndex(headers, options.sort.field)) : -1;
  return {
    rowColumns, valueColumns, columnColumn, columns, headerRows, groups, nodes, grandTotals,
    ...(hasColumns ? { grandCells: cellsOf(body) } : {}),
    height, width, fullHeight, fullWidth, blankLabels,
    ...(options.sort && sortLevel >= 0 ? { sort: { level: sortLevel, by: options.sort.by, order: options.sort.order } } : {}),
    filteredOut, warnings
  };
}

/** Числа сводной сравниваются с допуском: среднее и суммы дробей Excel округляет по-своему. */
export function samePivotNumber(actual: unknown, expected: number): boolean {
  if (!isNumber(actual)) return false;
  const scale = Math.max(1, Math.abs(expected));
  return Math.abs(actual - expected) <= 1e-9 * scale + 1e-6;
}

/**
 * Сверяет прочитанную из Excel сводную с расчётом панели.
 *
 * Макет табличный, итоги внизу групп — так его выставляет исполнитель.
 * Первые столбцы — подписи полей строк, остальные — поля значений.
 * Пустая подпись родителя значит «та же группа, что строкой выше», поэтому
 * полный путь строки восстанавливается переносом подписи вниз. Строка,
 * у которой заполнены не все столбцы подписей, — итог группы того уровня,
 * где стоит последняя подпись: он относится к группе, после которой стоит,
 * и его подпись («Москва Итог») обязана эту группу называть. Слово «Итог»
 * Excel пишет на языке интерфейса, поэтому сравнивается только имя группы,
 * а общий итог узнаётся по месту — последняя строка.
 *
 * С полем в столбцах элементы столбцов читаются из второй строки шапки и
 * сопоставляются по ключу, а не по порядку: порядок — дело сортировки Excel.
 * Пустая ячейка там, где по расчёту данных нет, — норма; число там же — нет.
 *
 * Сверяются итоги всех групп всех уровней, общий итог, отсутствие строк
 * и столбцов, которых по расчёту быть не должно, и — если просили —
 * порядок элементов по значению.
 */
export function pivotMismatches(expected: PivotExpectation, layout: readonly (readonly unknown[])[]): string[] {
  const problems: string[] = [];
  if (layout.length !== expected.height) problems.push(`строк ${layout.length} вместо ${expected.height}`);
  const width = layout[0]?.length ?? 0;
  if (width !== expected.width) problems.push(`столбцов ${width} вместо ${expected.width}`);
  if (problems.length) return problems;

  const levels = expected.rowColumns.length;
  const valueCount = expected.valueColumns.length;
  const hasColumns = expected.columnColumn >= 0;

  // Каждый столбец значений: элемент поля в столбцах (null — итоговый) и поле значений.
  const slots: { key: string | null; value: number }[] = [];
  if (hasColumns) {
    const labels = layout[1];
    const seen = new Set<string>();
    for (let group = 0; group <= expected.columns.length; group++) {
      let key: string | null = null;
      if (group < expected.columns.length) {
        const cell = labels[levels + group * valueCount];
        key = layoutKey(cell);
        if (!expected.columns.some((item) => item.key === key)) problems.push(`лишний столбец «${pivotLabel(cell)}»`);
        seen.add(key);
      }
      for (let value = 0; value < valueCount; value++) slots.push({ key, value });
    }
    for (const item of expected.columns) if (!seen.has(item.key)) problems.push(`нет столбца «${item.label}»`);
    if (problems.length) return problems;
  } else {
    for (let value = 0; value < valueCount; value++) slots.push({ key: null, value });
  }

  const slotName = (slot: { key: string | null; value: number }) => {
    if (!hasColumns) return valueCount > 1 ? `, столбец ${levels + slot.value + 1}` : "";
    const column = slot.key === null ? "итог" : expected.columns.find((item) => item.key === slot.key)!.label;
    return ` (${column}${valueCount > 1 ? `, поле значений ${slot.value + 1}` : ""})`;
  };
  const compare = (where: string, actual: readonly unknown[], totals: readonly number[], cells?: Record<string, number[]>) => {
    slots.forEach((slot, index) => {
      const expectedValue = slot.key === null ? totals[slot.value] : cells?.[slot.key]?.[slot.value];
      if (expectedValue === undefined) {
        if (!isBlank(actual[index])) problems.push(`${where}${slotName(slot)}: ${String(actual[index])} вместо пустой ячейки`);
        return;
      }
      if (!samePivotNumber(actual[index], expectedValue)) problems.push(`${where}${slotName(slot)}: ${String(actual[index])} вместо ${expectedValue}`);
    });
  };

  const total = layout[layout.length - 1];
  compare("общий итог", total.slice(levels), expected.grandTotals, expected.grandCells);

  // Путь каждой строки: ключи по уровням, перенесённые сверху.
  const actual = new Map<string, { path: string[]; values: readonly unknown[] }>();
  const order: string[][] = [];
  const keys: string[] = [];
  const texts: string[] = [];
  for (let r = expected.headerRows; r < layout.length - 1; r++) {
    const row = layout[r];
    const labels = row.slice(0, levels);
    let deepest = -1;
    labels.forEach((cell, index) => { if (!isBlank(cell)) deepest = index; });
    if (deepest < 0) {
      problems.push(`строка ${r + 1} сводной без подписи`);
      continue;
    }
    if (deepest === levels - 1) {
      labels.forEach((cell, index) => {
        if (isBlank(cell)) return;
        keys[index] = layoutKey(cell);
        texts[index] = pivotLabel(cell);
      });
    } else {
      // Итог группы уровня deepest: её имя — последняя подпись этого уровня.
      const text = pivotLabel(labels[deepest]);
      const group = texts[deepest];
      if (group === undefined || !text.toLowerCase().includes(group.toLowerCase())) {
        problems.push(`строка итога «${text}» стоит после группы «${group ?? "—"}» и не называет её`);
        continue;
      }
    }
    const rowKeys = keys.slice(0, deepest + 1);
    if (rowKeys.length !== deepest + 1 || rowKeys.some((key) => key === undefined)) {
      problems.push(`строка ${r + 1} сводной: родительскую группу не восстановить`);
      continue;
    }
    actual.set(pathKey(rowKeys), { path: texts.slice(0, deepest + 1), values: row.slice(levels) });
    order.push(rowKeys);
  }

  const expectedKeys = new Map<string, PivotNode>();
  for (const node of expected.nodes) {
    const key = pathKey(node.keys);
    expectedKeys.set(key, node);
    const row = actual.get(key);
    if (!row) problems.push(`нет строки ${pathText(node.path)}`);
    else compare(pathText(node.path), row.values, node.totals, node.cells);
  }
  for (const [key, row] of actual) {
    if (!expectedKeys.has(key)) problems.push(`лишняя строка ${pathText(row.path)}`);
  }

  // Порядок по значению: внутри каждого родителя группы уровня идут по итогу
  // выбранного поля значений. Равные итоги могут стоять в любом порядке.
  if (expected.sort) {
    const { level, by, order: direction } = expected.sort;
    let previous: { parent: string; node: PivotNode } | null = null;
    for (const rowKeys of order) {
      if (rowKeys.length !== level + 1) continue;
      const node = expectedKeys.get(pathKey(rowKeys));
      if (!node) continue;
      const parent = pathKey(rowKeys.slice(0, level));
      if (previous && previous.parent === parent) {
        const before = previous.node.totals[by];
        const after = node.totals[by];
        if (direction === "desc" ? after > before : after < before) {
          problems.push(
            `порядок ${direction === "desc" ? "по убыванию" : "по возрастанию"} нарушен: «${pathText(previous.node.path)}» (${before}) стоит перед «${pathText(node.path)}» (${after})`
          );
        }
      }
      previous = { parent, node };
    }
  }
  return problems.slice(0, 8);
}

/* --- группировка дат через вспомогательный столбец (этап 8, 8.2) --------------- */

export type DateGrouping = "year" | "quarter" | "month";

/** От крупного к мелкому: в таком порядке уровни встают в сводную. */
export const DATE_GROUPINGS: readonly DateGrouping[] = ["year", "quarter", "month"];

export const DATE_GROUPING_TEXT: Record<DateGrouping, string> = { year: "год", quarter: "квартал", month: "месяц" };

/**
 * Подпись группы для серийного номера даты Excel — та же, что выдаст формула
 * вспомогательного столбца (`dateGroupFormula`). Замер 27.09.2026: формулы
 * дали «2026-02» и «2026 К1»; год — число.
 */
export function dateGroupLabel(serial: number, by: DateGrouping): string | number {
  const date = new Date(Date.UTC(1899, 11, 30) + Math.floor(serial) * 86400000);
  const year = date.getUTCFullYear();
  const month = date.getUTCMonth() + 1;
  if (by === "year") return year;
  if (by === "quarter") return `${year} К${Math.ceil(month / 3)}`;
  return `${year}-${String(month).padStart(2, "0")}`;
}

/**
 * Формула вспомогательного столбца. Пишется по-английски через `formulas` —
 * Excel сам переводит её на язык интерфейса (замер: ГОД, ТЕКСТ, МЕСЯЦ,
 * ОКРУГЛВВЕРХ). Формат "00" — одни цифры, от языка не зависит. Пустая дата —
 * пустая подпись, как у самой даты в сводной.
 */
export function dateGroupFormula(cell: string, by: DateGrouping): string {
  if (by === "year") return `=IF(${cell}="","",YEAR(${cell}))`;
  if (by === "quarter") return `=IF(${cell}="","",YEAR(${cell})&" К"&ROUNDUP(MONTH(${cell})/3,0))`;
  return `=IF(${cell}="","",YEAR(${cell})&"-"&TEXT(MONTH(${cell}),"00"))`;
}

/** Числовой формат даты: есть день, месяц или год, и это не «Общий». */
export function isDateFormat(format: unknown): boolean {
  const text = String(format ?? "").replace(/"[^"]*"/g, "").replace(/\[[^\]]*\]/g, "").trim().toLowerCase();
  if (!text || text === "general" || text === "основной" || text === "@") return false;
  return /[dmyдмг]/i.test(text);
}
