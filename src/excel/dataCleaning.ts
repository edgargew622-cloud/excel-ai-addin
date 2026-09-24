/**
 * Очистка данных (этап 7, срез 7.2): профиль, а затем операции очистки.
 *
 * Профиль только читает. Большая таблица читается порциями по строкам,
 * и ответ называет проверенную область: если прочитано не всё, это сказано
 * прямо, а вывод по части не выдаётся за вывод по всей таблице.
 */

import { cultureDateOrder, profileData, type ColumnProfile, type DataProfile, type DateOrder, type NumberCulture } from "./dataProfile";
import { MAX_IO_CELLS, checkAddress, rangeOf, ToolError } from "./excelTools";
import { columnLetters } from "./formulaFill";
import { captureTarget, officeCapabilities } from "./workbookContext";

/** Больше этого профиль не читает: ответ был бы огромным, а время — долгим. */
export const MAX_PROFILE_CELLS = 100_000;

export type WorkbookCulture = NumberCulture & { dateOrder: DateOrder | null; dateSeparator: string; name: string };

/** Разделители и порядок даты книги — как их сообщает Excel (ExcelApi 1.12). */
export async function readCulture(ctx: Excel.RequestContext): Promise<WorkbookCulture> {
  const culture = ctx.workbook.application.cultureInfo;
  culture.load("name");
  culture.numberFormat.load(["numberDecimalSeparator", "numberGroupSeparator"]);
  culture.datetimeFormat.load(["dateSeparator", "shortDatePattern"]);
  await ctx.sync();
  return {
    name: String(culture.name),
    decimal: String(culture.numberFormat.numberDecimalSeparator),
    group: String(culture.numberFormat.numberGroupSeparator),
    dateSeparator: String(culture.datetimeFormat.dateSeparator),
    dateOrder: cultureDateOrder(String(culture.datetimeFormat.shortDatePattern))
  };
}

const FINDINGS: (keyof ColumnProfile)[] = [
  "edgeSpaces", "innerSpaces", "nonBreakingSpaces",
  "numbersAsText", "foreignNumbersAsText", "ambiguousNumbersAsText", "codesWithLeadingZeros",
  "datesAsText", "ambiguousDatesAsText"
];

/** Ответ модели: пустые находки опускаются, чтобы не тратить её контекст. */
function compact(profile: DataProfile) {
  return profile.columns.map((column) => {
    const findings = Object.fromEntries(FINDINGS
      .map((key) => [key, column[key]] as const)
      .filter(([, value]) => (value as { count: number }).count > 0));
    const { edgeSpaces, innerSpaces, nonBreakingSpaces, numbersAsText, foreignNumbersAsText, ambiguousNumbersAsText,
      codesWithLeadingZeros, datesAsText, ambiguousDatesAsText, ...counts } = column;
    return { ...counts, ...(Object.keys(findings).length ? { findings } : {}) };
  });
}

export async function profileRange(args: { sheet?: string; address?: string; hasHeaders?: boolean }) {
  const target = await captureTarget(args.sheet);
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(target.sheetId);
    sheet.load("name");
    let area: Excel.Range;
    if (args.address?.trim()) {
      area = await rangeOf(ctx, sheet, checkAddress(args.address));
    } else {
      const used = officeCapabilities().usedRangeOrNull ? sheet.getUsedRangeOrNullObject(true) : sheet.getUsedRange(true);
      used.load("isNullObject");
      await ctx.sync();
      if ((used as any).isNullObject) return { sheet: sheet.name, empty: true, note: "Лист пуст: профилировать нечего." };
      area = used;
    }
    area.load(["address", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
    await ctx.sync();
    const culture = await readCulture(ctx);

    // Порции по строкам, не больше MAX_IO_CELLS ячеек за чтение.
    const rowsPerChunk = Math.max(1, Math.floor(MAX_IO_CELLS / area.columnCount));
    const readableRows = Math.min(area.rowCount, Math.floor(MAX_PROFILE_CELLS / area.columnCount));
    if (readableRows < 1) throw new ToolError(`В ${area.address} слишком много столбцов для профиля: ${area.columnCount}.`);
    const values: unknown[][] = [];
    const formulas: unknown[][] = [];
    const valueTypes: unknown[][] = [];
    const numberFormat: unknown[][] = [];
    for (let start = 0; start < readableRows; start += rowsPerChunk) {
      const rows = Math.min(rowsPerChunk, readableRows - start);
      const chunk = sheet.getRangeByIndexes(area.rowIndex + start, area.columnIndex, rows, area.columnCount);
      chunk.load(["values", "formulas", "valueTypes", "numberFormat"]);
      await ctx.sync();
      values.push(...(chunk.values as unknown[][]));
      formulas.push(...(chunk.formulas as unknown[][]));
      valueTypes.push(...(chunk.valueTypes as unknown[][]));
      numberFormat.push(...(chunk.numberFormat as unknown[][]));
    }
    const lastColumn = columnLetters(area.columnIndex + area.columnCount);
    const checked = `${columnLetters(area.columnIndex + 1)}${area.rowIndex + 1}:${lastColumn}${area.rowIndex + readableRows}`;
    const incomplete = readableRows < area.rowCount;

    const profile = profileData({
      values, formulas, valueTypes, numberFormat,
      hasHeaders: args.hasHeaders !== false,
      origin: { rowIndex: area.rowIndex, columnIndex: area.columnIndex },
      address: checked,
      culture,
      columnName: columnLetters
    });

    return {
      sheet: sheet.name,
      address: area.address.slice(area.address.lastIndexOf("!") + 1),
      checkedAddress: checked,
      incomplete,
      ...(incomplete
        ? { incompleteNote: `Прочитаны первые ${readableRows} строк из ${area.rowCount}: профиль ограничен ${MAX_PROFILE_CELLS} ячейками. Выводы — только о проверенной области.` }
        : {}),
      hasHeaders: profile.hasHeaders,
      rows: profile.rows,
      emptyRows: profile.emptyRows,
      culture: { name: culture.name, decimal: culture.decimal, group: culture.group === " " ? "пробел" : culture.group, dateOrder: culture.dateOrder },
      columns: compact(profile),
      duplicateRows: profile.duplicateRows,
      duplicateRowsNormalized: profile.duplicateRowsNormalized,
      note:
        "Счётчики посчитаны панелью по прочитанным значениям. numbersAsText и datesAsText — однозначные числа и даты, " +
        "записанные текстом: в расчёт они не входят. ambiguous* и foreignNumbersAsText без ответа пользователя не преобразуются: " +
        "01.02.2026 может быть и 1 февраля, и 2 января, 1,500 — и полутора, и тысячей пятьсот. codesWithLeadingZeros — коды, " +
        "их нельзя превращать в числа. duplicateRowsNormalized — кандидаты после обрезки пробелов и без учёта регистра, не доказанные дубликаты."
    };
  });
}
