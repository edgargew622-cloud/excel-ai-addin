export const EXCEL_MAX_ROWS = 1_048_576;
export const EXCEL_MAX_COLUMNS = 16_384;

export interface A1Rect {
  rowStart: number;
  rowEnd: number;
  columnStart: number;
  columnEnd: number;
  kind: "cells" | "rows" | "columns";
}

const CELL = /^\$?([A-Za-z]{1,3})\$?([1-9]\d{0,6})$/;
const ROWS = /^\$?([1-9]\d{0,6}):\$?([1-9]\d{0,6})$/;
const COLUMNS = /^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$/;
const DEFINED_NAME = /^[A-Za-z_\\][A-Za-z0-9_.\\]*$/;

export function columnNumber(label: string): number {
  let result = 0;
  for (const character of label.toUpperCase()) {
    result = result * 26 + character.charCodeAt(0) - 64;
  }
  return result;
}

function checkedCell(value: string): { row: number; column: number } | null {
  const match = CELL.exec(value);
  if (!match) return null;
  const column = columnNumber(match[1]);
  const row = Number(match[2]);
  if (column > EXCEL_MAX_COLUMNS || row > EXCEL_MAX_ROWS) return null;
  return { row, column };
}

export function parseA1Rect(input: string): A1Rect | null {
  const value = input.trim();
  const rowMatch = ROWS.exec(value);
  if (rowMatch) {
    const first = Number(rowMatch[1]);
    const last = Number(rowMatch[2]);
    if (first > last || last > EXCEL_MAX_ROWS) return null;
    return { rowStart: first, rowEnd: last, columnStart: 1, columnEnd: EXCEL_MAX_COLUMNS, kind: "rows" };
  }
  const columnMatch = COLUMNS.exec(value);
  if (columnMatch) {
    const first = columnNumber(columnMatch[1]);
    const last = columnNumber(columnMatch[2]);
    if (first > last || last > EXCEL_MAX_COLUMNS) return null;
    return { rowStart: 1, rowEnd: EXCEL_MAX_ROWS, columnStart: first, columnEnd: last, kind: "columns" };
  }
  const parts = value.split(":");
  if (parts.length > 2) return null;
  const first = checkedCell(parts[0]);
  const last = checkedCell(parts[1] ?? parts[0]);
  if (!first || !last || first.row > last.row || first.column > last.column) return null;
  return {
    rowStart: first.row,
    rowEnd: last.row,
    columnStart: first.column,
    columnEnd: last.column,
    kind: "cells"
  };
}

export function isDefinedName(value: string): boolean {
  const trimmed = value.trim();
  return DEFINED_NAME.test(trimmed) && !checkedCell(trimmed);
}

export function assertRangeReference(value: unknown): string {
  if (typeof value !== "string") throw new Error(`Адрес "${String(value)}" должен быть строкой.`);
  const trimmed = value.trim();
  if (!parseA1Rect(trimmed) && !isDefinedName(trimmed)) {
    throw new Error(`Адрес "${trimmed}" некорректен. Используйте A1, A1:B20, A:A, 1:10 или имя диапазона.`);
  }
  return trimmed;
}

export function cellCount(rect: A1Rect): number {
  return (rect.rowEnd - rect.rowStart + 1) * (rect.columnEnd - rect.columnStart + 1);
}

export function intersects(a: A1Rect, b: A1Rect): boolean {
  return a.rowStart <= b.rowEnd && b.rowStart <= a.rowEnd && a.columnStart <= b.columnEnd && b.columnStart <= a.columnEnd;
}

export function contains(outer: A1Rect, inner: A1Rect): boolean {
  return outer.rowStart <= inner.rowStart && outer.rowEnd >= inner.rowEnd &&
    outer.columnStart <= inner.columnStart && outer.columnEnd >= inner.columnEnd;
}
