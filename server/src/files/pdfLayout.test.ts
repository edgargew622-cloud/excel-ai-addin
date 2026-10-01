import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { parsePdf } from "./pdf.js";
import { buildSheetLayout, cellValue, normalizeSegments, type SheetLayout } from "./pdfLayout.js";

// Образцы — fixtures/layout-*.pdf, их делает fixtures/make-layout-pdfs.py (fpdf2).
const layoutOf = async (name: string): Promise<SheetLayout> => {
  const parsed = await parsePdf(readFileSync(join(process.cwd(), "src", "files", "fixtures", name)), name);
  assert.equal(parsed.layoutError, undefined);
  return parsed.layout!;
};
const at = (layout: SheetLayout, text: string) => layout.cells.find((cell) => cell.text === text);
const row = (layout: SheetLayout, r: number) => layout.cells.filter((cell) => cell.r === r).sort((p, q) => p.c - q.c).map((cell) => cell.text);

test("values: numbers and dates as numbers, codes and times as text", () => {
  assert.deepEqual(cellValue("12 500,00"), { value: 12500, numberFormat: "#,##0.00" });
  assert.deepEqual(cellValue("1 500"), { value: 1500, numberFormat: "#,##0" });
  assert.deepEqual(cellValue("17"), { value: 17, numberFormat: "General" });
  assert.deepEqual(cellValue("01.09.2026"), { value: 46266, numberFormat: "dd.mm.yyyy" });
  for (const text of ["0042", "7714636130", "22:04", "31.02.2026", "+7 (495) 940-84-18", "12,5%"]) {
    assert.equal(cellValue(text).numberFormat, "@", text);
  }
});

test("touching and overlapping pieces of one line become one segment", () => {
  const { h, v } = normalizeSegments([
    { x1: 10, y1: 20, x2: 50, y2: 20, width: 0.5 },
    { x1: 50, y1: 20.4, x2: 90, y2: 20.4, width: 1.5 },
    { x1: 30, y1: 10, x2: 30, y2: 60, width: 0.5 },
    { x1: 0, y1: 0, x2: 2, y2: 0, width: 0.5 }
  ]);
  assert.equal(h.length, 1);
  assert.deepEqual([Math.round(h[0].x1), Math.round(h[0].x2), h[0].width], [10, 90, 1.5]);
  assert.equal(v.length, 1);
});

test("a ruled table: same columns as the PDF, one sheet row per table row, merged total, borders", async () => {
  const layout = await layoutOf("layout-grid.pdf");
  assert.deepEqual(layout.columnWidths.map(Math.round), [30, 75, 120, 200, 90]);
  const header = at(layout, "№")!;
  assert.deepEqual(row(layout, header.r), ["№", "Дата", "Контрагент", "Назначение платежа", "Сумма, руб."]);
  assert.equal(header.bold, true);
  const first = row(layout, header.r + 1);
  assert.deepEqual(first, ["1", "01.09.2026", "ООО «Альфа»", "Оплата по счёту 15", "12 500,00"]);
  assert.equal(at(layout, "12 500,00")!.value, 12500);
  assert.equal(at(layout, "12 500,00")!.align, "Right");
  const total = at(layout, "Итого")!;
  assert.deepEqual([total.colSpan, total.rowSpan, total.align], [4, 1, "Right"]);
  const title = at(layout, "Реестр платежей за сентябрь 2026")!;
  assert.deepEqual([title.colSpan, title.align, title.bold], [5, "Center", true]);
  // Строка таблицы — одна строка листа высотой как в PDF, а не три тонкие.
  assert.equal(Math.round(layout.rowHeights[header.r]), 20);
  assert.ok(layout.hEdges.length >= 30 && layout.vEdges.length >= 30);
});

test("a form: cells merged inside one frame, multi-line cell kept whole, empty fields kept", async () => {
  const layout = await layoutOf("layout-form.pdf");
  assert.equal(at(layout, "Сведения о транспортном средстве")!.colSpan, layout.columnWidths.length);
  assert.equal(at(layout, "SOLARIS")!.colSpan, 5);
  const fio = layout.cells.find((cell) => cell.text.startsWith("Иванов"))!;
  assert.equal(fio.text, "Иванов Иван Иванович,\nудостоверение 99 00 000000");
  assert.equal(fio.wrap, true);
  assert.equal(at(layout, "Выезд")!.align, "Center");
  assert.equal(at(layout, "114 106")!.value, 114106);
  assert.equal(at(layout, "08:00")!.numberFormat, "@");
  // Пустые поля бланка («Время» возвращения) остаются объединёнными ячейками.
  assert.ok(layout.cells.some((cell) => !cell.text && cell.colSpan > 1));
});

test("a table without lines: columns from the gaps, even right next to a right-aligned number", async () => {
  const layout = await layoutOf("layout-plain.pdf");
  const header = at(layout, "Товар")!;
  assert.deepEqual(row(layout, header.r), ["Товар", "Кол-во", "Ед.", "Цена", "Сумма"]);
  assert.deepEqual(row(layout, header.r + 1), ["Бумага A4", "120", "пач.", "310,00", "37 200,00"]);
  assert.equal(at(layout, "1 500")!.value, 1500);
  assert.equal(at(layout, "шт.")!.numberFormat, "@");
  assert.equal(layout.hEdges.length + layout.vEdges.length, 0);
});

test("two pages: pages one after another, the same columns, the repeated header kept", async () => {
  const layout = await layoutOf("layout-multipage.pdf");
  assert.equal(layout.pages, 2);
  assert.equal(layout.columnWidths.length, 5);
  assert.equal(layout.cells.filter((cell) => cell.text === "Дата").length, 2);
  assert.equal(layout.cells.filter((cell) => /^Операция \d+$/.test(cell.text)).length, 70);
  assert.equal(at(layout, "Операция 70")!.r > layout.pageStarts[1], true);
});

test("no text at all is a scan, not an empty copy", () => {
  assert.throws(() => buildSheetLayout([{ width: 100, height: 100, texts: [], segments: [] }]), /скан/);
});

test("text longer than its cell stays left-aligned so it runs on to the right, as in the PDF", async () => {
  const parsed = await parsePdf(readFileSync(join(process.cwd(), "src", "files", "fixtures", "layout-grid.pdf")), "g.pdf");
  // Живая проверка 01.10.2026: строка шире столбца A вышла по правому краю — видно было «000001».
  assert.equal(parsed.layout!.cells.find((cell) => cell.text.startsWith("Организация"))!.align, "Left");
});
