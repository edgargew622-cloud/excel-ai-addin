import test from "node:test";
import assert from "node:assert/strict";
import { executeRowOpPlan, prepareDeleteRowsPlan, prepareInsertRowsPlan } from "./excelTools";
import { PLANNED_TOOLS } from "./plans";

/**
 * Книга из двух листов: данные о продажах и отчёт, который на них ссылается.
 * Именно такая связка и ломается от удаления строк, причём ломается не там,
 * куда смотрит пользователь.
 */
function workbook(options: { protectedSheet?: boolean; deleteThrows?: boolean } = {}) {
  const sales = [
    ["Товар", "Цена", "Итог"],
    ["А", 10, "=B2*2"],
    ["Б", 20, "=B3*2"],
    ["В", 30, "=B4*2"],
    ["Г", 40, "=B5*2"]
  ];
  const report = [["Всего", "=SUM(Продажи!B2:B5)"], ["Первый", "=Продажи!B2"]];
  const grids: Record<string, any[][]> = { "Продажи": sales, "Отчёт": report };

  const letters = (index: number) => String.fromCharCode(64 + index);
  const evaluate = (grid: any[][], row: number, column: number) => {
    const raw = grid[row]?.[column];
    if (typeof raw !== "string" || !raw.startsWith("=")) return raw ?? "";
    // Достаточно одного правила: ссылка на исчезнувшую строку — ошибка.
    return /#REF!/.test(raw) ? "#REF!" : 1;
  };

  /**
   * Диапазон ведёт себя как настоящий: свойство, которое не запросили через
   * load, читать нельзя. Проверка 18 сентября 2026 года сорвалась именно на
   * этом — опрос объединений читал rowIndex у незагруженного диапазона.
   */
  function makeRange(name: string, rowIndex: number, columnIndex: number, rowCount: number, columnCount: number): any {
    const grid = grids[name];
    const loaded = new Set<string>();
    const slice = (reader: (row: number, column: number) => unknown) =>
      Array.from({ length: rowCount }, (_, row) =>
        Array.from({ length: columnCount }, (_, column) => reader(rowIndex + row, columnIndex + column)));
    const guard = <T>(property: string, value: () => T) => () => {
      if (!loaded.has(property)) {
        throw new Error(`Свойство "${property}" недоступно. Прежде чем прочесть его значение, вызовите метод загрузки.`);
      }
      return value();
    };
    const range: any = {
      isNullObject: false,
      load: (properties: string | string[]) => {
        for (const property of Array.isArray(properties) ? properties : String(properties).split(",")) {
          loaded.add(property.trim());
        }
      },
      // Объединений в книге нет, но опрос до них должен дойти без ошибки.
      getMergedAreasOrNullObject: () => ({
        isNullObject: true,
        areaCount: 0,
        areas: { items: [], load: () => undefined },
        load: () => undefined
      }),
      format: { protection: { locked: false, load: () => undefined } }
    };
    const properties: Record<string, () => unknown> = {
      rowIndex: () => rowIndex,
      columnIndex: () => columnIndex,
      rowCount: () => rowCount,
      columnCount: () => columnCount,
      address: () => `${name}!${letters(columnIndex + 1)}${rowIndex + 1}:${letters(columnIndex + columnCount)}${rowIndex + rowCount}`,
      formulas: () => slice((row, column) => grid[row]?.[column] ?? ""),
      values: () => slice((row, column) => evaluate(grid, row, column))
    };
    for (const [property, value] of Object.entries(properties)) {
      Object.defineProperty(range, property, { get: guard(property, value), enumerable: true });
    }
    return range;
  }

  function makeSheet(name: string, id: string): any {
    const grid = grids[name];
    return {
      id,
      name,
      load: () => undefined,
      protection: { protected: name === "Продажи" && options.protectedSheet === true, load: () => undefined },
      tables: { items: [], load: () => undefined },
      getUsedRangeOrNullObject: () => makeRange(name, 0, 0, grid.length, grid[0].length),
      getUsedRange: () => makeRange(name, 0, 0, grid.length, grid[0].length),
      getRangeByIndexes: (row: number, column: number, rows: number, columns: number) =>
        makeRange(name, row, column, rows, columns),
      getRange: (address: string) => {
        const rows = /^(\d+):(\d+)$/.exec(address);
        if (!rows) return makeRange(name, 0, 0, grid.length, grid[0].length);
        const first = Number(rows[1]) - 1;
        const count = Number(rows[2]) - Number(rows[1]) + 1;
        // Свойства диапазона — геттеры с проверкой загрузки, поэтому не копируем
        // их разворотом: он прочитал бы их прямо здесь.
        return Object.assign(makeRange(name, first, 0, count, grid[0].length), {
          insert: () => {
            grid.splice(first, 0, ...Array.from({ length: count }, () => grid[0].map(() => "")));
          },
          delete: () => {
            if (options.deleteThrows) throw new Error("Во время обработки запроса произошла внутренняя ошибка.");
            grid.splice(first, count);
            // Так Excel и ломает формулы: ссылка на исчезнувшую строку.
            for (const other of Object.values(grids)) {
              for (const row of other) {
                row.forEach((cell, index) => {
                  // Ссылку внутри диапазона Excel не ломает, а укорачивает — в макете тоже.
                  const single = new RegExp(`Продажи!B${first + 1}(?![:\\d])`);
                  if (typeof cell === "string" && single.test(cell)) {
                    row[index] = cell.replace(/Продажи!B\d+/, "#REF!");
                  }
                });
              }
            }
          }
        });
      }
    };
  }

  const sheets = [makeSheet("Продажи", "sheet-1"), makeSheet("Отчёт", "sheet-2")];
  (globalThis as any).Office = {
    context: {
      document: { url: "C:/книга.xlsx" },
      requirements: { isSetSupported: () => true }
    }
  };
  (globalThis as any).Excel = {
    InsertShiftDirection: { down: "Down" },
    DeleteShiftDirection: { up: "Up" },
    run: async (fn: any) => fn({
      workbook: {
        application: { calculationMode: "automatic", load: () => undefined },
        worksheets: {
          items: sheets,
          load: () => undefined,
          getActiveWorksheet: () => sheets[0],
          getItem: (key: string) => sheets.find((sheet) => sheet.id === key || sheet.name === key) ?? sheets[0]
        }
      },
      sync: async () => undefined
    })
  };
  return grids;
}

test("row operations go through the plan registry like every other change", () => {
  assert.ok(PLANNED_TOOLS.includes("insert_rows"));
  assert.ok(PLANNED_TOOLS.includes("delete_rows"));
});

test("the deletion preview names what is lost and which formulas break", async () => {
  workbook();
  const plan = await prepareDeleteRowsPlan({ sheet: "Продажи", startRow: 2, count: 1 });

  assert.equal(plan.kind, "delete_rows");
  assert.equal(plan.rowsAddress, "2:2");
  assert.equal(plan.filledCells, 3, "строка А, 10, формула");
  // Предпросмотр показывает то, что видно на листе, а не формулы.
  assert.deepEqual(plan.preview[0], ["А", 10, 1]);

  const broken = plan.formulaRisks.filter((risk) => risk.kind === "broken");
  const shrunk = plan.formulaRisks.filter((risk) => risk.kind === "shrunk");
  assert.equal(broken.length, 1, "=Продажи!B2 на листе Отчёт станет #ССЫЛКА!");
  assert.equal(broken[0].sheet, "Отчёт");
  assert.equal(broken[0].address, "B2");
  assert.equal(shrunk.length, 1, "=SUM(Продажи!B2:B5) молча пересчитается по укороченному диапазону");

  // Отката нет, и в предпросмотре это сказано прямо, вместе с отсутствием копии.
  assert.equal(plan.undoAvailable, false);
  assert.equal(plan.backup, null);
  assert.match(plan.undoNote, /не создавалось/);
});

test("the insertion preview finds the sums that will not cover the new rows", async () => {
  workbook();
  // Вставка строки 6 — сразу под диапазоном SUM(Продажи!B2:B5).
  const plan = await prepareInsertRowsPlan({ sheet: "Продажи", startRow: 6, count: 1 });
  const missed = plan.formulaRisks.filter((risk) => risk.kind === "missed");
  assert.equal(missed.length, 1);
  assert.equal(missed[0].sheet, "Отчёт");
  assert.match(plan.undoNote, /Отмены нет/);
});

test("deletion reports the reference errors it caused instead of only the target", async () => {
  const grids = workbook();
  const plan = await prepareDeleteRowsPlan({ sheet: "Продажи", startRow: 2, count: 1 });
  const result = await executeRowOpPlan(plan) as any;

  assert.equal(result.executionState, "verified");
  assert.equal(result.deleted, 1);
  assert.equal(result.lostFilledCells, 3);
  assert.equal(grids["Продажи"].length, 4, "строка действительно удалена");
  // Главное: сломанная ссылка на другом листе названа, а не пропущена.
  assert.equal(result.newRefErrors, 1);
  assert.deepEqual(result.brokenCells, ["Отчёт!B2"]);
  assert.match(result.refNote, /ошибок ссылок/);
  assert.equal(result.undoable, false);
});

test("a manual edit between preview and confirmation cancels the deletion", async () => {
  const grids = workbook();
  const plan = await prepareDeleteRowsPlan({ sheet: "Продажи", startRow: 2, count: 1 });
  grids["Продажи"][1][0] = "изменено вручную";

  await assert.rejects(() => executeRowOpPlan(plan), (error: any) => {
    assert.match(error.message, /изменились после предпросмотра/);
    assert.equal(error.executionState, "failed_before_write");
    return true;
  });
  assert.equal(grids["Продажи"].length, 5, "ничего не удалено");
});

test("a refusal from Excel leaves the outcome unknown, not successful", async () => {
  workbook({ deleteThrows: true });
  const plan = await prepareDeleteRowsPlan({ sheet: "Продажи", startRow: 2, count: 1 });
  await assert.rejects(() => executeRowOpPlan(plan), (error: any) => {
    assert.equal(error.executionState, "unknown");
    assert.match(error.message, /перечитайте лист/);
    return true;
  });
});

test("a protected sheet is refused before anything is attempted", async () => {
  workbook({ protectedSheet: true });
  await assert.rejects(
    () => prepareDeleteRowsPlan({ sheet: "Продажи", startRow: 2, count: 1 }),
    /защищён/
  );
});
