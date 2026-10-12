import test from "node:test";
import assert from "node:assert/strict";
import { appendRows } from "./appendPlans";
import { PLANNED_TOOLS } from "./plans";

test("tables with the same header are stacked by column name, with the source sheet; sums are kept", () => {
  const a = { sheet: "Данные", header: ["Дата", "Регион", "Выручка"], rows: [[44197, "Север", 10], ["", "", ""], [44198, "Юг", 5]], formats: [["dd.mm.yyyy", "General", "General"], [], ["dd.mm.yyyy", "General", "General"]] };
  const b = { sheet: "Данные 2023", header: ["Регион", "Дата", "Выручка "], rows: [["Запад", 44927, 7]], formats: [["General", "dd.mm.yyyy", "General"]] };
  const out = appendRows([a, b], "Источник");
  assert.deepEqual(out.header, ["Дата", "Регион", "Выручка", "Источник"]);
  assert.deepEqual(out.rows, [[44197, "Север", 10, "Данные"], [44198, "Юг", 5, "Данные"], [44927, "Запад", 7, "Данные 2023"]]);
  assert.equal(out.formats[2][0], "dd.mm.yyyy");
  assert.deepEqual(out.counts, [2, 1]);
  assert.equal(out.sums[2], 22);
  assert.ok(PLANNED_TOOLS.includes("append_ranges"));
});

test("different headers are refused with what is missing", () => {
  const out = appendRows([
    { sheet: "А", header: ["Дата", "Сумма"], rows: [[1, 2]], formats: [[]] },
    { sheet: "Б", header: ["Дата", "Итог"], rows: [[1, 2]], formats: [[]] }
  ], null);
  assert.match(out.problems.join(" "), /«Б»: нет столбцов «Сумма»; лишние столбцы «Итог»/);
});
