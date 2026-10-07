import test from "node:test";
import assert from "node:assert/strict";
import { unpivotRows } from "./unpivotPlans";
import { PLANNED_TOOLS } from "./plans";

const WIDE = [
  ["Город", "Товар", "Янв", "Фев", "Мар"],
  ["Москва", "Чай", 10, 12, ""],
  ["Казань", "Кофе", 5, "", 7],
  ["", "", "", "", ""]
];

test("a wide table turns into a long one: keys stay, each month becomes a row; blanks skipped", () => {
  const result = unpivotRows(WIDE, 2, { attribute: "Месяц", value: "Продажи" }, true);
  assert.deepEqual(result.header, ["Город", "Товар", "Месяц", "Продажи"]);
  assert.deepEqual(result.rows, [
    ["Москва", "Чай", "Янв", 10], ["Москва", "Чай", "Фев", 12],
    ["Казань", "Кофе", "Янв", 5], ["Казань", "Кофе", "Мар", 7]
  ]);
  assert.equal(result.skipped, 2);
  assert.equal(result.sum, 34);
});

test("blanks can be kept; the tool goes through the plan registry", () => {
  const result = unpivotRows(WIDE, 1, { attribute: "Показатель", value: "Значение" }, false);
  assert.equal(result.rows.length, 8);
  assert.deepEqual(result.rows[0], ["Москва", "Товар", "Чай"]);
  assert.ok(PLANNED_TOOLS.includes("unpivot_range"));
});
