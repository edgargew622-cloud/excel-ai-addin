import test from "node:test";
import assert from "node:assert/strict";
import { literalMentionsSheet, referencesSheet, sheetImpact, usesName } from "./sheetImpact";

const letters = (index: number) => String.fromCharCode(64 + index);

test("a plain reference, a quoted one and a reference inside text are told apart", () => {
  assert.equal(referencesSheet("=Продажи!A1*2", "Продажи"), true);
  assert.equal(referencesSheet("='Продажи 2026'!A1", "Продажи 2026"), true);
  assert.equal(referencesSheet("=SUM(продажи!A1:A3)", "Продажи"), true, "регистр в имени листа не важен");
  assert.equal(referencesSheet("=A1*2", "Продажи"), false);
  // Замер: INDIRECT с именем в тексте Excel при переименовании не переписал.
  assert.equal(referencesSheet('=INDIRECT("Продажи!A1")', "Продажи"), false);
  assert.equal(literalMentionsSheet('=INDIRECT("Продажи!A1")', "Продажи"), true);
  assert.equal(literalMentionsSheet('=HYPERLINK("#\'Продажи\'!A1","к продажам")', "Продажи"), true);
  assert.equal(literalMentionsSheet("=Продажи!A1", "Продажи"), false);
});

test("a named range is found as a whole word, not inside another name", () => {
  assert.equal(usesName("=Ставка*2", "Ставка"), true);
  assert.equal(usesName("=СтавкаНДС*2", "Ставка"), false);
  assert.equal(usesName('="Ставка"', "Ставка"), false, "в тексте — не имя");
});

test("the impact lists what breaks and what only mentions the sheet", () => {
  const impact = sheetImpact([
    { name: "Продажи", rowIndex: 0, columnIndex: 0, formulas: [[10, "=Продажи!A1"]], values: [[10, 10]] },
    {
      name: "Отчёт", rowIndex: 0, columnIndex: 0,
      formulas: [["=Продажи!A1", '=INDIRECT("Продажи!A1")', "=Ставка*2", "см. лист Продажи", "=B9"]],
      values: [[10, 10, 2, "см. лист Продажи", 0]]
    }
  ], "Продажи", ["Ставка"], letters);
  assert.deepEqual(impact.referencing.map((item) => `${item.sheet}!${item.cell}`), ["Отчёт!A1"], "ссылка листа на самого себя не считается");
  assert.deepEqual(impact.literal.map((item) => item.cell), ["B1"]);
  assert.deepEqual(impact.viaNames.map((item) => item.cell), ["C1"]);
  assert.deepEqual(impact.textMentions.map((item) => item.cell), ["D1"]);
});
