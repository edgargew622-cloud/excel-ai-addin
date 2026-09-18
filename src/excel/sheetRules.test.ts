import test from "node:test";
import assert from "node:assert/strict";
import {
  checkTableName,
  checkTableStyle,
  describeConditionalRule,
  describeFreeze,
  headerProblems,
  parseConditionalRequest,
  parseFreezeLocation,
  parseFreezeRequest,
  ruleFormula,
  ruleMatches
} from "./sheetRules";

test("the freeze location Excel reports is read as rows and columns", () => {
  assert.deepEqual(parseFreezeLocation(null), { rows: 0, columns: 0 });
  assert.deepEqual(parseFreezeLocation("Сотрудники!$1:$1"), { rows: 1, columns: 0 });
  assert.deepEqual(parseFreezeLocation("A:B"), { rows: 0, columns: 2 });
  assert.deepEqual(parseFreezeLocation("Лист1!A1:B3"), { rows: 3, columns: 2 });
  assert.equal(parseFreezeLocation("непонятно"), null);
  assert.equal(describeFreeze({ rows: 1, columns: 0 }), "строки 1–1");
  assert.equal(describeFreeze({ rows: 0, columns: 0 }), "ничего не закреплено");
});

test("a freeze request is checked before anything reaches Excel", () => {
  assert.deepEqual(parseFreezeRequest({ rows: 1 }), { rows: 1, columns: 0 });
  assert.deepEqual(parseFreezeRequest({}), { rows: 0, columns: 0 }, "без аргументов — снять закрепление");
  assert.throws(() => parseFreezeRequest({ rows: -1 }), /от 0 до 100/);
  assert.throws(() => parseFreezeRequest({ columns: 1.5 }), /целым/);
});

test("a comparison rule needs a value and a way to highlight", () => {
  const rule = parseConditionalRequest({ rule: "greaterThan", value: 150000, fillColor: "ffc7ce" });
  assert.equal(rule.highlight?.fillColor, "#FFC7CE");
  assert.match(describeConditionalRule(rule), /больше 150000/);

  assert.throws(() => parseConditionalRequest({ rule: "greaterThan", fillColor: "#FF0000" }), /нужно значение/);
  assert.throws(() => parseConditionalRequest({ rule: "greaterThan", value: 5 }), /как подсветить/);
  assert.throws(() => parseConditionalRequest({ rule: "between", value: 10, fillColor: "#FF0000" }), /два числа/);
  assert.throws(() => parseConditionalRequest({ rule: "between", value: 10, value2: 5, fillColor: "#FF0000" }), /не больше/);
  // «Больше, чем текст» Excel сравнивает по алфавиту — почти всегда не то, что имели в виду.
  assert.throws(() => parseConditionalRequest({ rule: "greaterThan", value: "Москва", fillColor: "#FF0000" }), /сравнивает числа/);
  assert.doesNotThrow(() => parseConditionalRequest({ rule: "equalTo", value: "Финансы", bold: true }));
});

test("colour scales and data bars need no highlight, but scales need both ends", () => {
  assert.throws(() => parseConditionalRequest({ rule: "colorScale", minColor: "#F8696B" }), /minColor и maxColor/);
  const scale = parseConditionalRequest({ rule: "colorScale", minColor: "#f8696b", midColor: "#FFEB84", maxColor: "#63BE7B" });
  assert.equal(scale.scale?.minColor, "#F8696B");
  assert.equal(parseConditionalRequest({ rule: "dataBar" }).barColor, "#638EC6", "цвет по умолчанию как в Excel");
});

test("text in a rule formula is quoted the way Excel expects", () => {
  assert.equal(ruleFormula(150000), "150000");
  assert.equal(ruleFormula("Финансы"), '="Финансы"');
  assert.equal(ruleFormula('Сказал "да"'), '="Сказал ""да"""');
});

test("the panel predicts matches by Excel's own rules", () => {
  const more = parseConditionalRequest({ rule: "greaterThan", value: 100000, fillColor: "#FF0000" });
  assert.equal(ruleMatches(more, 185000), true);
  assert.equal(ruleMatches(more, 95000), false);
  assert.equal(ruleMatches(more, "Оклад"), false, "текст с числом не сравнивается");

  const less = parseConditionalRequest({ rule: "lessThan", value: 10, fillColor: "#FF0000" });
  assert.equal(ruleMatches(less, ""), true, "пустая ячейка для сравнения с числом равна нулю");

  const contains = parseConditionalRequest({ rule: "textContains", text: "финанс", fillColor: "#FF0000" });
  assert.equal(ruleMatches(contains, "Финансы"), true, "регистр не различается");
  assert.equal(ruleMatches(contains, ""), false);

  const scale = parseConditionalRequest({ rule: "colorScale", minColor: "#FFFFFF", maxColor: "#000000" });
  assert.equal(ruleMatches(scale, 5), null, "у шкалы нет «совпало»: она красит всё");
});

test("only real table styles and valid names pass", () => {
  assert.equal(checkTableStyle("TableStyleMedium2"), "TableStyleMedium2");
  assert.equal(checkTableStyle("TableStyleLight21"), "TableStyleLight21");
  assert.throws(() => checkTableStyle("TableStyleMedium29"), /не существует/);
  assert.throws(() => checkTableStyle("Синий"), /не существует/);
  assert.equal(checkTableName("Сотрудники"), "Сотрудники");
  assert.throws(() => checkTableName("Мои данные"), /без пробелов/);
  assert.throws(() => checkTableName("AB12"), /адрес/, "имя вида AB12 Excel примет за ссылку");
});

test("headers Excel would change are named before the table exists", () => {
  assert.deepEqual(headerProblems(["ФИО", "Оклад"], ["ФИО", "Оклад"]), []);
  const problems = headerProblems(["ФИО", "", "ФИО", 2026, "Итого"], ["ФИО", "", "ФИО", 2026, "=B1&\"о\""]);
  assert.equal(problems.length, 4);
  assert.ok(problems.some((text) => /пустой заголовок/.test(text)));
  assert.ok(problems.some((text) => /одинаковый заголовок «ФИО»/.test(text)));
  assert.ok(problems.some((text) => /число 2026 станет текстом/.test(text)));
  assert.ok(problems.some((text) => /формула/.test(text)));
});
