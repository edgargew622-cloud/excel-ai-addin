import test from "node:test";
import assert from "node:assert/strict";
import {
  DEFAULT_TOLERANCES,
  detectColumns,
  invoiceNumbers,
  nameSimilarity,
  readSide,
  reconcile,
  toDay,
  toNumber
} from "./reconcileModel";

test("numbers and dates are read the way a Russian statement writes them", () => {
  assert.equal(toNumber("1 234,50"), 1234.5);
  assert.equal(toNumber("-12 000,00 руб."), -12000);
  assert.equal(toNumber("(1 500)"), -1500);
  assert.equal(toNumber("1,234.56"), 1234.56);
  assert.equal(toNumber("1.234.567"), 1234567);
  assert.equal(toNumber("ООО"), null);
  assert.equal(toDay("05.09.2026"), toDay("2026-09-05"));
  assert.equal(toDay("31.02.2026"), null);
  assert.equal(toDay(46270), 46270);
});

test("names compare without legal forms, quotes, case and word order; invoice numbers come out of the purpose", () => {
  assert.equal(nameSimilarity("ООО «Ромашка»", "Ромашка ООО"), 1);
  assert.ok(nameSimilarity("ООО Ромашка-Торг", "ООО Ромашка Торг") > 0.9);
  assert.ok(nameSimilarity("Ромашка", "Ромашкаа") > 0.8);
  assert.ok(nameSimilarity("Ромашка", "Василёк") < 0.3);
  assert.deepEqual(invoiceNumbers("Оплата по сч. № 00123 от 01.09.2026, в т.ч. НДС 20%"), ["123"]);
  assert.deepEqual(invoiceNumbers("Оплата по счету 45/А за товар"), ["45/а"]);
});

// Реестр платежей (слева) и выписка (справа): каждый трудный случай из плана.
const REGISTRY = [
  ["Дата", "Контрагент", "ИНН", "№ счёта", "Сумма"],
  ["01.09.2026", "ООО «Ромашка»", "7701234567", "101", "15 000,00"],          // 1 точная пара
  ["02.09.2026", "ООО Василёк", "7702345678", "102", "8 400,00"],             // 2 дата +3 дн.
  ["03.09.2026", "ИП Петров А.А.", "", "103", "3 000,00"],                    // 3 опечатка в названии, без ИНН
  ["04.09.2026", "АО Сигма", "7703456789", "104", "50 000,00"],               // 4 частичная оплата: 2 платежа
  ["05.09.2026", "ООО Омега", "7704567890", "105", "12 000,00"],              // 5 комиссия вычтена (11 950)
  ["06.09.2026", "ООО Только реестр", "7705678901", "106", "7 777,00"],       // 6 нет в выписке
  ["07.09.2026", "ООО Дубль", "7706789012", "107", "2 000,00"],               // 7 два одинаковых платежа
  ["07.09.2026", "ООО Дубль", "7706789012", "108", "2 000,00"],               // 8
  ["08.09.2026", "ООО Дельта", "7707890123", "109", "9 990,00"]                // 9 сумма разошлась (9 900)
];
const STATEMENT = [
  ["Дата", "Получатель", "ИНН получателя", "Назначение платежа", "Списание", "Поступление"],
  ["01.09.2026", "Ромашка ООО", "7701234567", "Оплата по сч. № 101 от 01.09.2026", "15 000,00", ""],
  ["05.09.2026", "ООО \"Василёк\"", "7702345678", "Оплата по счету 102", "8 400,00", ""],
  ["03.09.2026", "ИП Петрова А.А.", "", "Оплата услуг", "3 000,00", ""],
  ["04.09.2026", "АО «Сигма»", "7703456789", "Частичная оплата по сч. 104", "30 000,00", ""],
  ["10.09.2026", "АО «Сигма»", "7703456789", "Доплата по сч. 104", "20 000,00", ""],
  ["05.09.2026", "ООО Омега", "7704567890", "Оплата по сч. № 105", "11 950,00", ""],
  ["07.09.2026", "ООО Дубль", "7706789012", "Оплата", "2 000,00", ""],
  ["07.09.2026", "ООО Дубль", "7706789012", "Оплата", "2 000,00", ""],
  ["08.09.2026", "ООО Дельта", "7707890123", "Оплата по сч. 109", "9 900,00", ""],
  ["09.09.2026", "ПАО Банк", "", "Комиссия за ведение счёта", "1 200,00", ""],
  ["11.09.2026", "ООО Только выписка", "7708901234", "Возврат", "", "4 500,00"]
];

function run(decisions?: Record<string, boolean>) {
  const leftColumns = detectColumns(REGISTRY[0], REGISTRY.slice(1));
  const rightColumns = detectColumns(STATEMENT[0], STATEMENT.slice(1));
  return {
    leftColumns,
    rightColumns,
    result: reconcile({
      left: readSide(REGISTRY.slice(1), leftColumns),
      right: readSide(STATEMENT.slice(1), rightColumns),
      tolerances: DEFAULT_TOLERANCES,
      ...(decisions ? { decisions } : {})
    })
  };
}
const find = (items: any[], leftRow: number) => items.find((item) => item.left.includes(leftRow - 1));

test("columns are picked by header and content: date, amount or debit/credit, names, INN, purpose, invoice number", () => {
  const { leftColumns, rightColumns } = run();
  assert.deepEqual(leftColumns, { names: [1], inns: [2], texts: [], docs: [3], date: 0, amount: 4 });
  assert.deepEqual(rightColumns, { names: [1], inns: [2], texts: [3], docs: [], date: 0, debit: 4, credit: 5 });
});

test("every hard case lands in its section, and every row is used exactly once", () => {
  const { result } = run();
  assert.ok(result.coverage.ok, JSON.stringify(result.coverage));
  assert.equal(find(result.items, 1).section, "matched");
  const late = find(result.items, 2);
  assert.equal(late.section, "matched");
  assert.match(late.reason, /дата \+3 дн\./);
  const typo = find(result.items, 3);
  assert.equal(typo.section, "probable", "похожее название без ИНН — не доказано");
  assert.match(typo.reason, /название \d+%/);
  const partial = find(result.items, 4);
  assert.equal(partial.section, "probable");
  assert.deepEqual(partial.right, [3, 4]);
  assert.match(partial.reason, /частичная или сборная/);
  const fee = find(result.items, 5);
  assert.equal(fee.section, "amountDiff");
  assert.equal(fee.diff, 50);
  assert.match(fee.reason, /похоже на комиссию/);
  assert.equal(find(result.items, 6).section, "leftOnly");
  const twins = [find(result.items, 7), find(result.items, 8)];
  assert.ok(twins.every((item) => item.section === "probable" && /другой кандидат/.test(item.reason)), "два одинаковых платежа — спорно, не «совпало»");
  const delta = find(result.items, 9);
  assert.equal(delta.section, "amountDiff");
  assert.equal(delta.diff, 90);
  const bankFee = result.items.find((item) => item.right.includes(9));
  assert.equal(bankFee?.section, "fees");
  const onlyStatement = result.items.find((item) => item.right.includes(10));
  assert.equal(onlyStatement?.section, "rightOnly");
});

test("user decisions survive a rerun: «да» confirms, «нет» breaks the pair", () => {
  const first = run().result;
  const typo = find(first.items, 3);
  const partial = find(first.items, 4);
  const again = run({ [typo.key]: true, [partial.key]: false }).result;
  assert.equal(find(again.items, 3).section, "confirmed");
  assert.notEqual(find(again.items, 4).section, "probable");
  assert.ok(again.coverage.ok);
  // Тот же вход — тот же результат.
  assert.deepEqual(run().result.items.map((item) => [item.section, item.left, item.right]), first.items.map((item) => [item.section, item.left, item.right]));
});
