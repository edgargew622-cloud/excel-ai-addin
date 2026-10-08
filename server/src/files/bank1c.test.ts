import test from "node:test";
import assert from "node:assert/strict";
import { BANK_1C_HEADER, isBank1C, parseBank1C } from "./bank1c.js";
import { parseTxt } from "./csv.js";

const FILE = [
  "1CClientBankExchange",
  "ВерсияФормата=1.03",
  "Кодировка=Windows",
  "Отправитель=Бухгалтерия предприятия, редакция 3.0",
  "ДатаНачала=01.09.2026",
  "ДатаКонца=30.09.2026",
  "РасчСчет=40702810900000000001",
  "СекцияРасчСчет",
  "ДатаНачала=01.09.2026",
  "ДатаКонца=30.09.2026",
  "РасчСчет=40702810900000000001",
  "НачальныйОстаток=100000.00",
  "ВсегоПоступило=4500.00",
  "ВсегоСписано=15000.00",
  "КонечныйОстаток=89500.00",
  "КонецРасчСчет",
  "СекцияДокумент=Платежное поручение",
  "Номер=101",
  "Дата=01.09.2026",
  "Сумма=15000.00",
  "ПлательщикСчет=40702810900000000001",
  "ДатаСписано=01.09.2026",
  "Плательщик=ООО \"Наша компания\"",
  "ПлательщикИНН=7700000001",
  "ПолучательСчет=40702810100000000777",
  "Получатель=ООО \"Ромашка\"",
  "ПолучательИНН=7701234567",
  "НазначениеПлатежа=Оплата по сч. № 101 от 01.09.2026, в т.ч. НДС 20% 2500.00",
  "КонецДокумента",
  "СекцияДокумент=Платежное поручение",
  "Номер=55",
  "Дата=11.09.2026",
  "Сумма=4500.00",
  "ПлательщикСчет=40702810500000000555",
  "Плательщик=ООО \"Только выписка\"",
  "ПлательщикИНН=7708901234",
  "ПолучательСчет=40702810900000000001",
  "ДатаПоступило=11.09.2026",
  "Получатель=ООО \"Наша компания\"",
  "ПолучательИНН=7700000001",
  "НазначениеПлатежа=Возврат аванса",
  "КонецДокумента",
  "КонецФайла"
].join("\r\n");

test("a 1C bank exchange file becomes a table of operations: direction by own account, counterparty is the other side", () => {
  assert.ok(isBank1C(FILE));
  const { table, summary, warnings } = parseBank1C(FILE, "kl_to_1c.txt");
  assert.deepEqual(table.cells[0], BANK_1C_HEADER);
  assert.equal(table.rows, 3);
  const [out, inc] = table.cells.slice(1);
  assert.equal(out[2], "Списание");
  assert.equal(out[3], 15000);
  assert.equal(out[4], null);
  assert.equal(out[5], "ООО \"Ромашка\"");
  assert.equal(out[6], "7701234567");
  assert.equal(inc[2], "Поступление");
  assert.equal(inc[4], 4500);
  assert.equal(inc[5], "ООО \"Только выписка\"");
  assert.equal(table.dateFormats?.["1,0"], "dd.mm.yyyy");
  assert.match(summary, /документов 2, списано 15000, поступило 4500/);
  assert.deepEqual(warnings, [], "итоги секции счёта сошлись");
});

test("a TXT upload in the 1C format comes with the table; plain text stays text", () => {
  const parsed = parseTxt(Buffer.from("\uFEFF" + FILE, "utf8"), "выписка.txt");
  assert.equal(parsed.tables.length, 1);
  assert.ok(parsed.warnings.some((line) => /Выписка 1С за 01\.09\.2026–30\.09\.2026/.test(line)));
  assert.equal(parseTxt(Buffer.from("просто заметка", "utf8"), "note.txt").tables.length, 0);
  const broken = parseBank1C(FILE.replace("ВсегоСписано=15000.00", "ВсегоСписано=16000.00"), "x.txt");
  assert.match(broken.warnings.join(" "), /списаний по документам 15000 не равна итогу выписки 16000/);
});
