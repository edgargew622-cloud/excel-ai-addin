/**
 * Выписка в формате обмена 1С с банком (1CClientBankExchange, этап 11.3).
 *
 * Текстовый файл «Ключ=Значение» с секциями документов. Разбирается в одну
 * таблицу операций: дата, номер, вид (списание/поступление), суммы в
 * отдельных столбцах, контрагент, его ИНН и счёт, назначение платежа.
 * Наша сторона (свой счёт) определяется по РасчСчет из заголовка: если
 * плательщик — мы, это списание, контрагент — получатель; иначе наоборот.
 */

import type { FileTable } from "./types.js";

export const BANK_1C_HEADER = ["Дата", "Номер", "Вид операции", "Списание", "Поступление", "Контрагент", "ИНН контрагента", "Счёт контрагента", "Назначение платежа", "Вид документа"];

export function isBank1C(text: string): boolean {
  return /^﻿?\s*1CClientBankExchange\s*$/m.test(text.slice(0, 200).split(/\r?\n/)[0] ?? "") || text.trimStart().startsWith("1CClientBankExchange");
}

/** «01.09.2026» → номер дня Excel. */
function excelDay(text: string): number | null {
  const m = /^(\d{2})\.(\d{2})\.(\d{4})$/.exec(text.trim());
  if (!m) return null;
  const utc = Date.UTC(+m[3], +m[2] - 1, +m[1]);
  return Math.round(utc / 86_400_000) + 25569;
}

function money(text: string | undefined): number | null {
  if (!text) return null;
  const value = Number(text.trim().replace(/\s/g, "").replace(",", "."));
  return Number.isFinite(value) ? value : null;
}

export interface Bank1C { table: FileTable; summary: string; warnings: string[] }

export function parseBank1C(text: string, fileName: string): Bank1C {
  const lines = text.replace(/\r\n?/g, "\n").split("\n").map((line) => line.trim());
  const header: Record<string, string> = {};
  const accounts = new Set<string>();
  const documents: { kind: string; fields: Record<string, string> }[] = [];
  let current: { kind: string; fields: Record<string, string> } | null = null;
  let inAccount = false;
  const warnings: string[] = [];
  for (const line of lines) {
    if (!line) continue;
    if (line.startsWith("СекцияДокумент")) { current = { kind: line.split("=")[1]?.trim() ?? "", fields: {} }; continue; }
    if (line === "КонецДокумента") { if (current) documents.push(current); current = null; continue; }
    if (line === "СекцияРасчСчет") { inAccount = true; continue; }
    if (line === "КонецРасчСчет") { inAccount = false; continue; }
    if (line === "КонецФайла") break;
    const eq = line.indexOf("=");
    if (eq < 0) continue;
    const key = line.slice(0, eq).trim();
    const value = line.slice(eq + 1).trim();
    if (current) { current.fields[key] = value; continue; }
    if (key === "РасчСчет" && value) accounts.add(value);
    if (!inAccount) header[key] = value;
    else header[`Счёт.${key}`] = value;
  }
  if (current) warnings.push("Последний документ не закрыт строкой «КонецДокумента» — он взят как есть.");
  if (current) documents.push(current);

  const rows: (string | number | null)[][] = [BANK_1C_HEADER];
  const dateFormats: Record<string, string> = {};
  let outSum = 0;
  let inSum = 0;
  let undetermined = 0;
  for (const doc of documents) {
    const f = doc.fields;
    const amount = money(f["Сумма"]);
    let outgoing: boolean | null = null;
    if (accounts.size) {
      if (f["ПлательщикСчет"] && accounts.has(f["ПлательщикСчет"])) outgoing = true;
      else if (f["ПолучательСчет"] && accounts.has(f["ПолучательСчет"])) outgoing = false;
    }
    if (outgoing === null) {
      if (f["ДатаСписано"]) outgoing = true;
      else if (f["ДатаПоступило"]) outgoing = false;
    }
    if (outgoing === null) { undetermined++; outgoing = true; }
    const party = outgoing ? "Получатель" : "Плательщик";
    // Название: «Получатель1» часто полнее «Получатель» (там бывает «ИНН … КПП …»).
    const name = (f[`${party}1`] || f[party] || "").replace(/^ИНН\s*\d+\s*/i, "").trim();
    const dateText = (outgoing ? f["ДатаСписано"] : f["ДатаПоступило"]) || f["Дата"] || "";
    const day = excelDay(dateText);
    const rowIndex = rows.length;
    rows.push([
      day ?? dateText,
      f["Номер"] ?? "",
      outgoing ? "Списание" : "Поступление",
      outgoing ? amount : null,
      outgoing ? null : amount,
      name,
      f[`${party}ИНН`] ?? "",
      f[`${party}Счет`] ?? "",
      f["НазначениеПлатежа"] ?? "",
      doc.kind
    ]);
    if (day !== null) dateFormats[`${rowIndex},0`] = "dd.mm.yyyy";
    if (amount !== null) { if (outgoing) outSum += amount; else inSum += amount; }
  }
  if (undetermined) warnings.push(`У ${undetermined} документов не удалось понять направление по счёту — они считаются списанием; проверьте столбец «Вид операции».`);
  const round = (n: number) => Math.round(n * 100) / 100;
  // Сверка с итогами секции счёта, если они есть.
  const totalOut = money(header["Счёт.ВсегоСписано"]);
  const totalIn = money(header["Счёт.ВсегоПоступило"]);
  if (totalOut !== null && Math.abs(totalOut - outSum) > 0.005) warnings.push(`Сумма списаний по документам ${round(outSum)} не равна итогу выписки ${totalOut}.`);
  if (totalIn !== null && Math.abs(totalIn - inSum) > 0.005) warnings.push(`Сумма поступлений по документам ${round(inSum)} не равна итогу выписки ${totalIn}.`);
  const period = header["ДатаНачала"] && header["ДатаКонца"] ? ` за ${header["ДатаНачала"]}–${header["ДатаКонца"]}` : "";
  const summary = `Выписка 1С${period}: документов ${documents.length}, списано ${round(outSum)}, поступило ${round(inSum)}` +
    `${accounts.size ? `, счёт ${[...accounts].join(", ")}` : ""}${header["Счёт.НачальныйОстаток"] ? `, остаток на начало ${header["Счёт.НачальныйОстаток"]}, на конец ${header["Счёт.КонечныйОстаток"] ?? "—"}` : ""}.`;
  const width = BANK_1C_HEADER.length;
  return {
    table: { name: `Выписка 1С — ${fileName}`, rows: rows.length, columns: width, cells: rows, dateFormats },
    summary,
    warnings
  };
}
