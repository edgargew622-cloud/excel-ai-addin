/**
 * Сжатие истории беседы (08.10.2026: беседа «Книга602» после ~117 сообщений
 * упёрлась в предел запроса 1 МБ — в истории лежали целиком прочитанные
 * страницы, таблицы и сводные всех прошлых задач, — и агент просил начать
 * новый чат; запросы при этом доходили до 250 тыс. токенов).
 *
 * 1. Всегда: результаты инструментов из задач старше двух последних
 *    урезаются до начала с пометкой — нужное модель перечитает. Последние
 *    две задачи целы: «продолжай» опирается на них.
 * 2. Если запрос всё равно больше предела — ранние задачи опускаются
 *    целиком (по границе просьб, чтобы вызовы и ответы инструментов не
 *    разрывались), а модель получает об этом заметку.
 * 3. Если не помещается и текущая задача — урезаются её ранние результаты.
 *
 * Пункт 1 меняет сообщения на месте: так и сохранённая беседа не растёт
 * без конца (хранилище панели — несколько мегабайт).
 */

import type { ChatMessage } from "../taskpane/api/client";

export const KEEP_FULL_TURNS = 2;
const OLD_RESULT_CHARS = 600;
const CURRENT_RESULT_CHARS = 2_000;
const CUT_NOTE = "\n…[результат прошлой задачи сокращён панелью, чтобы беседа помещалась; нужные данные перечитай инструментом]";

const size = (value: unknown) => new TextEncoder().encode(JSON.stringify(value)).length;
const isUser = (message: ChatMessage) => message.role === "user";

/** Индексы начала задач — сообщения пользователя. */
function turnStarts(messages: readonly ChatMessage[]): number[] {
  return messages.map((message, index) => (isUser(message) ? index : -1)).filter((index) => index >= 0);
}

function shorten(message: ChatMessage, chars: number): boolean {
  const content = typeof message.content === "string" ? message.content : "";
  if (message.role !== "tool" || content.length <= chars + CUT_NOTE.length) return false;
  (message as { content: string }).content = content.slice(0, chars) + CUT_NOTE;
  return true;
}

/** Пункт 1: урезать старые результаты на месте. Возвращает, сколько сообщений сжато. */
export function compactOldResults(messages: ChatMessage[], keepTurns = KEEP_FULL_TURNS): number {
  const starts = turnStarts(messages);
  if (starts.length <= keepTurns) return 0;
  const boundary = starts[starts.length - keepTurns];
  let changed = 0;
  for (let index = 0; index < boundary; index++) {
    if (shorten(messages[index], OLD_RESULT_CHARS)) changed++;
  }
  return changed;
}

/**
 * Пункты 2–3: запрос под предел. Возвращает новый список сообщений
 * (исходный не меняется) и сколько задач опущено.
 */
export function fitRequest(messages: readonly ChatMessage[], fits: (list: ChatMessage[]) => boolean): { messages: ChatMessage[]; droppedTurns: number; shortenedCurrent: number } {
  let list = [...messages];
  if (fits(list)) return { messages: list, droppedTurns: 0, shortenedCurrent: 0 };
  const system = list.filter((message) => message.role === "system");
  let rest = list.filter((message) => message.role !== "system");
  let dropped = 0;
  // Опускаем по одной самой ранней задаче, пока не поместится; последнюю — никогда.
  while (turnStarts(rest).length > 1) {
    const starts = turnStarts(rest);
    rest = rest.slice(starts[1]);
    dropped++;
    const note: ChatMessage = { role: "system", content: `Ранние задачи этой беседы (${dropped}) опущены панелью из-за размера запроса; пользователь видит их в панели. Если нужно что-то из них — перечитай книгу.` };
    list = [...system, note, ...rest];
    if (fits(list)) return { messages: list, droppedTurns: dropped, shortenedCurrent: 0 };
  }
  // Текущая задача сама велика: урезаем её результаты с начала, последние четыре — целы.
  const copies = list.map((message) => ({ ...message }));
  const tools = copies.map((message, index) => (message.role === "tool" ? index : -1)).filter((index) => index >= 0);
  let shortened = 0;
  for (const index of tools.slice(0, Math.max(0, tools.length - 4))) {
    if (shorten(copies[index], CURRENT_RESULT_CHARS)) shortened++;
    if (fits(copies)) break;
  }
  return { messages: copies, droppedTurns: dropped, shortenedCurrent: shortened };
}

export { size as requestSize };
