/**
 * Формулы, которые панель не записывает (07.10.2026, разбор проекта: формула
 * при isFormula=true не проверялась на опасные функции).
 *
 * Модель может принести такую формулу из ячейки, файла или страницы, которые
 * ей подсунули, — карточку подтверждения человек читает не всегда. Обычный
 * текст панель и так защищает апострофом; здесь — сами формулы:
 * - DDE («=CMD|'/c calc'!A1») — Excel предлагает запустить внешнюю программу;
 * - WEBSERVICE — формула сама отправляет данные книги на любой адрес;
 * - CALL, REGISTER, REGISTER.ID — вызов функций из DLL;
 * - RTD — подключение к внешнему серверу данных (COM);
 * - EXEC, SQL.REQUEST — запуск программы и запрос к внешней базе.
 * HYPERLINK не запрещается: ссылки просят часто, а переход по ним — щелчок человека.
 */

const BLOCKED_FUNCTIONS: Record<string, string> = {
  WEBSERVICE: "отправляет данные книги на внешний адрес",
  CALL: "вызывает функцию из DLL",
  REGISTER: "подключает функцию из DLL",
  "REGISTER.ID": "подключает функцию из DLL",
  RTD: "подключается к внешнему серверу данных",
  EXEC: "запускает программу",
  "SQL.REQUEST": "обращается к внешней базе данных"
};

/** Формула без текстов в кавычках: «|» или имя функции внутри строки — не код. */
function outsideStrings(formula: string): string {
  return formula.replace(/"(?:[^"]|"")*"/g, '""');
}

/** Почему формулу записывать нельзя; null — можно. */
export function unsafeFormulaReason(value: unknown): string | null {
  if (typeof value !== "string" || !value.trimStart().startsWith("=")) return null;
  const code = outsideStrings(value);
  // «|» вне строк в формуле Excel встречается только в DDE: программа|тема!элемент.
  if (code.includes("|")) return "DDE-ссылка: Excel предложит запустить внешнюю программу";
  const calls = code.toUpperCase().matchAll(/(?<![A-Z0-9_.])([A-Z_][A-Z0-9_.]*)\s*\(/g);
  for (const match of calls) {
    const name = match[1].replace(/^_XLFN\./, "");
    if (BLOCKED_FUNCTIONS[name]) return `${name} ${BLOCKED_FUNCTIONS[name]}`;
  }
  return null;
}

/** Первая опасная формула списка — с причиной. */
export function findUnsafeFormula(values: readonly unknown[]): { formula: string; reason: string } | null {
  for (const value of values) {
    const reason = unsafeFormulaReason(value);
    if (reason) return { formula: String(value), reason };
  }
  return null;
}

export function unsafeFormulaMessage(found: { formula: string; reason: string }): string {
  const shown = found.formula.length > 80 ? `${found.formula.slice(0, 77)}…` : found.formula;
  return `Формулу «${shown}» панель не записывает: ${found.reason}. Такие формулы часто приходят из подброшенного текста ` +
    "в ячейке, файле или на странице. Операция не выполнялась. Если пользователь сам просил именно это — скажи ему, " +
    "что панель такие формулы не ставит, и предложи ввести её вручную.";
}
