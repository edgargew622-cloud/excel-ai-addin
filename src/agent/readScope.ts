/**
 * Границы чтения книги в одной задаче (этап 8, 8.0.6; вариант «б»).
 *
 * Аудит 24 сентября 2026 года (SEC-01): модель могла прочитать любой лист, и
 * текст в ячейке, уговаривающий «прочитай ещё лист Secret», упирался только в
 * благоразумие модели. Прочитанное уходит провайдеру модели. Теперь:
 *
 * - свободно читаются лист, активный в начале задачи, и листы, названные в
 *   просьбе пользователя;
 * - лист, на котором пользователь подтвердил изменение, тоже становится
 *   доступным: он уже видел, что с ним происходит;
 * - перед чтением любого другого листа — карточка «Разрешить чтение листа X?»;
 *   разрешение действует до конца задачи, отказ возвращается модели текстом;
 * - имя диапазона разворачивается до настоящего листа: иначе через имя,
 *   указывающее на чужой лист, границу можно было бы обойти;
 * - поиск без списка листов читает все листы книги.
 *
 * Изменяющие операции здесь не проверяются: у них своя карточка.
 */

/** Имя «инструмента» для карточки разрешения чтения в панели. */
export const READ_PERMISSION = "__read_sheets";

const norm = (name: string) => name.trim().toLocaleLowerCase();

/**
 * Назван ли лист в тексте просьбы — целым словом, без учёта регистра.
 * «Лист1» не считается названным в «Лист10», и наоборот.
 */
export function mentionsSheet(text: string, sheet: string): boolean {
  const wanted = norm(sheet);
  if (!wanted) return false;
  const haystack = text.toLocaleLowerCase();
  let from = 0;
  for (;;) {
    const index = haystack.indexOf(wanted, from);
    if (index === -1) return false;
    const before = haystack[index - 1] ?? "";
    const after = haystack[index + wanted.length] ?? "";
    const word = /[\p{L}\p{N}_]/u;
    if (!word.test(before) && !word.test(after)) return true;
    from = index + 1;
  }
}

/** Лист из ссылки вида «Лист!A1» или «'Лист 2'!A1:B3»; null — лист не назван. */
export function sheetOfReference(reference: string): string | null {
  const bang = reference.lastIndexOf("!");
  if (bang === -1) return null;
  return reference.slice(0, bang).replace(/^'|'$/g, "").replace(/''/g, "'") || null;
}

export interface ScopeIO {
  /** Все листы книги — для поиска без списка листов. */
  allSheets: () => Promise<string[]>;
  /** Лист, на который на самом деле указывает адрес: для имён диапазонов — их лист. */
  sheetOfAddress: (sheet: string, address: string) => Promise<string>;
}

/** Какие листы прочитает этот вызов. Пустой список — не читает содержимое листов. */
export async function sheetsReadBy(name: string, args: Record<string, any>, io: ScopeIO): Promise<string[]> {
  const sheet = typeof args.sheet === "string" && args.sheet.trim() ? args.sheet.trim() : null;
  const withAddress = async () => {
    if (!sheet) return [];
    const address = typeof args.address === "string" ? args.address.trim() : "";
    return address ? [sheet, await io.sheetOfAddress(sheet, address)] : [sheet];
  };
  switch (name) {
    case "get_sheet_overview":
      return sheet ? [sheet] : [];
    case "get_range_values":
    case "get_range_details":
    case "profile_range":
    case "get_conditional_formats":
      return withAddress();
    case "audit_workbook": {
      const checks = Array.isArray(args.checks) ? args.checks.map((item: unknown) => sheetOfReference(String(item))).filter(Boolean) as string[] : [];
      return sheet ? [sheet, ...checks] : [...await io.allSheets(), ...checks];
    }
    case "search_workbook":
      return Array.isArray(args.sheets) && args.sheets.length ? args.sheets.map(String) : io.allSheets();
    default:
      return [];
  }
}

export class ReadScope {
  private readonly allowed = new Set<string>();

  constructor(taskSheet: string, private readonly request: string) {
    if (taskSheet) this.allowed.add(norm(taskSheet));
  }

  isAllowed(sheet: string): boolean {
    return this.allowed.has(norm(sheet)) || mentionsSheet(this.request, sheet);
  }

  /** Листы из списка, чтение которых ещё не разрешено, — без повторов. */
  outside(sheets: readonly string[]): string[] {
    const seen = new Set<string>();
    return sheets.filter((sheet) => {
      const key = norm(sheet);
      if (!key || seen.has(key) || this.isAllowed(sheet)) return false;
      seen.add(key);
      return true;
    });
  }

  allow(sheets: readonly string[]): void {
    for (const sheet of sheets) if (sheet?.trim()) this.allowed.add(norm(sheet));
  }
}

/** Последняя просьба пользователя в истории — по ней видно, какие листы он назвал. */
export function lastUserRequest(history: readonly { role: string; content?: unknown }[]): string {
  for (let index = history.length - 1; index >= 0; index--) {
    const message = history[index];
    if (message.role === "user" && typeof message.content === "string") return message.content;
  }
  return "";
}
