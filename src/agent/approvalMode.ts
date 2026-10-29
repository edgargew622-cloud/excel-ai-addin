/**
 * Режимы подтверждения (30.10.2026, просьба пользователя: «все выполняется
 * в ручном режиме» — три режима: ручной, полуручной, автомат).
 *
 * - manual — как было: каждое изменение ждёт «Выполнить», чтение чужого
 *   листа — «Разрешить».
 * - semi — агент сам делает то, что отменяется кнопкой «Отменить».
 *   Подтверждение — для удалений, вставки строк и столбцов, таблицы Excel,
 *   всего, что план объявил неотменяемым, и для записи поверх больше чем
 *   OVERWRITE_LIMIT заполненных ячеек.
 * - auto — подтверждаются только удаление строк, столбцов, листа и
 *   дубликатов на месте.
 *
 * «Только анализ» сильнее любого режима: изменения в нём запрещены вовсе.
 */
export type ApprovalMode = "manual" | "semi" | "auto";

export const APPROVAL_MODES: { id: ApprovalMode; label: string; title: string }[] = [
  { id: "manual", label: "Ручной", title: "Каждое изменение книги ждёт кнопки «Выполнить»." },
  { id: "semi", label: "Полуручной", title: "Обратимое агент делает сам; удаления, вставку строк и столбцов, большие перезаписи — после вашего подтверждения." },
  { id: "auto", label: "Автомат", title: "Агент делает всё сам; подтверждаются только удаление строк, столбцов, листов и дубликатов." }
];

export const DEFAULT_APPROVAL_MODE: ApprovalMode = "semi";

/** Перезапись больше стольких заполненных ячеек в полуручном режиме подтверждается. */
export const OVERWRITE_LIMIT = 500;

export function parseApprovalMode(value: unknown): ApprovalMode {
  return value === "manual" || value === "semi" || value === "auto" ? value : DEFAULT_APPROVAL_MODE;
}

const filled = (cells: unknown): number =>
  Array.isArray(cells)
    ? cells.reduce((sum: number, row) => sum + (Array.isArray(row) ? row.filter((value) => value !== "" && value !== null && value !== undefined).length : 0), 0)
    : 0;

/** Сколько заполненных ячеек затрёт запись (по снимку области в плане). */
export function overwrittenCells(plan: unknown): number {
  const p = plan as Record<string, any> | null;
  if (!p || typeof p !== "object") return 0;
  if (p.kind === "set_ranges_values" && Array.isArray(p.items)) return p.items.reduce((sum: number, item: unknown) => sum + overwrittenCells(item), 0);
  return filled(p.before ?? p.beforeFormulas);
}

/** Удаление дубликатов на месте — необратимо; результат на другом листе — нет. */
function inPlaceDuplicates(name: string, plan: unknown, args: unknown): boolean {
  if (name !== "remove_duplicates") return false;
  const a = (plan ?? args) as Record<string, unknown> | null;
  return !(a && typeof a === "object" && (a.destSheet || (args as Record<string, unknown> | null)?.destSheet));
}

const ALWAYS_ASK = new Set(["delete_rows", "delete_columns", "delete_sheet"]);
const SEMI_ASK = new Set(["insert_rows", "insert_columns", "create_table", "convert_table_to_range", "delete_pivot", "delete_chart", "delete_slicer"]);
// Оформление данных не трогает: даже без точной отмены (большая область) подтверждать его незачем.
const FORMAT_ONLY = new Set(["format_range", "format_chart", "set_sheet_view", "freeze_panes", "edit_chart", "arrange_charts"]);

/** Нужно ли подтверждение изменяющей операции в этом режиме. */
export function needsApproval(mode: ApprovalMode, name: string, plan: unknown, args: unknown): boolean {
  if (mode === "manual") return true;
  if (ALWAYS_ASK.has(name) || inPlaceDuplicates(name, plan, args)) return true;
  if (mode === "auto") return false;
  if (SEMI_ASK.has(name)) return true;
  const p = plan as Record<string, unknown> | null;
  if (!FORMAT_ONLY.has(name) && p && typeof p === "object" && p.undoAvailable === false) return true;
  return overwrittenCells(plan) > OVERWRITE_LIMIT;
}

/** Чтение листа, которого пользователь не называл, — спрашивается только в ручном режиме. */
export function needsReadApproval(mode: ApprovalMode): boolean {
  return mode === "manual";
}
