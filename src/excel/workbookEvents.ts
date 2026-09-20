import { invalidateAfterStructuralChange, setUndoMonitorReady } from "./undo";
import { bumpWorkbookRevision, setRevisionCoverage } from "./workbookRevision";

export interface StructuralInvalidationNotice {
  kind: "structure" | "worksheet";
  changeType: string;
  address?: string;
  removedUndo: number;
}

const STRUCTURAL_CHANGE_TYPES = new Set([
  "RowInserted",
  "RowDeleted",
  "ColumnInserted",
  "ColumnDeleted",
  "CellInserted",
  "CellDeleted"
]);

export function isStructuralChangeType(changeType: unknown): boolean {
  return typeof changeType === "string" && STRUCTURAL_CHANGE_TYPES.has(changeType);
}

const listeners = new Set<(notice: StructuralInvalidationNotice) => void>();
let registrationPromise: Promise<boolean> | null = null;

export function subscribeStructuralInvalidation(
  listener: (notice: StructuralInvalidationNotice) => void
): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

function notify(notice: StructuralInvalidationNotice) {
  for (const listener of listeners) listener(notice);
}

/**
 * Регистрирует один monitor на коллекции листов.
 *
 * Custom undo хранит A1-адреса. Любая вставка/удаление строк, столбцов или
 * ячеек — в том числе выполненная пользователем через обычный UI Excel —
 * способна сдвинуть эти адреса. Поэтому при структурном событии весь стек
 * инвалидируется. Это намеренно консервативно: потерять возможность undo
 * безопаснее, чем записать старый snapshot уже в другую ячейку.
 *
 * WorksheetCollection.onChanged доступен с ExcelApi 1.9.
 */
export function ensureStructuralChangeMonitor(): Promise<boolean> {
  if (registrationPromise) return registrationPromise;

  // Until registration succeeds, custom undo must stay unavailable.
  setUndoMonitorReady(false);
  registrationPromise = (async () => {
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.9")) {
      setUndoMonitorReady(false);
      setRevisionCoverage({ content: false, structure: false, format: false, protection: false, sheetNames: false });
      return false;
    }

    await Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;

      sheets.onChanged.add(async (event) => {
        if (!isStructuralChangeType(event.changeType)) {
          bumpWorkbookRevision("content");
          return;
        }
        bumpWorkbookRevision("structure");
        const removedUndo = invalidateAfterStructuralChange();
        if (removedUndo > 0) {
          notify({
            kind: "structure",
            changeType: String(event.changeType),
            address: event.address,
            removedUndo
          });
        }
      });

      // Новый лист приходит пустым и ничего не сдвигает на других листах,
      // поэтому прежние отмены остаются в силе. Проверка 21 сентября 2026
      // года: прежнее правило чистило историю на любое событие с листами,
      // и собственное создание листа стирало свою же отмену сразу после
      // записи — отменить его было нельзя.
      sheets.onAdded.add(async () => {
        bumpWorkbookRevision("structure");
      });
      // Удаление листа уносит цели отмен вместе с содержимым: здесь очистка
      // истории оправдана.
      sheets.onDeleted.add(async () => {
        bumpWorkbookRevision("structure");
        const removedUndo = invalidateAfterStructuralChange();
        if (removedUndo > 0) notify({ kind: "worksheet", changeType: "WorksheetDeleted", removedUndo });
      });

      // Formula results may change without a direct edit in the observed range.
      sheets.onCalculated.add(async () => { bumpWorkbookRevision("content"); });
      sheets.onFormatChanged.add(async () => { bumpWorkbookRevision("format"); });
      const protectionEvents = Office.context.requirements.isSetSupported("ExcelApi", "1.14");
      const nameEvents = Office.context.requirements.isSetSupported("ExcelApi", "1.17");
      if (protectionEvents) sheets.onProtectionChanged.add(async () => { bumpWorkbookRevision("format"); });
      if (nameEvents) sheets.onNameChanged.add(async () => { bumpWorkbookRevision("structure"); });

      await ctx.sync();
      setRevisionCoverage({
        content: true,
        structure: true,
        format: true,
        protection: protectionEvents,
        sheetNames: nameEvents
      });
    });

    // Registration and initial ctx.sync completed successfully. Start a fresh
    // undo era; entries created before this point were intentionally discarded.
    setUndoMonitorReady(true);
    return true;
  })().catch((error) => {
    setUndoMonitorReady(false);
    setRevisionCoverage({ content: false, structure: false, format: false, protection: false, sheetNames: false });
    registrationPromise = null;
    throw error;
  });

  return registrationPromise;
}
