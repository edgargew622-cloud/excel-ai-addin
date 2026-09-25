/**
 * Создание резервной копии книги.
 *
 * Панель выгружает книгу срезами через Office и отправляет их на локальный
 * сервер, который проверяет порядок и полноту и публикует готовый файл.
 * Книга при этом не меняется: выгрузка — операция чтения.
 *
 * Важная оговорка этапа 5: копия отражает то, что отдал Excel, а не то, что
 * лежит на диске. Насколько они расходятся при несохранённых правках,
 * показывает отдельный замер `measure_workbook_export`; здесь мы честно
 * сообщаем, что копия снята из открытой книги.
 */

import { MAX_SLICE_BYTES, openWorkbookFile, readSlice, closeWorkbookFile, type OfficeFile } from "./workbookExport";
import { apiHeaders } from "../taskpane/api/panelToken";

export interface BackupResult {
  ok: boolean;
  name?: string;
  directory?: string;
  sizeBytes?: number;
  sliceCount?: number;
  elapsedMs?: number;
  removedOldCopies?: string[];
  error?: string;
  note: string;
}

/**
 * Последняя удачная копия за этот сеанс панели.
 *
 * Нужна предпросмотру вставки и удаления строк: у этих операций нет отката,
 * и единственный честный ответ на вопрос «что делать, если пойдёт не так» —
 * назвать копию, к которой можно вернуться, или признать, что её нет.
 */
let lastBackup: { name: string; directory: string; at: string } | null = null;

export function lastWorkbookBackup(): Readonly<{ name: string; directory: string; at: string }> | null {
  return lastBackup;
}

async function post(path: string, body: unknown): Promise<any> {
  const response = await fetch(path, {
    method: "POST",
    headers: apiHeaders(),
    body: JSON.stringify(body)
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(String(payload?.error?.message ?? `Сервер вернул ${response.status}.`));
  }
  return payload;
}

/** Срез приходит от Office как массив байтов; сервер ждёт base64. */
export function toBase64(data: unknown): string {
  const bytes = data instanceof Uint8Array ? data : Uint8Array.from((data as number[]) ?? []);
  let binary = "";
  for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
  return btoa(binary);
}

export async function createWorkbookBackup(
  options: { fileName?: string; sliceSizeBytes?: number } = {}
): Promise<BackupResult> {
  const startedAt = Date.now();
  const sliceSize = Math.min(Math.max(options.sliceSizeBytes ?? MAX_SLICE_BYTES, 1024), MAX_SLICE_BYTES);
  const fileName = options.fileName?.trim() || documentName();

  let file: OfficeFile;
  try {
    file = await openWorkbookFile(sliceSize);
  } catch (error: any) {
    return { ok: false, error: error?.message ?? String(error), note: "Копия не создавалась: Excel не выдал файл книги." };
  }

  let uploadId = "";
  try {
    const begin = await post("/api/backup/begin", {
      fileName,
      sizeBytes: file.size,
      sliceCount: file.sliceCount
    });
    uploadId = String(begin.uploadId);

    for (let index = 0; index < file.sliceCount; index++) {
      const slice = await readSlice(file, index);
      await post("/api/backup/slice", { uploadId, index, data: toBase64(slice.data) });
    }

    const finished = await post("/api/backup/finish", { uploadId });
    lastBackup = {
      name: String(finished.name),
      directory: String(finished.directory),
      at: new Date().toISOString()
    };
    return {
      ok: true,
      name: String(finished.name),
      directory: String(finished.directory),
      sizeBytes: Number(finished.sizeBytes),
      sliceCount: file.sliceCount,
      elapsedMs: Date.now() - startedAt,
      ...(Array.isArray(finished.removed) && finished.removed.length ? { removedOldCopies: finished.removed } : {}),
      note: "Копия снята из открытой книги вместе с несохранёнными правками, поэтому она может отличаться от файла на диске. " +
        "Восстановление ручное: откройте копию в Excel как обычный файл. Перенос отдельного листа из копии не восстанавливает межлистовые ссылки."
    };
  } catch (error: any) {
    // Незавершённую загрузку закрываем сами: иначе она займёт место до истечения
    // срока, а недособранный файл не должен пережить сбой.
    if (uploadId) {
      try { await post("/api/backup/abort", { uploadId }); } catch { /* сервер уже закрыл */ }
    }
    return {
      ok: false,
      error: error?.message ?? String(error),
      note: "Копия не опубликована. Готового файла нет — на него нельзя рассчитывать."
    };
  } finally {
    await closeWorkbookFile(file);
  }
}

function documentName(): string {
  try {
    const url = String((globalThis as any).Office?.context?.document?.url ?? "");
    return url || "книга.xlsx";
  } catch {
    return "книга.xlsx";
  }
}
