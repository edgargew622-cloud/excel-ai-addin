/**
 * Замер выгрузки книги — первый шаг к резервным копиям этапа 5.
 *
 * План требует сначала выяснить на установленном Excel: доступна ли выгрузка
 * вообще, сколько занимает по времени и размеру, попадают ли туда несохранённые
 * изменения. Пока это только измерение: срезы читаются, считаются и сразу
 * отбрасываются. Никуда не отправляются, на диск не пишутся, книга не меняется.
 */

/** Предел Office для одного среза. */
export const MAX_SLICE_BYTES = 4_194_304;
/** Выше этого замер прекращается: держать в памяти огромную книгу незачем. */
export const MAX_MEASURED_BYTES = 200 * 1024 * 1024;

export interface WorkbookExportMeasurement {
  available: boolean;
  reason?: string;
  sizeBytes?: number;
  sliceCount?: number;
  sliceSizeBytes?: number;
  openMs?: number;
  readMs?: number;
  totalMs?: number;
  /** Срезы прочитаны все до одного и сошлись по объёму с заявленным размером. */
  complete?: boolean;
  measuredAt: string;
}

export type OfficeFile = {
  size: number;
  sliceCount: number;
  getSliceAsync: (index: number, callback: (result: any) => void) => void;
  closeAsync: (callback?: (result: any) => void) => void;
};

export function openWorkbookFile(sliceSize: number): Promise<OfficeFile> {
  return new Promise((resolve, reject) => {
    const document: any = (globalThis as any).Office?.context?.document;
    if (!document || typeof document.getFileAsync !== "function") {
      reject(new Error("Office.context.document.getFileAsync недоступен в этой среде."));
      return;
    }
    const compressed = (globalThis as any).Office?.FileType?.Compressed ?? "compressed";
    document.getFileAsync(compressed, { sliceSize }, (result: any) => {
      if (result?.status === "succeeded" || result?.value) resolve(result.value as OfficeFile);
      else reject(new Error(String(result?.error?.message ?? "Excel не выдал файл книги.")));
    });
  });
}

export interface WorkbookSlice {
  data: unknown;
  byteLength: number;
}

/** Возвращает и сами байты, и их число: замеру нужна длина, копии — данные. */
export function readSlice(file: OfficeFile, index: number): Promise<WorkbookSlice> {
  return new Promise((resolve, reject) => {
    file.getSliceAsync(index, (result: any) => {
      if (result?.status === "failed" || !result?.value) {
        reject(new Error(String(result?.error?.message ?? `Срез ${index} не прочитан.`)));
        return;
      }
      const data = result.value.data;
      resolve({ data, byteLength: typeof data?.length === "number" ? data.length : 0 });
    });
  });
}

/** Закрывать файл обязательно: иначе Excel держит его до конца сессии. */
export function closeWorkbookFile(file: OfficeFile): Promise<void> {
  return new Promise((resolve) => {
    try { file.closeAsync(() => resolve()); }
    catch { resolve(); }
  });
}

export async function measureWorkbookExport(
  options: { sliceSizeBytes?: number } = {}
): Promise<WorkbookExportMeasurement> {
  const measuredAt = new Date().toISOString();
  const sliceSize = Math.min(Math.max(options.sliceSizeBytes ?? MAX_SLICE_BYTES, 1024), MAX_SLICE_BYTES);
  const startedAt = Date.now();

  let file: OfficeFile;
  try {
    file = await openWorkbookFile(sliceSize);
  } catch (error: any) {
    return { available: false, reason: error?.message ?? String(error), measuredAt };
  }
  const openMs = Date.now() - startedAt;

  try {
    if (file.size > MAX_MEASURED_BYTES) {
      return {
        available: true,
        reason: `Книга занимает ${file.size} байт — больше предела замера ${MAX_MEASURED_BYTES}. Срезы не читались.`,
        sizeBytes: file.size,
        sliceCount: file.sliceCount,
        sliceSizeBytes: sliceSize,
        openMs,
        complete: false,
        measuredAt
      };
    }
    const readStartedAt = Date.now();
    let read = 0;
    for (let index = 0; index < file.sliceCount; index++) {
      read += (await readSlice(file, index)).byteLength;
    }
    const readMs = Date.now() - readStartedAt;
    return {
      available: true,
      sizeBytes: file.size,
      sliceCount: file.sliceCount,
      sliceSizeBytes: sliceSize,
      openMs,
      readMs,
      totalMs: openMs + readMs,
      // Расхождение означает, что копия из таких срезов была бы неполной.
      complete: read === file.size,
      ...(read === file.size ? {} : { reason: `Срезы дали ${read} байт вместо ${file.size}.` }),
      measuredAt
    };
  } catch (error: any) {
    return {
      available: true,
      reason: error?.message ?? String(error),
      sizeBytes: file.size,
      sliceCount: file.sliceCount,
      sliceSizeBytes: sliceSize,
      openMs,
      complete: false,
      measuredAt
    };
  } finally {
    await closeWorkbookFile(file);
  }
}
