/**
 * Приём резервной копии книги по срезам.
 *
 * Панель выгружает книгу кусками и присылает их по одному. Сервер держит
 * незавершённую загрузку во временном виде, проверяет порядок, объём и полноту
 * и только потом публикует готовый файл. Здесь — вся проверяемая логика без
 * работы с диском: имена, пределы, состояния. Файловые операции живут в server.ts.
 *
 * Правила этапа 5, закреплённые здесь:
 * — путь и имя задаёт сервер, а не клиент: имя из книги только очищается;
 * — исходное расширение сохраняется, ничто не переименовывается в .xlsx;
 * — объём и срок хранения ограничены, незавершённые загрузки истекают.
 */

export const MAX_BACKUP_BYTES = 100 * 1024 * 1024;
export const MAX_SLICE_BYTES = 4 * 1024 * 1024;
export const MAX_SLICES = 2048;
export const MAX_ACTIVE_UPLOADS = 4;
/** Незавершённая загрузка живёт не дольше: панель могла закрыться посреди. */
export const UPLOAD_TTL_MS = 10 * 60_000;
/** Готовые копии старше этого удаляются. */
export const BACKUP_RETENTION_MS = 7 * 24 * 60 * 60_000;
export const MAX_BACKUPS_TOTAL_BYTES = 1024 * 1024 * 1024;

/** Форматы книг Excel. Переименовывать чужой формат в .xlsx нельзя: файл
 * перестанет открываться, а пользователь будет считать копию целой. */
const ALLOWED_EXTENSIONS = [".xlsx", ".xlsm", ".xlsb", ".xls", ".csv"];

export interface UploadRequest {
  fileName: string;
  sizeBytes: number;
  sliceCount: number;
}

export interface Upload extends UploadRequest {
  id: string;
  extension: string;
  safeName: string;
  receivedBytes: number;
  nextIndex: number;
  startedAt: number;
}

export class BackupError extends Error {}

/** Оставляет от имени книги только безопасную часть. Клиентскому имени
 * доверять нельзя: там могут быть разделители пути и переходы вверх. */
export function sanitizeWorkbookName(raw: string): string {
  const base = String(raw ?? "").split(/[\\/]/).pop() ?? "";
  const withoutExtension = base.replace(/\.[^.]*$/, "");
  const cleaned = withoutExtension
    .replace(/[^\p{L}\p{N} ._-]/gu, "")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 80);
  return cleaned || "книга";
}

export function extensionOf(raw: string): string {
  const base = String(raw ?? "").split(/[\\/]/).pop() ?? "";
  const match = /\.[^.]+$/.exec(base);
  const found = match ? match[0].toLowerCase() : "";
  return ALLOWED_EXTENSIONS.includes(found) ? found : ".xlsx";
}

/** Имя готовой копии задаёт сервер: время, очищенное имя книги, исходное
 * расширение. Время впереди, чтобы копии сортировались по порядку. */
export function backupFileName(workbookName: string, at: Date = new Date()): string {
  const stamp = at.toISOString().replace(/[-:]/g, "").replace(/\.\d+Z$/, "Z");
  return `${stamp}-${sanitizeWorkbookName(workbookName)}${extensionOf(workbookName)}`;
}

export function validateUploadRequest(request: UploadRequest): void {
  if (!Number.isInteger(request.sizeBytes) || request.sizeBytes <= 0) {
    throw new BackupError("Размер книги должен быть целым положительным числом.");
  }
  if (request.sizeBytes > MAX_BACKUP_BYTES) {
    throw new BackupError(`Книга занимает ${request.sizeBytes} байт — больше предела копии ${MAX_BACKUP_BYTES}.`);
  }
  if (!Number.isInteger(request.sliceCount) || request.sliceCount <= 0 || request.sliceCount > MAX_SLICES) {
    throw new BackupError(`Число срезов должно быть целым от 1 до ${MAX_SLICES}.`);
  }
}

export function createUpload(id: string, request: UploadRequest, now = Date.now()): Upload {
  validateUploadRequest(request);
  return {
    id,
    fileName: request.fileName,
    sizeBytes: request.sizeBytes,
    sliceCount: request.sliceCount,
    extension: extensionOf(request.fileName),
    safeName: sanitizeWorkbookName(request.fileName),
    receivedBytes: 0,
    nextIndex: 0,
    startedAt: now
  };
}

export function isExpired(upload: Upload, now = Date.now()): boolean {
  return now - upload.startedAt > UPLOAD_TTL_MS;
}

/**
 * Проверяет очередной срез. Порядок строгий: пропуск или повтор означает, что
 * собранный файл не совпадёт с книгой, а молча склеенная копия хуже отсутствия
 * копии — на неё будут рассчитывать.
 */
export function acceptSlice(upload: Upload, index: number, byteLength: number): Upload {
  if (index !== upload.nextIndex) {
    throw new BackupError(`Ожидался срез ${upload.nextIndex}, пришёл ${index}. Загрузка прервана.`);
  }
  if (!Number.isInteger(byteLength) || byteLength <= 0) {
    throw new BackupError(`Срез ${index} пуст.`);
  }
  if (byteLength > MAX_SLICE_BYTES) {
    throw new BackupError(`Срез ${index} занимает ${byteLength} байт — больше предела ${MAX_SLICE_BYTES}.`);
  }
  const received = upload.receivedBytes + byteLength;
  if (received > upload.sizeBytes) {
    throw new BackupError(`Срезы дали ${received} байт при заявленных ${upload.sizeBytes}. Загрузка прервана.`);
  }
  return { ...upload, receivedBytes: received, nextIndex: index + 1 };
}

/** Публиковать можно только полную копию: недостача или недосчитанные срезы
 * означают, что файл не откроется. */
export function assertComplete(upload: Upload): void {
  if (upload.nextIndex !== upload.sliceCount) {
    throw new BackupError(`Получено ${upload.nextIndex} срезов из ${upload.sliceCount}. Копия не опубликована.`);
  }
  if (upload.receivedBytes !== upload.sizeBytes) {
    throw new BackupError(`Получено ${upload.receivedBytes} байт из ${upload.sizeBytes}. Копия не опубликована.`);
  }
}

export interface StoredBackup {
  name: string;
  sizeBytes: number;
  createdAt: number;
}

/** Что удалить, чтобы уложиться в срок и объём хранения. Старые уходят первыми. */
export function backupsToRemove(backups: StoredBackup[], now = Date.now()): string[] {
  const doomed = new Set<string>();
  const alive: StoredBackup[] = [];
  for (const backup of backups) {
    if (now - backup.createdAt > BACKUP_RETENTION_MS) doomed.add(backup.name);
    else alive.push(backup);
  }
  alive.sort((a, b) => a.createdAt - b.createdAt);
  let total = alive.reduce((sum, backup) => sum + backup.sizeBytes, 0);
  for (const backup of alive) {
    if (total <= MAX_BACKUPS_TOTAL_BYTES) break;
    doomed.add(backup.name);
    total -= backup.sizeBytes;
  }
  return [...doomed];
}
