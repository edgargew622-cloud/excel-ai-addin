import test from "node:test";
import assert from "node:assert/strict";
import {
  BackupError,
  BACKUP_RETENTION_MS,
  MAX_BACKUPS_TOTAL_BYTES,
  MAX_SLICE_BYTES,
  UPLOAD_TTL_MS,
  acceptSlice,
  assertComplete,
  backupFileName,
  backupsToRemove,
  createUpload,
  extensionOf,
  isExpired,
  sanitizeWorkbookName,
  validateUploadRequest
} from "./backupStore.js";

const request = { fileName: "Продажи.xlsx", sizeBytes: 900, sliceCount: 3 };

test("the client never decides where its copy lands", () => {
  // Имя приходит от клиента и может содержать путь или переход вверх.
  assert.equal(sanitizeWorkbookName("C:\\\\Users\\\\1\\\\Продажи.xlsx"), "Продажи");
  assert.equal(sanitizeWorkbookName("../../etc/passwd"), "passwd");
  assert.equal(sanitizeWorkbookName("отчёт: за март?.xlsx"), "отчёт за март");
  assert.equal(sanitizeWorkbookName(""), "книга");
  assert.equal(sanitizeWorkbookName("///"), "книга");

  const name = backupFileName("C:/книги/Продажи.xlsm", new Date("2026-09-16T08:30:00Z"));
  assert.equal(name, "20260916T083000Z-Продажи.xlsm");
  // Разделителей пути в готовом имени быть не может.
  assert.equal(/[\\/]/.test(name), false);
});

test("the original format survives: nothing is renamed to xlsx", () => {
  // Переименованный .xlsb просто не откроется, а копия будет считаться целой.
  for (const extension of [".xlsx", ".xlsm", ".xlsb", ".xls", ".csv"]) {
    assert.equal(extensionOf(`книга${extension}`), extension);
  }
  assert.equal(extensionOf("книга.XLSM"), ".xlsm", "регистр не важен");
  // Неизвестное расширение не тащим в имя файла вовсе.
  assert.equal(extensionOf("книга.exe"), ".xlsx");
  assert.equal(extensionOf("книга"), ".xlsx");
});

test("an upload is refused before it starts when the numbers make no sense", () => {
  assert.throws(() => validateUploadRequest({ ...request, sizeBytes: 0 }), BackupError);
  assert.throws(() => validateUploadRequest({ ...request, sizeBytes: 1.5 }), BackupError);
  assert.throws(() => validateUploadRequest({ ...request, sizeBytes: 999 ** 4 }), /больше предела копии/);
  assert.throws(() => validateUploadRequest({ ...request, sliceCount: 0 }), BackupError);
  assert.throws(() => validateUploadRequest({ ...request, sliceCount: 99_999 }), /от 1 до/);
});

test("slices must arrive in order: a gap would silently corrupt the copy", () => {
  let upload = createUpload("u1", request);
  upload = acceptSlice(upload, 0, 300);
  assert.equal(upload.nextIndex, 1);

  assert.throws(() => acceptSlice(upload, 2, 300), /Ожидался срез 1, пришёл 2/);
  assert.throws(() => acceptSlice(upload, 0, 300), /Ожидался срез 1, пришёл 0/, "повтор тоже отклоняется");
});

test("slices that exceed the declared size stop the upload", () => {
  let upload = createUpload("u2", request);
  upload = acceptSlice(upload, 0, 300);
  upload = acceptSlice(upload, 1, 300);
  assert.throws(() => acceptSlice(upload, 2, 500), /при заявленных 900/);
  assert.throws(() => acceptSlice(upload, 2, MAX_SLICE_BYTES + 1), /больше предела/);
  assert.throws(() => acceptSlice(upload, 2, 0), /пуст/);
});

test("an incomplete upload is never published", () => {
  let upload = createUpload("u3", request);
  upload = acceptSlice(upload, 0, 300);
  upload = acceptSlice(upload, 1, 300);
  // Срезов не хватает: файл не откроется, публиковать нечего.
  assert.throws(() => assertComplete(upload), /Получено 2 срезов из 3/);

  const shortByteCount = createUpload("u4", { ...request, sliceCount: 1 });
  assert.throws(() => assertComplete(acceptSlice(shortByteCount, 0, 100)), /Получено 100 байт из 900/);
});

test("a complete upload passes", () => {
  let upload = createUpload("u5", request);
  for (let index = 0; index < 3; index++) upload = acceptSlice(upload, index, 300);
  assert.doesNotThrow(() => assertComplete(upload));
});

test("an abandoned upload expires instead of holding space forever", () => {
  const upload = createUpload("u6", request, 1000);
  assert.equal(isExpired(upload, 1000 + UPLOAD_TTL_MS - 1), false);
  assert.equal(isExpired(upload, 1000 + UPLOAD_TTL_MS + 1), true);
});

test("retention removes the old first, by age and then by volume", () => {
  const now = Date.now();
  const day = 24 * 60 * 60_000;
  const old = { name: "старая", sizeBytes: 10, createdAt: now - BACKUP_RETENTION_MS - day };
  const fresh = { name: "свежая", sizeBytes: 10, createdAt: now };
  assert.deepEqual(backupsToRemove([old, fresh], now), ["старая"]);

  // Объём превышен: уходят самые старые, пока не уложимся.
  const big = MAX_BACKUPS_TOTAL_BYTES;
  const crowded = [
    { name: "1", sizeBytes: big, createdAt: now - 3 * day },
    { name: "2", sizeBytes: big, createdAt: now - 2 * day },
    { name: "3", sizeBytes: big, createdAt: now - day }
  ];
  assert.deepEqual(backupsToRemove(crowded, now).sort(), ["1", "2"]);

  assert.deepEqual(backupsToRemove([fresh], now), [], "укладываемся — ничего не удаляем");
});
