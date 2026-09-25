/**
 * Точки приёма резервной копии книги.
 *
 * Панель выгружает книгу срезами и присылает их по одному. Сервер копит их во
 * временном файле и публикует готовую копию только после проверки порядка,
 * объёма и полноты. Клиент не выбирает ни путь, ни имя: он присылает только
 * имя книги, из которого берётся очищенная часть и исходное расширение.
 *
 * Вся проверяемая логика — в backupStore.ts; здесь только работа с диском
 * и разбор запросов.
 */

import type { Express, Request, Response } from "express";
import { appendFileSync, existsSync, mkdirSync, readdirSync, renameSync, rmSync, statSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import {
  BackupError,
  MAX_ACTIVE_UPLOADS,
  acceptSlice,
  assertComplete,
  backupFileName,
  backupsToRemove,
  createUpload,
  isExpired,
  type Upload
} from "./backupStore.js";

const uploads = new Map<string, Upload>();

const newId = () =>
  globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`;

function fail(res: Response, error: unknown): void {
  const message = error instanceof BackupError ? error.message : String((error as Error)?.message ?? error);
  res.status(error instanceof BackupError ? 400 : 500).json({ error: { message } });
}

export function registerBackupRoutes(app: Express, projectRoot: string): void {
  const backupDir = join(projectRoot, "backups");
  const incomingDir = join(backupDir, "incoming");
  const partPath = (id: string) => join(incomingDir, `${id}.part`);

  /** Брошенные загрузки не должны занимать место до перезапуска сервера. */
  function sweepExpired(now = Date.now()): void {
    for (const [id, upload] of uploads) {
      if (!isExpired(upload, now)) continue;
      uploads.delete(id);
      try { rmSync(partPath(id), { force: true }); } catch { /* уже нет */ }
    }
  }

  function applyRetention(): string[] {
    if (!existsSync(backupDir)) return [];
    const stored = readdirSync(backupDir, { withFileTypes: true })
      .filter((entry) => entry.isFile())
      .map((entry) => {
        const stats = statSync(join(backupDir, entry.name));
        return { name: entry.name, sizeBytes: stats.size, createdAt: stats.mtimeMs };
      });
    const removed: string[] = [];
    for (const name of backupsToRemove(stored)) {
      try { rmSync(join(backupDir, name), { force: true }); removed.push(name); }
      catch { /* файл занят или уже удалён */ }
    }
    return removed;
  }

  app.post("/api/backup/begin", (req: Request, res: Response) => {
    try {
      sweepExpired();
      if (uploads.size >= MAX_ACTIVE_UPLOADS) {
        throw new BackupError(`Уже идёт ${uploads.size} загрузок копии. Дождитесь их завершения.`);
      }
      const { fileName, sizeBytes, sliceCount } = req.body ?? {};
      const upload = createUpload(newId(), {
        fileName: String(fileName ?? ""),
        sizeBytes: Number(sizeBytes),
        sliceCount: Number(sliceCount)
      });
      mkdirSync(incomingDir, { recursive: true });
      // Пустой файл создаётся сразу: дальше срезы только дописываются, и это
      // исключает случай, когда часть их ушла в несуществующий файл.
      writeFileSync(partPath(upload.id), Buffer.alloc(0));
      uploads.set(upload.id, upload);
      res.json({ uploadId: upload.id, sliceCount: upload.sliceCount, sizeBytes: upload.sizeBytes });
    } catch (error) {
      fail(res, error);
    }
  });

  app.post("/api/backup/slice", (req: Request, res: Response) => {
    try {
      const { uploadId, index, data } = req.body ?? {};
      const upload = uploads.get(String(uploadId ?? ""));
      if (!upload) throw new BackupError("Загрузка не найдена или уже завершена.");
      if (isExpired(upload)) {
        uploads.delete(upload.id);
        rmSync(partPath(upload.id), { force: true });
        throw new BackupError("Загрузка истекла по времени. Начните заново.");
      }
      if (typeof data !== "string") throw new BackupError("Срез должен быть строкой в base64.");
      const bytes = Buffer.from(data, "base64");
      const next = acceptSlice(upload, Number(index), bytes.byteLength);
      appendFileSync(partPath(upload.id), bytes);
      uploads.set(upload.id, next);
      res.json({ receivedBytes: next.receivedBytes, nextIndex: next.nextIndex, sliceCount: next.sliceCount });
    } catch (error) {
      // Любая ошибка среза делает копию недостоверной: загрузку закрываем,
      // чтобы недособранный файл нельзя было опубликовать позже.
      const id = String(req.body?.uploadId ?? "");
      if (uploads.has(id)) {
        uploads.delete(id);
        try { rmSync(partPath(id), { force: true }); } catch { /* уже нет */ }
      }
      fail(res, error);
    }
  });

  app.post("/api/backup/finish", (req: Request, res: Response) => {
    const id = String(req.body?.uploadId ?? "");
    try {
      const upload = uploads.get(id);
      if (!upload) throw new BackupError("Загрузка не найдена или уже завершена.");
      assertComplete(upload);

      const onDisk = statSync(partPath(upload.id)).size;
      if (onDisk !== upload.sizeBytes) {
        throw new BackupError(`Собранный файл занимает ${onDisk} байт вместо ${upload.sizeBytes}. Копия не опубликована.`);
      }

      mkdirSync(backupDir, { recursive: true });
      const name = backupFileName(upload.fileName);
      renameSync(partPath(upload.id), join(backupDir, name));
      uploads.delete(upload.id);
      const removed = applyRetention();
      console.log(`Резервная копия сохранена: ${name}, ${onDisk} байт.`);
      res.json({ name, sizeBytes: onDisk, directory: backupDir, removed });
    } catch (error) {
      if (uploads.has(id)) {
        uploads.delete(id);
        try { rmSync(partPath(id), { force: true }); } catch { /* уже нет */ }
      }
      fail(res, error);
    }
  });

  // Удаляется только файл своей активной загрузки. Прежде abort строил путь
  // из присланного id без проверки, и «..\..\logs\x» удалял logs\x.part вне
  // incoming/ (аудит 24 сентября 2026 года, SEC-04).
  app.post("/api/backup/abort", (req: Request, res: Response) => {
    const id = String(req.body?.uploadId ?? "");
    if (!uploads.has(id)) {
      return res.status(404).json({ error: { message: "Такой загрузки нет: отменять нечего." } });
    }
    uploads.delete(id);
    try { rmSync(partPath(id), { force: true }); } catch { /* уже нет */ }
    res.json({ ok: true });
  });

  app.get("/api/backup/list", (_req: Request, res: Response) => {
    try {
      if (!existsSync(backupDir)) return res.json({ directory: backupDir, backups: [] });
      const backups = readdirSync(backupDir, { withFileTypes: true })
        .filter((entry) => entry.isFile())
        .map((entry) => {
          const stats = statSync(join(backupDir, entry.name));
          return { name: entry.name, sizeBytes: stats.size, createdAt: new Date(stats.mtimeMs).toISOString() };
        })
        .sort((a, b) => b.createdAt.localeCompare(a.createdAt));
      res.json({ directory: backupDir, backups });
    } catch (error) {
      fail(res, error);
    }
  });
}
