/**
 * Беседа в файл (этап 10, 10.1): окно «Беседы» в панели сохраняет выбранную
 * беседу текстом Markdown в «Документы\am.AI\Беседы».
 *
 * Папку «Документы» спрашиваем у Windows: у многих она перенесена в OneDrive,
 * и %USERPROFILE%\Documents тогда не та. Имя файла приходит из панели, но
 * сервер проверяет его сам: только имя, без папок и запрещённых символов, с
 * расширением .md; существующий файл не перезаписывается — рядом ложится
 * «(2)», «(3)». Текст уже есть у панели, сервер его только записывает.
 * Маршруты — за общим токеном панели, как остальные /api.
 */

import type { Express } from "express";
import { execFile } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdir, writeFile } from "node:fs/promises";
import { homedir } from "node:os";
import { basename, extname, join } from "node:path";

export const MAX_EXPORT_BYTES = 3 * 1024 * 1024;

export class ExportError extends Error {}

/** «Документы» пользователя по Windows; вне Windows или при сбое — ~/Documents. */
export function documentsFolder(run: typeof execFile = execFile): Promise<string> {
  const fallback = join(homedir(), "Documents");
  if (process.platform !== "win32") return Promise.resolve(fallback);
  return new Promise((resolve) => {
    run(
      "powershell.exe",
      ["-NoProfile", "-NonInteractive", "-Command", "[Console]::OutputEncoding=[Text.Encoding]::UTF8; [Environment]::GetFolderPath('MyDocuments')"],
      { windowsHide: true, timeout: 10_000 },
      (error, stdout) => {
        const path = String(stdout ?? "").trim();
        resolve(!error && path ? path : fallback);
      }
    );
  });
}

export function checkedFileName(name: unknown): string {
  if (typeof name !== "string") throw new ExportError("Нет имени файла.");
  const trimmed = name.trim();
  if (!trimmed || trimmed !== basename(trimmed) || /[\\/:*?"<>|\u0000-\u001f]/.test(trimmed) || /^\.+$/.test(trimmed.replace(/\.md$/i, ""))) {
    throw new ExportError("Имя файла недопустимо.");
  }
  if (extname(trimmed).toLowerCase() !== ".md") throw new ExportError("Беседа сохраняется только в файл .md.");
  if (trimmed.length > 150) throw new ExportError("Имя файла слишком длинное.");
  return trimmed;
}

/** Свободное имя: «Книга.md», затем «Книга (2).md» и так далее. */
export function freePath(folder: string, name: string, exists: (path: string) => boolean = existsSync): string {
  const stem = name.slice(0, -extname(name).length);
  for (let index = 1; index < 1000; index++) {
    const candidate = join(folder, index === 1 ? name : `${stem} (${index}).md`);
    if (!exists(candidate)) return candidate;
  }
  throw new ExportError("В папке слишком много файлов с таким именем.");
}

export async function exportConversation(folder: string, name: unknown, text: unknown): Promise<string> {
  const file = checkedFileName(name);
  if (typeof text !== "string" || !text.trim()) throw new ExportError("Беседа пуста.");
  if (Buffer.byteLength(text, "utf8") > MAX_EXPORT_BYTES) throw new ExportError("Беседа слишком большая для файла.");
  await mkdir(folder, { recursive: true });
  const path = freePath(folder, file);
  await writeFile(path, text, { encoding: "utf8", flag: "wx" });
  return path;
}

export function registerConversationExportRoutes(app: Express, folder: () => Promise<string>): void {
  app.post("/api/conversations/export", async (req, res) => {
    try {
      const path = await exportConversation(await folder(), req.body?.name, req.body?.text);
      res.json({ path });
    } catch (error: any) {
      const known = error instanceof ExportError;
      if (!known) console.error(`Беседа в файл: ${String(error?.message ?? error)}`);
      res.status(known ? 400 : 500).json({ error: { message: known ? error.message : "Не удалось записать файл. Подробности — в журнале сервера." } });
    }
  });
}
