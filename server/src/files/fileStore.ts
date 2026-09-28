/**
 * Прикреплённые файлы (этап 8, 8.6): разбираются на этом компьютере и лежат
 * только в памяти сервера — на диск не пишутся, в облако целиком не уходят.
 * Модель получает содержимое частями, как при чтении листа.
 */

import { randomUUID } from "node:crypto";
import { Worker } from "node:worker_threads";
import express, { type Express } from "express";
import type { ParsedFile } from "./types.js";
import { FILE_LIMITS } from "./types.js";

export const STORE_LIMITS = {
  files: 5,
  /** Сколько файл живёт после загрузки. */
  ttlMs: 4 * 60 * 60 * 1000,
  /** Разбор одного файла. */
  parseMs: 60_000,
  parseMemoryMb: 1024,
  /** Сколько строк таблицы и частей текста за одно чтение. */
  rowsPerRead: 200,
  partsPerRead: 5,
  /** Сколько знаков за одно чтение — чтобы ответ модели не раздулся. */
  charsPerRead: 30_000
};

export class FileStoreError extends Error {}

interface StoredFile {
  id: string;
  parsed: ParsedFile;
  uploadedAt: number;
}

/** Разбор в отдельном потоке с пределом памяти и времени. */
export function parseInWorker(bytes: Buffer, name: string): Promise<ParsedFile> {
  return new Promise((resolve, reject) => {
    const worker = new Worker(new URL("./parseWorker.js", import.meta.url), {
      workerData: { bytes: new Uint8Array(bytes), name },
      resourceLimits: { maxOldGenerationSizeMb: STORE_LIMITS.parseMemoryMb }
    });
    const timer = setTimeout(() => {
      void worker.terminate();
      reject(new FileStoreError(`Разбор файла дольше ${STORE_LIMITS.parseMs / 1000} секунд — остановлен. Файл слишком сложный или повреждён.`));
    }, STORE_LIMITS.parseMs);
    worker.once("message", (message: any) => {
      clearTimeout(timer);
      void worker.terminate();
      if (message.ok) resolve(message.parsed);
      else reject(new FileStoreError(message.message));
    });
    worker.once("error", (error: any) => {
      clearTimeout(timer);
      reject(new FileStoreError(/memory|heap/i.test(String(error?.message ?? error))
        ? "Разбор файла превысил предел памяти — остановлен."
        : `Разбор файла сорвался: ${error?.message ?? error}.`));
    });
  });
}

export function fileSummary(file: StoredFile) {
  const { parsed } = file;
  return {
    id: file.id,
    name: parsed.name,
    kind: parsed.kind,
    size: parsed.size,
    uploadedAt: new Date(file.uploadedAt).toISOString(),
    tables: parsed.tables.map((table, index) => ({
      index,
      name: table.name,
      rows: table.rows,
      columns: table.columns,
      firstRow: (table.cells[0] ?? []).slice(0, 30)
    })),
    textParts: parsed.text.length,
    textChars: parsed.text.reduce((total, part) => total + part.length, 0),
    warnings: parsed.warnings
  };
}

export class FileStore {
  private files = new Map<string, StoredFile>();

  constructor(private now: () => number = Date.now, private parse: (bytes: Buffer, name: string) => Promise<ParsedFile> = parseInWorker) {}

  private sweep() {
    for (const [id, file] of this.files) if (this.now() - file.uploadedAt > STORE_LIMITS.ttlMs) this.files.delete(id);
  }

  async add(bytes: Buffer, name: string) {
    this.sweep();
    if (!bytes.length) throw new FileStoreError("Файл пустой.");
    if (bytes.length > FILE_LIMITS.bytes) throw new FileStoreError(`Файл больше ${FILE_LIMITS.bytes / 1024 / 1024} МБ.`);
    if (this.files.size >= STORE_LIMITS.files) throw new FileStoreError(`Прикреплено уже ${STORE_LIMITS.files} файлов: уберите ненужный.`);
    const cleanName = name.replace(/[\u0000-\u001f\u007f]/g, "").trim().slice(0, 200) || "файл";
    const parsed = await this.parse(bytes, cleanName);
    const id = randomUUID();
    const file = { id, parsed, uploadedAt: this.now() };
    this.files.set(id, file);
    return fileSummary(file);
  }

  list() {
    this.sweep();
    return [...this.files.values()].map(fileSummary);
  }

  get(id: string): StoredFile {
    this.sweep();
    const file = this.files.get(id);
    if (!file) throw new FileStoreError("Такого файла нет: его убрали или он устарел — прикрепите заново.");
    return file;
  }

  remove(id: string) {
    this.get(id);
    this.files.delete(id);
  }

  /** Строки таблицы для модели: не больше rowsPerRead строк и charsPerRead знаков. */
  readTable(id: string, index: number, from = 0, count: number = STORE_LIMITS.rowsPerRead) {
    const file = this.get(id);
    const table = file.parsed.tables[index];
    if (!table) throw new FileStoreError(`В файле нет таблицы №${index}. Таблиц: ${file.parsed.tables.length}.`);
    const start = Math.max(0, Math.floor(from));
    const wanted = Math.min(Math.max(1, Math.floor(count)), STORE_LIMITS.rowsPerRead);
    const rows: { row: number; cells: unknown[] }[] = [];
    let chars = 0;
    for (let r = start; r < Math.min(table.rows, start + wanted); r++) {
      const cells = table.cells[r];
      chars += JSON.stringify(cells).length;
      if (rows.length && chars > STORE_LIMITS.charsPerRead) break;
      rows.push({ row: r + 1, cells });
    }
    const next = start + rows.length;
    return { file: file.parsed.name, table: table.name, rows: table.rows, columns: table.columns, from: start + 1, to: next, data: rows, ...(next < table.rows ? { continueFrom: next } : {}) };
  }

  /** Части текста для модели. */
  readText(id: string, from = 0, count: number = STORE_LIMITS.partsPerRead) {
    const file = this.get(id);
    const parts = file.parsed.text;
    const start = Math.max(0, Math.floor(from));
    const wanted = Math.min(Math.max(1, Math.floor(count)), STORE_LIMITS.partsPerRead);
    const out: { part: number; text: string }[] = [];
    let chars = 0;
    for (let index = start; index < Math.min(parts.length, start + wanted); index++) {
      chars += parts[index].length;
      if (out.length && chars > STORE_LIMITS.charsPerRead) break;
      out.push({ part: index + 1, text: parts[index].slice(0, STORE_LIMITS.charsPerRead) });
    }
    const next = start + out.length;
    return { file: file.parsed.name, kind: file.parsed.kind, parts: parts.length, from: start + 1, to: next, data: out, ...(next < parts.length ? { continueFrom: next } : {}) };
  }

  /** Вся таблица — для переноса в книгу панелью, не для модели. */
  fullTable(id: string, index: number) {
    const file = this.get(id);
    const table = file.parsed.tables[index];
    if (!table) throw new FileStoreError(`В файле нет таблицы №${index}.`);
    return { file: file.parsed.name, kind: file.parsed.kind, ...table };
  }
}

export function registerFileRoutes(app: Express, store: FileStore): void {
  const fail = (res: any, error: any) => res.status(error instanceof FileStoreError ? 400 : 500).json({ error: { message: String(error?.message ?? error) } });
  const number = (value: unknown, fallback: number) => (value === undefined || value === "" ? fallback : Number(value));
  app.post("/api/files", express.raw({ type: "application/octet-stream", limit: FILE_LIMITS.bytes }), (req, res) => {
    let name = "файл";
    try { name = decodeURIComponent(String(req.headers["x-file-name"] ?? "файл")); } catch { /* имя как есть */ }
    if (!Buffer.isBuffer(req.body)) return fail(res, new FileStoreError("Файл не дошёл до сервера."));
    store.add(req.body, name).then((summary) => res.json(summary), (error) => fail(res, error));
  });
  app.get("/api/files", (_req, res) => res.json({ files: store.list(), limits: { ...STORE_LIMITS, ...FILE_LIMITS } }));
  app.get("/api/files/:id", (req, res) => { try { res.json(fileSummary(store.get(req.params.id))); } catch (error) { fail(res, error); } });
  app.get("/api/files/:id/tables/:index", (req, res) => {
    try { res.json(store.readTable(req.params.id, Number(req.params.index), number(req.query.from, 0), number(req.query.count, STORE_LIMITS.rowsPerRead))); } catch (error) { fail(res, error); }
  });
  app.get("/api/files/:id/tables/:index/all", (req, res) => { try { res.json(store.fullTable(req.params.id, Number(req.params.index))); } catch (error) { fail(res, error); } });
  app.get("/api/files/:id/text", (req, res) => {
    try { res.json(store.readText(req.params.id, number(req.query.from, 0), number(req.query.count, STORE_LIMITS.partsPerRead))); } catch (error) { fail(res, error); }
  });
  app.delete("/api/files/:id", (req, res) => { try { store.remove(req.params.id); res.json({ files: store.list() }); } catch (error) { fail(res, error); } });
}
