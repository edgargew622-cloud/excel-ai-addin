/**
 * Отдельный поток разбора (этап 8, 8.6): испорченный или вредный файл может
 * съесть память или уйти в долгий разбор. В потоке у него свой предел памяти,
 * а основной сервер его просто завершит по таймеру — и продолжит работать.
 */

import { parentPort, workerData } from "node:worker_threads";
import { parseFile } from "./parse.js";

const { bytes, name } = workerData as { bytes: Uint8Array; name: string };
parseFile(Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength), name).then(
  (parsed) => parentPort!.postMessage({ ok: true, parsed }),
  (error) => parentPort!.postMessage({ ok: false, message: String(error?.message ?? error), expected: error?.constructor?.name === "FileParseError" })
);
