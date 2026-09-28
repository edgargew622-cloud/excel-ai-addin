/**
 * Чтение zip-архива (XLSX и DOCX — это zip с XML внутри) без сторонних
 * библиотек (этап 8, 8.6).
 *
 * Файл пользователя — недоверенные данные. Поэтому распаковка ограничена:
 * число записей, размер каждой распакованной записи и суммарный размер —
 * так «zip-бомба» (маленький архив, огромная распаковка) не съест память.
 * Предел проверяет сам zlib (maxOutputLength), а не заявленный в архиве
 * размер: заголовок могут подделать.
 */

import { inflateRawSync } from "node:zlib";

export const ZIP_LIMITS = { entries: 2_000, entryBytes: 50 * 1024 * 1024, totalBytes: 150 * 1024 * 1024 };

export class FileParseError extends Error {}

interface ZipEntry {
  name: string;
  method: number;
  compressedSize: number;
  offset: number;
}

function zipEntries(buffer: Buffer): ZipEntry[] {
  // Конец центрального каталога — в последних 64 КБ + 22 байта.
  const start = Math.max(0, buffer.length - 65_557);
  let end = -1;
  for (let i = buffer.length - 22; i >= start; i--) {
    if (buffer.readUInt32LE(i) === 0x06054b50) { end = i; break; }
  }
  if (end < 0) throw new FileParseError("Файл не похож на XLSX/DOCX: не найден каталог архива. Возможно, он повреждён или это другой формат.");
  const count = buffer.readUInt16LE(end + 10);
  const directory = buffer.readUInt32LE(end + 16);
  if (count > ZIP_LIMITS.entries) throw new FileParseError(`В архиве ${count} записей — больше ${ZIP_LIMITS.entries}: такой файл не разбирается.`);
  const entries: ZipEntry[] = [];
  let at = directory;
  for (let n = 0; n < count; n++) {
    if (at + 46 > buffer.length || buffer.readUInt32LE(at) !== 0x02014b50) throw new FileParseError("Каталог архива повреждён.");
    const nameLength = buffer.readUInt16LE(at + 28);
    entries.push({
      method: buffer.readUInt16LE(at + 10),
      compressedSize: buffer.readUInt32LE(at + 20),
      offset: buffer.readUInt32LE(at + 42),
      name: buffer.toString("utf8", at + 46, at + 46 + nameLength)
    });
    at += 46 + nameLength + buffer.readUInt16LE(at + 30) + buffer.readUInt16LE(at + 32);
  }
  return entries;
}

/** Читатель записей с общим пределом распакованного объёма. */
export function zipReader(buffer: Buffer) {
  const byName = new Map(zipEntries(buffer).map((entry) => [entry.name.replace(/^\/+/, ""), entry]));
  let total = 0;
  return {
    names: () => [...byName.keys()],
    text(name: string): string | null {
      const entry = byName.get(name);
      if (!entry) return null;
      const local = entry.offset;
      if (local + 30 > buffer.length || buffer.readUInt32LE(local) !== 0x04034b50) throw new FileParseError(`Запись ${name} в архиве повреждена.`);
      const dataStart = local + 30 + buffer.readUInt16LE(local + 26) + buffer.readUInt16LE(local + 28);
      const data = buffer.subarray(dataStart, dataStart + entry.compressedSize);
      let out: Buffer;
      if (entry.method === 0) out = Buffer.from(data);
      else if (entry.method === 8) {
        try {
          out = inflateRawSync(data, { maxOutputLength: ZIP_LIMITS.entryBytes });
        } catch (error: any) {
          throw new FileParseError(error?.code === "ERR_BUFFER_TOO_LARGE"
            ? `Запись ${name} распаковывается больше чем в ${ZIP_LIMITS.entryBytes / 1024 / 1024} МБ — такой файл не разбирается.`
            : `Запись ${name} не распаковалась: файл повреждён.`);
        }
      } else throw new FileParseError(`Запись ${name} сжата неизвестным способом (${entry.method}).`);
      total += out.length;
      if (total > ZIP_LIMITS.totalBytes) throw new FileParseError(`Архив распаковывается больше чем в ${ZIP_LIMITS.totalBytes / 1024 / 1024} МБ — такой файл не разбирается.`);
      return out.toString("utf8");
    }
  };
}

/** Текст XML: сущности → символы. */
export function xmlText(xml: string): string {
  return xml.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-f]+);/gi, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec) => String.fromCodePoint(Number(dec)))
    .replace(/&amp;/g, "&");
}

/** Атрибут XML-тега. */
export function attr(tag: string, name: string): string | null {
  const match = new RegExp(`\\s${name}="([^"]*)"`).exec(tag);
  return match ? xmlText(match[1]) : null;
}
