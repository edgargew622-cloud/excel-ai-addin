/**
 * Ключи провайдеров, введённые в панели.
 *
 * Ключ никогда не покидает этот компьютер и никогда не отдаётся обратно в
 * панель: наружу выходит только признак наличия и последние символы. На диске
 * лежит только шифротекст; расшифровать его может только та же учётная запись
 * Windows (DPAPI, см. dpapi.ts). Логика хранения отделена от шифрования, чтобы
 * её можно было проверить без Windows.
 */

import { existsSync, readFileSync, renameSync, writeFileSync } from "node:fs";

/**
 * Замена файла, которая переживает кратковременную блокировку в Windows.
 *
 * Антивирус или индексатор открывают только что записанный файл на доли
 * секунды, и переименование поверх него получает EPERM, EBUSY или EACCES.
 * Так изредка падал тест «удаление одного ключа не трогает остальные», а у
 * пользователя «Сохранить» отвечало бы ошибкой. Повторяем с короткой паузой.
 */
export async function replaceFile(
  from: string,
  to: string,
  rename: (a: string, b: string) => void = renameSync,
  attempts = 8
): Promise<void> {
  for (let attempt = 1; ; attempt++) {
    try {
      rename(from, to);
      return;
    } catch (error: any) {
      const transient = error?.code === "EPERM" || error?.code === "EBUSY" || error?.code === "EACCES";
      if (!transient || attempt >= attempts) throw error;
      await new Promise((resolve) => setTimeout(resolve, 25 * attempt));
    }
  }
}

export interface Protector {
  /** Есть ли на этой системе защищённое хранилище. */
  readonly available: boolean;
  protect(plain: Buffer): Promise<Buffer>;
  unprotect(blob: Buffer): Promise<Buffer>;
}

export class KeyError extends Error {}

const FORMAT = "excel-ai-keys";
const VERSION = 1;
const MIN_KEY_LENGTH = 8;
const MAX_KEY_LENGTH = 512;

/** Ключ — одна строка без пробелов и управляющих символов. */
export function validateKey(raw: unknown): string {
  if (typeof raw !== "string") throw new KeyError("Ключ должен быть строкой.");
  const key = raw.trim();
  if (!key) throw new KeyError("Ключ пуст.");
  if (/[\s\u0000-\u001f\u007f]/.test(key)) {
    throw new KeyError("В ключе есть пробелы или переводы строки — скопируйте его заново целиком.");
  }
  if (key.length < MIN_KEY_LENGTH || key.length > MAX_KEY_LENGTH) {
    throw new KeyError(`Длина ключа должна быть от ${MIN_KEY_LENGTH} до ${MAX_KEY_LENGTH} символов.`);
  }
  return key;
}

/** Достаточно, чтобы узнать свой ключ, и недостаточно, чтобы им воспользоваться. */
export function keyHint(key: string): string {
  return `…${key.slice(-4)}`;
}

export class KeyStore {
  private keys = new Map<string, string>();
  private queue: Promise<unknown> = Promise.resolve();
  /** Почему сохранённые ключи не прочитаны. Текст без содержимого файла. */
  loadError: string | null = null;

  constructor(private readonly file: string, private readonly protector: Protector) {}

  get storageAvailable(): boolean {
    return this.protector.available;
  }

  async load(): Promise<void> {
    this.keys.clear();
    this.loadError = null;
    if (!existsSync(this.file)) return;
    try {
      const envelope = JSON.parse(readFileSync(this.file, "utf8"));
      if (envelope?.format !== FORMAT || envelope?.version !== VERSION || typeof envelope?.data !== "string") {
        throw new Error("неизвестный формат файла");
      }
      if (!this.protector.available) throw new Error("защищённое хранилище недоступно на этой системе");
      const plain = await this.protector.unprotect(Buffer.from(envelope.data, "base64"));
      const payload = JSON.parse(plain.toString("utf8"));
      for (const [id, key] of Object.entries(payload?.keys ?? {})) {
        if (typeof key === "string" && key) this.keys.set(id, key);
      }
    } catch (error: any) {
      this.keys.clear();
      this.loadError =
        `Сохранённые в панели ключи не прочитаны (${error?.message ?? error}). ` +
        "Файл не тронут; введите ключи заново, и он будет перезаписан.";
    }
  }

  get(id: string): string | undefined {
    return this.keys.get(id);
  }

  async set(id: string, raw: unknown): Promise<void> {
    const key = validateKey(raw);
    return this.enqueue(async () => {
      const next = new Map(this.keys);
      next.set(id, key);
      await this.persist(next);
      this.keys = next;
    });
  }

  remove(id: string): Promise<boolean> {
    return this.enqueue(async () => {
      if (!this.keys.has(id)) return false;
      const next = new Map(this.keys);
      next.delete(id);
      await this.persist(next);
      this.keys = next;
      return true;
    });
  }

  /** Записи по очереди: две параллельные не должны потерять ключ друг друга. */
  private enqueue<T>(job: () => Promise<T>): Promise<T> {
    const run = this.queue.then(job, job);
    this.queue = run.catch(() => undefined);
    return run;
  }

  private async persist(keys: Map<string, string>): Promise<void> {
    if (!this.protector.available) {
      throw new KeyError(
        "Сохранять ключи из панели можно только в Windows: там они шифруются средствами системы. " +
          "Здесь впишите ключ в server/.env."
      );
    }
    const plain = Buffer.from(JSON.stringify({ keys: Object.fromEntries(keys) }), "utf8");
    const blob = await this.protector.protect(plain);
    const envelope = { format: FORMAT, version: VERSION, scope: "dpapi-current-user", data: blob.toString("base64") };
    // Сначала целиком во временный файл, затем замена: оборванная запись не
    // должна оставить полфайла вместо всех ключей.
    const temp = `${this.file}.tmp`;
    writeFileSync(temp, JSON.stringify(envelope), { encoding: "utf8", mode: 0o600 });
    await replaceFile(temp, this.file);
  }
}
