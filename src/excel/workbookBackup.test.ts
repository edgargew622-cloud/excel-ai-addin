import test from "node:test";
import assert from "node:assert/strict";
import { createWorkbookBackup, toBase64 } from "./workbookBackup";

/** Макет Office и сервера: книга из нескольких срезов и запись вызовов. */
function stage(options: {
  size: number;
  sliceCount: number;
  failAtSlice?: number;
  serverFailsAt?: string;
  noFile?: boolean;
}) {
  const calls: { path: string; body: any }[] = [];
  let closed = false;

  const file = {
    size: options.size,
    sliceCount: options.sliceCount,
    getSliceAsync: (index: number, callback: (result: any) => void) => {
      if (options.failAtSlice === index) {
        callback({ status: "failed", error: { message: "срез потерян" } });
        return;
      }
      callback({ status: "succeeded", value: { data: [1, 2, 3] } });
    },
    closeAsync: (callback?: (result: any) => void) => { closed = true; callback?.({}); }
  };

  (globalThis as any).Office = {
    FileType: { Compressed: "compressed" },
    context: {
      document: options.noFile ? { url: "C:/книги/Отчёт.xlsm" } : {
        url: "C:/книги/Отчёт.xlsm",
        getFileAsync: (_t: string, _o: unknown, cb: (result: any) => void) =>
          cb({ status: "succeeded", value: file })
      }
    }
  };

  (globalThis as any).fetch = async (path: string, init: any) => {
    const body = JSON.parse(init.body);
    calls.push({ path, body });
    if (options.serverFailsAt && path.includes(options.serverFailsAt)) {
      return { ok: false, status: 400, json: async () => ({ error: { message: "сервер отказал" } }) };
    }
    if (path.endsWith("/begin")) return { ok: true, status: 200, json: async () => ({ uploadId: "u1" }) };
    if (path.endsWith("/slice")) return { ok: true, status: 200, json: async () => ({ nextIndex: body.index + 1 }) };
    if (path.endsWith("/finish")) {
      return { ok: true, status: 200, json: async () => ({ name: "20260916T000000Z-Отчёт.xlsm", directory: "C:/backups", sizeBytes: options.size, removed: [] }) };
    }
    return { ok: true, status: 200, json: async () => ({ ok: true }) };
  };

  return { calls, wasClosed: () => closed };
}

test("a full backup sends every slice in order and reports where it landed", async () => {
  const s = stage({ size: 9, sliceCount: 3 });
  const result = await createWorkbookBackup();

  assert.equal(result.ok, true);
  assert.equal(result.name, "20260916T000000Z-Отчёт.xlsm");
  assert.equal(result.sliceCount, 3);
  assert.deepEqual(s.calls.filter((c) => c.path.endsWith("/slice")).map((c) => c.body.index), [0, 1, 2]);
  // Имя книги уходит на сервер как есть: очищает его сервер, а не панель.
  assert.equal(s.calls[0].body.fileName, "C:/книги/Отчёт.xlsm");
  assert.equal(s.wasClosed(), true);
});

test("the copy never claims to match the file on disk", async () => {
  stage({ size: 3, sliceCount: 1 });
  const result = await createWorkbookBackup();
  // Копия снимается из открытой книги: при несохранённых правках это разные вещи.
  assert.match(result.note, /из открытой книги/);
  assert.match(result.note, /несохранённ[а-я]+ правк[а-я]+/);
  // Восстановление ручное, и перенос листа не равен восстановлению книги.
  assert.match(result.note, /Восстановление ручное/);
  assert.match(result.note, /межлистовые ссылки/);
});

test("a lost slice aborts the upload instead of leaving half a copy", async () => {
  const s = stage({ size: 9, sliceCount: 3, failAtSlice: 1 });
  const result = await createWorkbookBackup();

  assert.equal(result.ok, false);
  assert.match(result.error ?? "", /срез потерян/);
  assert.match(result.note, /не опубликована/);
  // Незавершённая загрузка закрывается, чтобы не занимать место до истечения.
  assert.equal(s.calls.some((c) => c.path.endsWith("/abort")), true);
  assert.equal(s.calls.some((c) => c.path.endsWith("/finish")), false);
  assert.equal(s.wasClosed(), true);
});

test("a server refusal is passed on plainly, with no finished copy", async () => {
  const s = stage({ size: 3, sliceCount: 1, serverFailsAt: "/finish" });
  const result = await createWorkbookBackup();

  assert.equal(result.ok, false);
  assert.match(result.error ?? "", /сервер отказал/);
  assert.equal(s.calls.some((c) => c.path.endsWith("/abort")), true);
});

test("an Excel without file export says so and sends nothing", async () => {
  const s = stage({ size: 0, sliceCount: 0, noFile: true });
  const result = await createWorkbookBackup();

  assert.equal(result.ok, false);
  assert.match(result.note, /Копия не создавалась/);
  assert.deepEqual(s.calls, []);
});

test("slice bytes are encoded as base64 for transport", () => {
  assert.equal(toBase64([80, 75, 3, 4]), "UEsDBA==");
  assert.equal(toBase64(new Uint8Array([80, 75, 3, 4])), "UEsDBA==");
  assert.equal(toBase64([]), "");
});
