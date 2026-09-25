import test from "node:test";
import assert from "node:assert/strict";
import express from "express";
import { existsSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import type { AddressInfo } from "node:net";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { registerBackupRoutes } from "./backupRoutes.js";

async function startApp(t: any) {
  const root = mkdtempSync(join(tmpdir(), "backup-routes-"));
  const app = express();
  app.use(express.json());
  registerBackupRoutes(app, root);
  const server = app.listen(0, "127.0.0.1");
  await new Promise((resolve) => server.once("listening", resolve));
  const base = `http://127.0.0.1:${(server.address() as AddressInfo).port}`;
  t.after(() => {
    server.close();
    rmSync(root, { recursive: true, force: true });
  });
  const post = async (path: string, body: unknown) => {
    const res = await fetch(`${base}${path}`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
    return { status: res.status, json: (await res.json()) as any };
  };
  return { root, post };
}

test("abort cannot reach a file outside incoming/ through a crafted upload id", async (t) => {
  const { root, post } = await startApp(t);
  // Аудит 24 сентября 2026 года (SEC-04): abort удалял <incoming>/<id>.part без
  // проверки, что такая загрузка есть, а join нормализует «..».
  mkdirSync(join(root, "logs"), { recursive: true });
  const victim = join(root, "logs", "victim.part");
  writeFileSync(victim, "не трогать");
  for (const uploadId of ["..\\..\\logs\\victim", "../../logs/victim"]) {
    const res = await post("/api/backup/abort", { uploadId });
    assert.equal(existsSync(victim), true, `файл вне incoming/ удалён по id ${uploadId}`);
    assert.equal(res.status, 404, `id ${uploadId}`);
  }
  assert.equal(existsSync(victim), true);
});

test("abort removes the caller's own active upload", async (t) => {
  const { root, post } = await startApp(t);
  const begun = await post("/api/backup/begin", { fileName: "Книга.xlsx", sizeBytes: 10, sliceCount: 1 });
  assert.equal(begun.status, 200);
  const id = begun.json.uploadId;
  const part = join(root, "backups", "incoming", `${id}.part`);
  assert.equal(existsSync(part), true);
  const res = await post("/api/backup/abort", { uploadId: id });
  assert.equal(res.status, 200);
  assert.equal(existsSync(part), false);
});
