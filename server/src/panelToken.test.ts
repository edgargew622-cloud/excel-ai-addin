import test from "node:test";
import assert from "node:assert/strict";
import express from "express";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import type { AddressInfo } from "node:net";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { loadOrCreatePanelToken, PANEL_TOKEN_HEADER, requirePanelToken, tokenMatches } from "./panelToken.js";

function tempDir(t: any) {
  const dir = mkdtempSync(join(tmpdir(), "panel-token-"));
  t.after(() => rmSync(dir, { recursive: true, force: true }));
  return dir;
}

test("the token is created once, random and long, and then reused", (t) => {
  const file = join(tempDir(t), "panel-token");
  const first = loadOrCreatePanelToken(file);
  assert.match(first, /^[A-Za-z0-9_-]{43}$/);
  assert.equal(loadOrCreatePanelToken(file), first);
  assert.notEqual(loadOrCreatePanelToken(join(tempDir(t), "panel-token")), first);
});

test("a token written by the installer script is used as is; a broken file is refused", (t) => {
  const dir = tempDir(t);
  const file = join(dir, "panel-token");
  const fromScript = "Abcdefghijklmnopqrstuvwxyz0123456789_-ABCDE";
  writeFileSync(file, `${fromScript}\r\n`);
  assert.equal(loadOrCreatePanelToken(file), fromScript);
  writeFileSync(file, "короткий");
  assert.throws(() => loadOrCreatePanelToken(file), /повреждён/);
});

test("only the exact token passes", () => {
  const token = "Abcdefghijklmnopqrstuvwxyz0123456789_-ABCDE";
  assert.equal(tokenMatches(token, token), true);
  assert.equal(tokenMatches(token, token.slice(0, -1)), false);
  assert.equal(tokenMatches(token, `${token}x`), false);
  assert.equal(tokenMatches(token, undefined), false);
  assert.equal(tokenMatches(token, ["x"]), false);
});

test("the API answers only with the token, before the body is even parsed; health stays open", async (t) => {
  const token = loadOrCreatePanelToken(join(tempDir(t), "panel-token"));
  const app = express();
  let parsed = 0;
  app.use("/api", requirePanelToken(token));
  app.use(express.json());
  app.use((req, _res, next) => { if (req.body && Object.keys(req.body).length) parsed += 1; next(); });
  app.get("/api/health", (_req, res) => res.json({ app: "excel-ai-addin" }));
  app.post("/api/chat", (_req, res) => res.json({ ok: true }));
  app.put("/api/keys/:id", (_req, res) => res.json({ ok: true }));
  const server = app.listen(0, "127.0.0.1");
  await new Promise((resolve) => server.once("listening", resolve));
  t.after(() => server.close());
  const base = `http://127.0.0.1:${(server.address() as AddressInfo).port}`;
  const call = (method: string, path: string, headers: Record<string, string> = {}) =>
    fetch(`${base}${path}`, { method, headers: { "Content-Type": "application/json", ...headers }, ...(method === "GET" ? {} : { body: JSON.stringify({ key: "sk-чужой" }) }) });

  assert.equal((await call("GET", "/api/health")).status, 200);
  assert.equal((await call("POST", "/api/chat")).status, 401);
  assert.equal((await call("PUT", "/api/keys/deepseek", { [PANEL_TOKEN_HEADER]: "guessed-wrong-token" })).status, 401);
  assert.equal(parsed, 0, "тело чужого запроса не разбиралось");
  assert.equal((await call("POST", "/api/chat", { [PANEL_TOKEN_HEADER]: token })).status, 200);
  assert.equal((await call("PUT", "/api/keys/deepseek", { [PANEL_TOKEN_HEADER]: token })).status, 200);
});
