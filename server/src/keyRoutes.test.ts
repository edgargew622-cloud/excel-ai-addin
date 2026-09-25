import test from "node:test";
import assert from "node:assert/strict";
import express from "express";
import { mkdtempSync, rmSync } from "node:fs";
import type { AddressInfo } from "node:net";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { KeyStore, type Protector } from "./keyStore.js";
import { registerKeyRoutes } from "./keyRoutes.js";
import { availableProviders, setStoredKeyLookup } from "./providers.js";

const KEY = "sk-panel-0123456789wxyz";
const ENV_KEY = "sk-env-0123456789abcd";

function flipProtector(): Protector {
  const flip = (data: Buffer) => Buffer.from(data.map((byte) => byte ^ 0x33));
  return { available: true, protect: async (b) => flip(b), unprotect: async (b) => flip(b) };
}

async function startApp(t: any, protector: Protector = flipProtector()) {
  const dir = mkdtempSync(join(tmpdir(), "key-routes-"));
  const store = new KeyStore(join(dir, "keys.dpapi"), protector);
  await store.load();
  setStoredKeyLookup((id) => store.get(id));
  const app = express();
  app.use(express.json());
  registerKeyRoutes(app, store);
  const server = app.listen(0, "127.0.0.1");
  await new Promise((resolve) => server.once("listening", resolve));
  const base = `http://127.0.0.1:${(server.address() as AddressInfo).port}`;
  const saved = { DEEPSEEK_API_KEY: process.env.DEEPSEEK_API_KEY, OPENAI_API_KEY: process.env.OPENAI_API_KEY };
  delete process.env.DEEPSEEK_API_KEY;
  process.env.OPENAI_API_KEY = ENV_KEY;
  t.after(() => {
    server.close();
    setStoredKeyLookup(() => undefined);
    for (const [name, value] of Object.entries(saved)) {
      if (value === undefined) delete process.env[name];
      else process.env[name] = value;
    }
    rmSync(dir, { recursive: true, force: true });
  });
  const call = async (method: string, path: string, body?: unknown) => {
    const res = await fetch(`${base}${path}`, {
      method,
      headers: { "Content-Type": "application/json" },
      ...(body === undefined ? {} : { body: JSON.stringify(body) })
    });
    const text = await res.text();
    return { status: res.status, text, json: JSON.parse(text) };
  };
  return { call, store };
}

const status = (json: any, id: string) => json.providers.find((p: any) => p.id === id);

test("a key saved in the panel makes the provider ready and is never sent back", async (t) => {
  const { call } = await startApp(t);

  const before = await call("GET", "/api/keys");
  assert.equal(before.json.storage.available, true);
  assert.deepEqual(status(before.json, "deepseek"), {
    id: "deepseek", label: "DeepSeek", source: null, hint: null, ready: false, keyOptional: false
  });
  assert.equal(availableProviders().some((p) => p.id === "deepseek"), false);

  const saved = await call("PUT", "/api/keys/deepseek", { key: KEY });
  assert.equal(saved.status, 200);
  assert.equal(saved.text.includes(KEY), false);
  assert.deepEqual(status(saved.json, "deepseek"), {
    id: "deepseek", label: "DeepSeek", source: "panel", hint: "…wxyz", ready: true, keyOptional: false
  });
  assert.equal(availableProviders().some((p) => p.id === "deepseek"), true);

  const listed = await call("GET", "/api/keys");
  assert.equal(listed.text.includes(KEY), false);
  assert.equal(listed.text.includes(ENV_KEY), false);
});

test("a panel key takes precedence over server/.env, and removing it falls back to .env", async (t) => {
  const { call } = await startApp(t);
  assert.equal(status((await call("GET", "/api/keys")).json, "openai").source, "env");
  assert.equal(status((await call("GET", "/api/keys")).json, "openai").hint, "…abcd");

  await call("PUT", "/api/keys/openai", { key: KEY });
  assert.equal(status((await call("GET", "/api/keys")).json, "openai").source, "panel");

  const removed = await call("DELETE", "/api/keys/openai");
  assert.equal(removed.status, 200);
  assert.equal(status(removed.json, "openai").source, "env");
});

test("unknown providers and malformed keys are refused", async (t) => {
  const { call, store } = await startApp(t);
  assert.equal((await call("PUT", "/api/keys/nope", { key: KEY })).status, 404);
  assert.equal((await call("DELETE", "/api/keys/nope")).status, 404);
  const bad = await call("PUT", "/api/keys/deepseek", { key: "sk bad key 123" });
  assert.equal(bad.status, 400);
  assert.match(bad.json.error.message, /пробелы/);
  assert.equal((await call("PUT", "/api/keys/deepseek", {})).status, 400);
  assert.equal(store.get("deepseek"), undefined);
});

test("an encryption failure is reported without its details", async (t) => {
  const broken: Protector = {
    available: true,
    protect: async () => { throw new Error(`секретная подробность ${KEY}`); },
    unprotect: async (b) => b
  };
  const { call } = await startApp(t, broken);
  const originalError = console.error;
  console.error = () => undefined;
  t.after(() => { console.error = originalError; });
  const res = await call("PUT", "/api/keys/deepseek", { key: KEY });
  assert.equal(res.status, 500);
  assert.equal(res.text.includes(KEY), false);
  assert.match(res.json.error.message, /журнале сервера/);
});
