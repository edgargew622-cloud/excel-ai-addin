import test from "node:test";
import assert from "node:assert/strict";
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { KeyError, KeyStore, keyHint, replaceFile, validateKey, type Protector } from "./keyStore.js";
import { dpapiArgs, dpapiScript, DPAPI_INPUT_VARIABLE, unavailableProtector, windowsDpapi } from "./dpapi.js";

/** Обратимое «шифрование» для тестов: без Windows DPAPI недоступен. */
function fakeProtector(): Protector & { calls: number } {
  const flip = (data: Buffer) => Buffer.from(data.map((byte) => byte ^ 0x5a));
  const protector = {
    available: true,
    calls: 0,
    protect: async (plain: Buffer) => { protector.calls += 1; return Buffer.concat([Buffer.from("FAKE"), flip(plain)]); },
    unprotect: async (blob: Buffer) => {
      protector.calls += 1;
      if (blob.subarray(0, 4).toString() !== "FAKE") throw new Error("чужой шифротекст");
      return flip(blob.subarray(4));
    }
  };
  return protector;
}

function tempFile(t: { after: (fn: () => void) => void }) {
  const dir = mkdtempSync(join(tmpdir(), "keys-"));
  t.after(() => rmSync(dir, { recursive: true, force: true }));
  return join(dir, "keys.dpapi");
}

const KEY = "sk-test-0123456789abcdef";

test("a saved key survives a restart and never lands on disk in plain text", async (t) => {
  const file = tempFile(t);
  const store = new KeyStore(file, fakeProtector());
  await store.load();
  await store.set("deepseek", `  ${KEY}\n`);
  assert.equal(store.get("deepseek"), KEY);

  const onDisk = readFileSync(file, "utf8");
  assert.equal(onDisk.includes(KEY), false);
  assert.equal(onDisk.includes("deepseek"), false, "даже имена провайдеров внутри шифротекста");
  assert.equal(existsSync(`${file}.tmp`), false);

  const restarted = new KeyStore(file, fakeProtector());
  await restarted.load();
  assert.equal(restarted.get("deepseek"), KEY);
  assert.equal(restarted.loadError, null);
});

test("removing one key keeps the others", async (t) => {
  const store = new KeyStore(tempFile(t), fakeProtector());
  await store.set("deepseek", KEY);
  await store.set("openai", `${KEY}-2`);
  assert.equal(await store.remove("deepseek"), true);
  assert.equal(await store.remove("deepseek"), false);
  assert.equal(store.get("deepseek"), undefined);
  assert.equal(store.get("openai"), `${KEY}-2`);
});

test("parallel saves do not lose each other's keys", async (t) => {
  const file = tempFile(t);
  const store = new KeyStore(file, fakeProtector());
  await Promise.all([store.set("deepseek", KEY), store.set("openai", `${KEY}-2`), store.set("xai", `${KEY}-3`)]);
  const restarted = new KeyStore(file, fakeProtector());
  await restarted.load();
  assert.deepEqual(["deepseek", "openai", "xai"].map((id) => restarted.get(id)), [KEY, `${KEY}-2`, `${KEY}-3`]);
});

test("a key with spaces, line breaks or a wrong length is refused before anything is written", async (t) => {
  const file = tempFile(t);
  const protector = fakeProtector();
  const store = new KeyStore(file, protector);
  for (const bad of ["", "   ", "sk-abc def-0123456", "sk-abc\ndef-0123456", "short", "x".repeat(513), 42, null]) {
    await assert.rejects(() => store.set("deepseek", bad), KeyError, String(bad));
  }
  assert.equal(protector.calls, 0);
  assert.equal(existsSync(file), false);
  assert.equal(validateKey(`\t${KEY} `), KEY);
});

test("the hint shows only the last four characters", () => {
  assert.equal(keyHint(KEY), "…cdef");
});

test("an unreadable file is reported, left untouched, and replaced only on the next save", async (t) => {
  const file = tempFile(t);
  const foreign = JSON.stringify({ format: "excel-ai-keys", version: 1, data: Buffer.from("чужое").toString("base64") });
  writeFileSync(file, foreign);

  const store = new KeyStore(file, fakeProtector());
  await store.load();
  assert.match(store.loadError ?? "", /не прочитаны.*введите ключи заново/);
  assert.equal(store.get("deepseek"), undefined);
  assert.equal(readFileSync(file, "utf8"), foreign);

  await store.set("deepseek", KEY);
  const restarted = new KeyStore(file, fakeProtector());
  await restarted.load();
  assert.equal(restarted.get("deepseek"), KEY);
});

test("without Windows the panel cannot store keys and says to use server/.env", async (t) => {
  const file = tempFile(t);
  const store = new KeyStore(file, unavailableProtector());
  await store.load();
  assert.equal(store.storageAvailable, false);
  await assert.rejects(() => store.set("deepseek", KEY), (error: any) => {
    assert.ok(error instanceof KeyError);
    assert.match(error.message, /только в Windows.*server\/\.env/);
    return true;
  });
  assert.equal(existsSync(file), false);
});

test("the key reaches PowerShell only through the environment, never the command line", () => {
  for (const mode of ["protect", "unprotect"] as const) {
    const args = dpapiArgs(mode);
    const encoded = args[args.indexOf("-EncodedCommand") + 1];
    assert.equal(Buffer.from(encoded, "base64").toString("utf16le"), dpapiScript(mode));
    assert.match(dpapiScript(mode), new RegExp(`\\$env:${DPAPI_INPUT_VARIABLE}`));
    assert.match(dpapiScript(mode), /DataProtectionScope\]::CurrentUser/);
  }
  assert.match(dpapiScript("protect"), /ProtectedData\]::Protect\(/);
  assert.match(dpapiScript("unprotect"), /ProtectedData\]::Unprotect\(/);
});

test("real DPAPI round-trips a key and the file holds only ciphertext", { skip: process.platform !== "win32" && "только в Windows" }, async (t) => {
  const file = tempFile(t);
  const store = new KeyStore(file, windowsDpapi());
  await store.set("deepseek", KEY);
  assert.equal(readFileSync(file, "utf8").includes(KEY), false);

  const restarted = new KeyStore(file, windowsDpapi());
  await restarted.load();
  assert.equal(restarted.loadError, null);
  assert.equal(restarted.get("deepseek"), KEY);
});

test("a file locked for a moment by Windows is replaced after a short retry", async () => {
  let calls = 0;
  const flaky = () => {
    calls += 1;
    if (calls < 3) throw Object.assign(new Error("operation not permitted"), { code: "EPERM" });
  };
  await replaceFile("a", "b", flaky);
  assert.equal(calls, 3);
  // Настоящая ошибка — не блокировка: повторять нечего.
  const missing = () => { throw Object.assign(new Error("no such file"), { code: "ENOENT" }); };
  await assert.rejects(() => replaceFile("a", "b", missing), /no such file/);
  // Блокировка, которая не проходит, не зацикливает сохранение.
  let tries = 0;
  const stuck = () => { tries += 1; throw Object.assign(new Error("busy"), { code: "EBUSY" }); };
  await assert.rejects(() => replaceFile("a", "b", stuck, 3), /busy/);
  assert.equal(tries, 3);
});
