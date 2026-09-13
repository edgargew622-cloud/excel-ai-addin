import test from "node:test";
import assert from "node:assert/strict";
import { isAllowedOrigin, isLoopbackAddress } from "./localOnly.js";

test("loopback is recognized in every form Node reports", () => {
  for (const address of [
    "127.0.0.1",
    "::1",
    "::ffff:127.0.0.1",
    "::1%lo0",
    "127.0.0.53",
    "127.1.2.3",
    " 127.0.0.1 ",
    "::FFFF:127.0.0.1"
  ]) {
    assert.equal(isLoopbackAddress(address), true, address);
  }
});

test("non-loopback addresses are rejected", () => {
  for (const address of [
    "192.168.1.10",
    "10.0.0.1",
    "0.0.0.0",
    "::",
    "128.0.0.1",
    "227.0.0.1",
    "::ffff:192.168.1.10",
    "example.com",
    "",
    undefined,
    null
  ]) {
    assert.equal(isLoopbackAddress(address as string | undefined), false, String(address));
  }
});

test("malformed addresses do not pass as loopback", () => {
  for (const address of ["127.0.0", "127.0.0.256", "127.0.0.1.5", "127.a.b.c", "127-0-0-1"]) {
    assert.equal(isLoopbackAddress(address), false, address);
  }
});

test("absent Origin is allowed: own navigation and static GET send none", () => {
  assert.equal(isAllowedOrigin(undefined, 3000), true);
  assert.equal(isAllowedOrigin("", 3000), true);
});

test("own origin is allowed on the serving port only", () => {
  assert.equal(isAllowedOrigin("https://localhost:3000", 3000), true);
  assert.equal(isAllowedOrigin("https://127.0.0.1:3000", 3000), true);
  assert.equal(isAllowedOrigin("https://[::1]:3000", 3000), true);
  // Другой порт — это уже другое происхождение, включая dev-панель на 3100.
  assert.equal(isAllowedOrigin("https://localhost:3100", 3000), false);
});

test("foreign and downgraded origins are rejected", () => {
  for (const origin of [
    "https://evil.example",
    "http://localhost:3000",
    "https://localhost",
    "null",
    "not a url",
    "https://localhost.evil.example:3000"
  ]) {
    assert.equal(isAllowedOrigin(origin, 3000), false, origin);
  }
});
