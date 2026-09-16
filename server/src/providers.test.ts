import test from "node:test";
import assert from "node:assert/strict";
import { PROVIDERS } from "./providers.js";

const openrouter = PROVIDERS.find((provider) => provider.id === "openrouter")!;

test("openrouter is enabled with a default that is actually in its list", () => {
  assert.equal(openrouter.enabled, true);
  assert.ok(openrouter.models.includes(openrouter.defaultModel), openrouter.defaultModel);
});

test("every family is represented by four models", () => {
  const family = (prefix: string) => openrouter.models.filter((model) => model.startsWith(prefix));
  assert.equal(family("anthropic/").length, 4);
  assert.equal(family("google/").length, 4);
  assert.equal(family("mistralai/").length, 4);
  assert.equal(openrouter.models.filter((model) => model.endsWith(":free")).length, 4);
  assert.equal(openrouter.models.length, 16);
});

test("no batch variants: they are asynchronous and useless for an interactive pane", () => {
  for (const model of openrouter.models) assert.equal(model.includes(":batch"), false, model);
});

test("models rejected during the live check are not offered", () => {
  // Отвечали текстом даже при tool_choice required, либо недоступны вовсе.
  for (const model of [
    "google/gemini-3.5-flash",
    "google/gemini-3.1-pro-preview",
    "thinkingmachines/inkling:free",
    "nex-agi/nex-n2.5-mini:free"
  ]) {
    assert.equal(openrouter.models.includes(model), false, model);
  }
});

test("every provider keeps a default model it actually offers", () => {
  for (const provider of PROVIDERS.filter((item) => item.enabled)) {
    assert.ok(provider.models.length > 0, provider.id);
    assert.ok(provider.models.includes(provider.defaultModel), `${provider.id}: ${provider.defaultModel}`);
  }
});
