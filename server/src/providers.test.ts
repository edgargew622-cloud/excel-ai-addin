import test from "node:test";
import assert from "node:assert/strict";
import { availableProviders, PROVIDERS, providerBaseURL, providerModels, providerReady, setDetectedModels } from "./providers.js";

const openrouter = PROVIDERS.find((provider) => provider.id === "openrouter")!;

test("openrouter is enabled with a default that is actually in its list", () => {
  assert.equal(openrouter.enabled, true);
  assert.ok(openrouter.models.includes(openrouter.defaultModel), openrouter.defaultModel);
});

test("every family is represented by four models, Claude by six", () => {
  const family = (prefix: string) => openrouter.models.filter((model) => model.startsWith(prefix));
  // Пятая у Claude — opus-5.5 (24 сентября 2026 года), шестая — sonnet-5.5
  // (29 сентября): обе по просьбе пользователя сверх отобранных четырёх.
  assert.equal(family("anthropic/").length, 6);
  assert.equal(family("google/").length, 4);
  assert.equal(family("mistralai/").length, 4);
  assert.equal(openrouter.models.filter((model) => model.endsWith(":free")).length, 3);
  assert.equal(openrouter.models.length, 17);
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
  for (const provider of PROVIDERS.filter((item) => item.enabled && !item.detectModels)) {
    assert.ok(provider.models.length > 0, provider.id);
    assert.ok(provider.models.includes(provider.defaultModel), `${provider.id}: ${provider.defaultModel}`);
  }
});

const qwen = PROVIDERS.find((provider) => provider.id === "qwen")!;

test("the own Qwen server needs no key but must be switched on by its address", () => {
  const saved = { url: process.env.QWEN_BASE_URL, key: process.env.QWEN_API_KEY };
  try {
    delete process.env.QWEN_BASE_URL;
    delete process.env.QWEN_API_KEY;
    // Без адреса его нет в списке: иначе он был бы у всех, у кого порт пуст.
    assert.equal(providerReady(qwen), false);
    assert.equal(availableProviders().some((p) => p.id === "qwen"), false);

    process.env.QWEN_BASE_URL = "http://127.0.0.1:8080/v1/";
    assert.equal(providerReady(qwen), true, "ключ не нужен");
    assert.equal(providerBaseURL(qwen), "http://127.0.0.1:8080/v1", "лишний слеш в конце убран");
    const listed = availableProviders().find((p) => p.id === "qwen") as any;
    assert.equal(listed.taskBudgetMinutes, 30, "медленной модели — больше времени на задачу");
  } finally {
    if (saved.url === undefined) delete process.env.QWEN_BASE_URL; else process.env.QWEN_BASE_URL = saved.url;
    if (saved.key === undefined) delete process.env.QWEN_API_KEY; else process.env.QWEN_API_KEY = saved.key;
  }
});

test("cloud providers are still switched on by their key alone", () => {
  const deepseek = PROVIDERS.find((provider) => provider.id === "deepseek")!;
  const saved = process.env.DEEPSEEK_API_KEY;
  try {
    delete process.env.DEEPSEEK_API_KEY;
    assert.equal(providerReady(deepseek), false);
    process.env.DEEPSEEK_API_KEY = "ключ";
    assert.equal(providerReady(deepseek), true);
    assert.equal(providerBaseURL(deepseek), "https://api.deepseek.com");
  } finally {
    if (saved === undefined) delete process.env.DEEPSEEK_API_KEY; else process.env.DEEPSEEK_API_KEY = saved;
  }
});

test("Ollama switches itself on only when a model with tools is found on this computer", () => {
  const ollama = PROVIDERS.find((provider) => provider.id === "ollama")!;
  try {
    setDetectedModels("ollama", []);
    assert.equal(providerReady(ollama), false, "Ollama нет или моделей нет — провайдера нет в списке");
    assert.equal(availableProviders().some((p) => p.id === "ollama"), false);

    setDetectedModels("ollama", ["qwen3:8b", "llama3.2:3b"]);
    assert.equal(providerReady(ollama), true, "ключ и server/.env не нужны");
    const listed = availableProviders().find((p) => p.id === "ollama") as any;
    assert.deepEqual(listed.models, ["qwen3:8b", "llama3.2:3b"]);
    assert.equal(listed.defaultModel, "qwen3:8b", "по умолчанию — первая найденная");
    assert.equal(listed.taskBudgetMinutes, 30, "на процессоре задача идёт долго");
    assert.deepEqual(providerModels(ollama), ["qwen3:8b", "llama3.2:3b"]);
  } finally {
    setDetectedModels("ollama", []);
  }
});
