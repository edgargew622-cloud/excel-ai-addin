/**
 * Веб-режим панели (Mac, Excel в браузере): модели напрямую из панели.
 *
 * Локального сервера нет, поэтому ключи хранятся в хранилище панели
 * (localStorage источника панели) на этом компьютере — это слабее, чем
 * шифрование Windows в обычном режиме, и панель прямо говорит об этом.
 *
 * Поставщики — те, что разрешают запросы из браузера (CORS; проверено
 * 30.09.2026 предварительным запросом с адреса сайта панели):
 * - DeepSeek — те же модели и то же тело запроса, что у сервера на Windows;
 * - OpenRouter — только бесплатные модели, дешёвый Mistral и Gemini Flash
 *   (выбор пользователя 30.09.2026), а не весь список Windows.
 * Формат сообщений — общий с сервером модуль.
 */

import { PROVIDERS } from "../../../server/src/providers";
import { serializeMessages, type InternalMessage } from "../../../server/src/protocol";
import type { KeysState, KeyStatus, ProviderInfo } from "./client";

export const OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions";
export const DEEPSEEK_URL = "https://api.deepseek.com/chat/completions";

const deepseek = PROVIDERS.find((p) => p.id === "deepseek")!;
const openrouter = PROVIDERS.find((p) => p.id === "openrouter")!;

/** Платные модели OpenRouter в веб-панели сверх бесплатных. */
// gemini-3.8-flash:batch не годится: OpenRouter отказывает ему в chat/completions
// (пакетная обработка с ответом позже) — проверено 30.09.2026.
export const WEB_OPENROUTER_PAID = ["mistralai/mistral-small-2603", "google/gemini-3.8-flash"];

interface WebProvider {
  id: "deepseek" | "openrouter";
  label: string;
  url: string;
  storageKey: string;
  models: string[];
  defaultModel: string;
  /** Где взять ключ — подсказка в «Ключах». */
  site: string;
}

const openrouterFree = openrouter.models.filter((model) => model.endsWith(":free"));

export const WEB_PROVIDERS: WebProvider[] = [
  {
    id: "deepseek",
    label: deepseek.label,
    url: DEEPSEEK_URL,
    storageKey: "amai.deepseek-key",
    models: deepseek.models,
    defaultModel: deepseek.defaultModel,
    site: "platform.deepseek.com"
  },
  {
    id: "openrouter",
    label: "OpenRouter (бесплатные, Mistral, Gemini)",
    url: OPENROUTER_URL,
    storageKey: "amai.openrouter-key",
    models: [...openrouterFree, ...WEB_OPENROUTER_PAID],
    defaultModel: openrouterFree[0] ?? WEB_OPENROUTER_PAID[0],
    site: "openrouter.ai/keys"
  }
];

function webProvider(id: string): WebProvider {
  const provider = WEB_PROVIDERS.find((p) => p.id === id);
  if (!provider) throw new Error("В этой панели работают только DeepSeek и OpenRouter.");
  return provider;
}

export interface KeyStorage {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
  removeItem(key: string): void;
}

function storage(): KeyStorage | null {
  try {
    return globalThis.localStorage ?? null;
  } catch {
    return null;
  }
}

/** Последние символы ключа, как у сервера: сам ключ панель не показывает. */
function hint(key: string): string {
  return `…${key.slice(-4)}`;
}

export function webKey(id: string, store: KeyStorage | null = storage()): string {
  const provider = WEB_PROVIDERS.find((p) => p.id === id);
  if (!provider) return "";
  try {
    return store?.getItem(provider.storageKey)?.trim() ?? "";
  } catch {
    return "";
  }
}

export function webKeysState(store: KeyStorage | null = storage()): KeysState {
  return {
    storage: store ? { available: true } : { available: false, error: "Хранилище панели недоступно: ключ не сохранить." },
    providers: WEB_PROVIDERS.map((provider): KeyStatus => {
      const key = webKey(provider.id, store);
      return {
        id: provider.id,
        label: provider.label,
        source: key ? "panel" : null,
        hint: key ? hint(key) : null,
        ready: Boolean(key),
        keyOptional: false,
        kind: "model",
        site: provider.site
      };
    })
  };
}

export function saveWebKey(id: string, key: string, store: KeyStorage | null = storage()): KeysState {
  const provider = webProvider(id);
  const value = key.trim();
  if (!value) throw new Error("Ключ пуст.");
  if (/\s/.test(value) || value.length > 512) throw new Error(`Это не похоже на ключ ${provider.label}.`);
  if (!store) throw new Error("Хранилище панели недоступно: ключ не сохранить.");
  store.setItem(provider.storageKey, value);
  return webKeysState(store);
}

export function deleteWebKey(id: string, store: KeyStorage | null = storage()): KeysState {
  const provider = WEB_PROVIDERS.find((p) => p.id === id);
  if (provider) store?.removeItem(provider.storageKey);
  return webKeysState(store);
}

/** Поставщики, для которых сохранён ключ. */
export function webProviders(store: KeyStorage | null = storage()): ProviderInfo[] {
  return WEB_PROVIDERS.filter((provider) => webKey(provider.id, store)).map((provider) => ({
    id: provider.id,
    label: provider.label,
    models: provider.models,
    defaultModel: provider.defaultModel
  }));
}

/** Запрос к поставщику — то же тело, что строит сервер (server.ts, chatBody). */
export function webChatRequest(
  opts: { provider: string; model: string; messages: unknown[]; tools: unknown[] },
  key: string,
  origin: string
): { url: string; init: RequestInit } {
  const provider = webProvider(opts.provider);
  if (!provider.models.includes(opts.model)) {
    throw new Error(`Модель "${opts.model}" не разрешена для ${provider.label}.`);
  }
  const headers: Record<string, string> = { "Content-Type": "application/json", Authorization: `Bearer ${key}` };
  if (provider.id === "openrouter") {
    headers["HTTP-Referer"] = origin;
    headers["X-Title"] = "am.AI";
  }
  return {
    url: provider.url,
    init: {
      method: "POST",
      headers,
      body: JSON.stringify({
        model: opts.model,
        messages: serializeMessages(opts.messages as InternalMessage[], provider.id),
        tools: opts.tools,
        stream: true,
        stream_options: { include_usage: true },
        ...(provider.id === "deepseek"
          ? { thinking: { type: "enabled" }, reasoning_effort: "high" }
          : { tool_choice: "auto" })
      })
    }
  };
}
