/**
 * Веб-режим панели (Mac, Excel в браузере): OpenRouter напрямую из панели.
 *
 * Локального сервера нет, поэтому ключ OpenRouter хранится в хранилище
 * панели (localStorage источника панели) на этом компьютере — это слабее,
 * чем шифрование Windows в обычном режиме, и панель прямо говорит об этом.
 * OpenRouter разрешает запросы из браузера (CORS), остальные провайдеры —
 * как правило, нет, поэтому здесь только он. Список моделей и формат
 * сообщений — те же, что у сервера: модули общие.
 */

import { PROVIDERS } from "../../../server/src/providers";
import { serializeMessages, type InternalMessage } from "../../../server/src/protocol";
import type { KeysState, ProviderInfo } from "./client";

export const OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions";
const KEY_STORAGE = "amai.openrouter-key";

const openrouter = PROVIDERS.find((p) => p.id === "openrouter")!;

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

export function webKey(store: KeyStorage | null = storage()): string {
  try {
    return store?.getItem(KEY_STORAGE)?.trim() ?? "";
  } catch {
    return "";
  }
}

export function webKeysState(store: KeyStorage | null = storage()): KeysState {
  const key = webKey(store);
  return {
    storage: store ? { available: true } : { available: false, error: "Хранилище панели недоступно: ключ не сохранить." },
    providers: [{
      id: openrouter.id,
      label: openrouter.label,
      source: key ? "panel" : null,
      hint: key ? hint(key) : null,
      ready: Boolean(key),
      keyOptional: false,
      kind: "model"
    }]
  };
}

export function saveWebKey(id: string, key: string, store: KeyStorage | null = storage()): KeysState {
  if (id !== openrouter.id) throw new Error("В этой панели работает только OpenRouter.");
  const value = key.trim();
  if (!value) throw new Error("Ключ пуст.");
  if (/\s/.test(value) || value.length > 512) throw new Error("Это не похоже на ключ OpenRouter.");
  if (!store) throw new Error("Хранилище панели недоступно: ключ не сохранить.");
  store.setItem(KEY_STORAGE, value);
  return webKeysState(store);
}

export function deleteWebKey(id: string, store: KeyStorage | null = storage()): KeysState {
  if (id === openrouter.id) store?.removeItem(KEY_STORAGE);
  return webKeysState(store);
}

export function webProviders(store: KeyStorage | null = storage()): ProviderInfo[] {
  if (!webKey(store)) return [];
  return [{ id: openrouter.id, label: openrouter.label, models: openrouter.models, defaultModel: openrouter.defaultModel }];
}

/** Запрос к OpenRouter — то же тело, что строит сервер для этого провайдера. */
export function webChatRequest(opts: { model: string; messages: unknown[]; tools: unknown[] }, key: string, origin: string): { url: string; init: RequestInit } {
  if (!openrouter.models.includes(opts.model)) {
    throw new Error(`Модель "${opts.model}" не разрешена для OpenRouter.`);
  }
  return {
    url: OPENROUTER_URL,
    init: {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${key}`,
        "HTTP-Referer": origin,
        "X-Title": "am.AI"
      },
      body: JSON.stringify({
        model: opts.model,
        messages: serializeMessages(opts.messages as InternalMessage[], openrouter.id),
        tools: opts.tools,
        stream: true,
        stream_options: { include_usage: true },
        tool_choice: "auto"
      })
    }
  };
}
