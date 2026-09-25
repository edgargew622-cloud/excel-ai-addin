/**
 * Ключи провайдеров из панели: посмотреть состояние, сохранить, удалить.
 *
 * Ключ приходит только внутрь и никогда не уходит обратно: в ответах есть
 * лишь источник и последние символы. Доступ ограничен общими проверками
 * сервера (петлевой адрес, свой Host и Origin), поэтому сторонняя страница в
 * браузере не может ни подменить ключ, ни прочитать список.
 */

import type { Express, Response } from "express";
import { KeyError, keyHint, type KeyStore } from "./keyStore.js";
import { enabledProviders, getProvider, keySource, providerKey, providerReady } from "./providers.js";

export interface KeyStatus {
  id: string;
  label: string;
  source: "panel" | "env" | null;
  hint: string | null;
  ready: boolean;
  /** Своему серверу ключ может быть не нужен — нужен адрес в server/.env. */
  keyOptional: boolean;
}

export function keyStatuses(): KeyStatus[] {
  return enabledProviders().map((p) => {
    const key = providerKey(p);
    return {
      id: p.id,
      label: p.label,
      source: keySource(p),
      hint: key ? keyHint(key) : null,
      ready: providerReady(p),
      keyOptional: Boolean(p.keyOptional)
    };
  });
}

function fail(res: Response, error: unknown): void {
  const known = error instanceof KeyError;
  // Текст ошибки шифрования не содержит ключа, но подробности всё равно не
  // нужны панели: они остаются в журнале сервера.
  if (!known) console.error(`Ключи: ${String((error as Error)?.message ?? error)}`);
  res.status(known ? 400 : 500).json({
    error: { message: known ? (error as Error).message : "Не удалось сохранить ключ. Подробности — в журнале сервера." }
  });
}

export function registerKeyRoutes(app: Express, store: KeyStore): void {
  const snapshot = () => ({
    storage: { available: store.storageAvailable, ...(store.loadError ? { error: store.loadError } : {}) },
    providers: keyStatuses()
  });

  app.get("/api/keys", (_req, res) => res.json(snapshot()));

  app.put("/api/keys/:id", async (req, res) => {
    const provider = getProvider(req.params.id);
    if (!provider) return res.status(404).json({ error: { message: `Провайдер "${req.params.id}" неизвестен или отключён.` } });
    try {
      await store.set(provider.id, req.body?.key);
      console.log(`Ключ ${provider.label} сохранён в панели.`);
      res.json(snapshot());
    } catch (error) {
      fail(res, error);
    }
  });

  app.delete("/api/keys/:id", async (req, res) => {
    const provider = getProvider(req.params.id);
    if (!provider) return res.status(404).json({ error: { message: `Провайдер "${req.params.id}" неизвестен или отключён.` } });
    try {
      const removed = await store.remove(provider.id);
      if (removed) console.log(`Ключ ${provider.label} удалён из панели.`);
      res.json(snapshot());
    } catch (error) {
      fail(res, error);
    }
  });
}
