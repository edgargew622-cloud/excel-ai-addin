/**
 * Память панели (этап 8, 8.5): предпочтения пользователя и сохранённые сценарии.
 *
 * Хранится на этом компьютере, в файле рядом с ключами — в папке надстройки,
 * закрытой для других пользователей Windows (8.0.1). Это не секрет, поэтому
 * без шифрования, но и никуда не отправляется: модель получает предпочтения
 * коротким блоком в начале задачи через панель.
 *
 * Записи создаются только из панели: после карточки подтверждения или в
 * окне «Память». Сервер сам ограничивает число и длину записей — панели и
 * модели он не доверяет.
 */

import type { Express } from "express";
import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { randomUUID } from "node:crypto";
import { replaceFile } from "./keyStore.js";

export const PREFERENCE_CATEGORIES = ["numbers", "headers", "colors", "charts", "other"] as const;
export type PreferenceCategory = (typeof PREFERENCE_CATEGORIES)[number];

export interface Preference {
  id: string;
  category: PreferenceCategory;
  text: string;
  createdAt: string;
}

export interface Scenario {
  id: string;
  name: string;
  steps: string[];
  createdAt: string;
}

export interface MemoryState {
  preferences: Preference[];
  scenarios: Scenario[];
}

export const MEMORY_LIMITS = { preferences: 30, preferenceText: 200, scenarios: 30, scenarioName: 60, steps: 20, stepText: 500 };

export class MemoryError extends Error {}

const clean = (value: unknown) => String(value ?? "").replace(/[\u0000-\u001f\u007f]/g, " ").replace(/\s+/g, " ").trim();

export function validatePreference(raw: { category?: unknown; text?: unknown }): { category: PreferenceCategory; text: string } {
  const category = raw.category as PreferenceCategory;
  if (!PREFERENCE_CATEGORIES.includes(category)) throw new MemoryError(`Категория — одна из: ${PREFERENCE_CATEGORIES.join(", ")}.`);
  const text = clean(raw.text);
  if (!text) throw new MemoryError("Пустое предпочтение не сохраняется.");
  if (text.length > MEMORY_LIMITS.preferenceText) throw new MemoryError(`Предпочтение длиннее ${MEMORY_LIMITS.preferenceText} знаков — сформулируйте короче.`);
  return { category, text };
}

export function validateScenario(raw: { name?: unknown; steps?: unknown }): { name: string; steps: string[] } {
  const name = clean(raw.name);
  if (!name) throw new MemoryError("У сценария нужно название.");
  if (name.length > MEMORY_LIMITS.scenarioName) throw new MemoryError(`Название сценария длиннее ${MEMORY_LIMITS.scenarioName} знаков.`);
  if (!Array.isArray(raw.steps) || !raw.steps.length) throw new MemoryError("В сценарии нужен хотя бы один шаг.");
  if (raw.steps.length > MEMORY_LIMITS.steps) throw new MemoryError(`Шагов больше ${MEMORY_LIMITS.steps}: разбейте сценарий.`);
  const steps = raw.steps.map(clean);
  if (steps.some((step) => !step)) throw new MemoryError("Пустой шаг в сценарии.");
  if (steps.some((step) => step.length > MEMORY_LIMITS.stepText)) throw new MemoryError(`Шаг длиннее ${MEMORY_LIMITS.stepText} знаков.`);
  return { name, steps };
}

export class MemoryStore {
  private state: MemoryState = { preferences: [], scenarios: [] };
  loadError: string | null = null;

  constructor(private readonly file: string, private readonly now: () => Date = () => new Date()) {}

  load(): void {
    if (!existsSync(this.file)) return;
    try {
      const parsed = JSON.parse(readFileSync(this.file, "utf8")) as Partial<MemoryState>;
      this.state = {
        preferences: Array.isArray(parsed.preferences) ? parsed.preferences.filter((item) => item && typeof item.text === "string") : [],
        scenarios: Array.isArray(parsed.scenarios) ? parsed.scenarios.filter((item) => item && Array.isArray(item.steps)) : []
      };
    } catch (error: any) {
      // Испорченный файл не роняет сервер: память пуста, файл остаётся для разбора.
      this.loadError = `Память панели не прочиталась (${error?.message ?? error}): начинаем с пустой, файл ${this.file} не тронут.`;
    }
  }

  snapshot(): MemoryState {
    return JSON.parse(JSON.stringify(this.state));
  }

  private async save(): Promise<void> {
    if (this.loadError) throw new MemoryError("Файл памяти испорчен — чтобы не затереть его, изменения не сохраняются. Удалите или исправьте его.");
    const temp = `${this.file}.tmp`;
    writeFileSync(temp, JSON.stringify(this.state, null, 2), "utf8");
    await replaceFile(temp, this.file);
  }

  async addPreference(raw: { category?: unknown; text?: unknown }): Promise<Preference> {
    const { category, text } = validatePreference(raw);
    const existing = this.state.preferences.find((item) => item.text.toLowerCase() === text.toLowerCase());
    if (existing) return existing;
    if (this.state.preferences.length >= MEMORY_LIMITS.preferences) throw new MemoryError(`Предпочтений уже ${MEMORY_LIMITS.preferences}: удалите лишние в окне «Память».`);
    const item: Preference = { id: randomUUID(), category, text, createdAt: this.now().toISOString() };
    this.state.preferences.push(item);
    await this.save();
    return item;
  }

  async updatePreference(id: string, raw: { category?: unknown; text?: unknown }): Promise<Preference> {
    const item = this.state.preferences.find((entry) => entry.id === id);
    if (!item) throw new MemoryError("Такого предпочтения нет.");
    const { category, text } = validatePreference({ category: raw.category ?? item.category, text: raw.text });
    item.category = category;
    item.text = text;
    await this.save();
    return item;
  }

  async removePreference(id: string): Promise<void> {
    const before = this.state.preferences.length;
    this.state.preferences = this.state.preferences.filter((item) => item.id !== id);
    if (this.state.preferences.length === before) throw new MemoryError("Такого предпочтения нет.");
    await this.save();
  }

  async saveScenario(raw: { name?: unknown; steps?: unknown }): Promise<Scenario> {
    const { name, steps } = validateScenario(raw);
    const existing = this.state.scenarios.find((item) => item.name.toLowerCase() === name.toLowerCase());
    if (existing) {
      existing.steps = steps;
      await this.save();
      return existing;
    }
    if (this.state.scenarios.length >= MEMORY_LIMITS.scenarios) throw new MemoryError(`Сценариев уже ${MEMORY_LIMITS.scenarios}: удалите лишние в окне «Память».`);
    const item: Scenario = { id: randomUUID(), name, steps, createdAt: this.now().toISOString() };
    this.state.scenarios.push(item);
    await this.save();
    return item;
  }

  async removeScenario(id: string): Promise<void> {
    const before = this.state.scenarios.length;
    this.state.scenarios = this.state.scenarios.filter((item) => item.id !== id);
    if (this.state.scenarios.length === before) throw new MemoryError("Такого сценария нет.");
    await this.save();
  }
}

/** Маршруты /api/memory — за токеном панели, как все /api (8.0.1). */
export function registerMemoryRoutes(app: Express, store: MemoryStore): void {
  const send = (res: any, work: () => Promise<unknown>) => {
    work()
      .then(() => res.json({ ...store.snapshot(), limits: MEMORY_LIMITS }))
      .catch((error: any) => res.status(error instanceof MemoryError ? 400 : 500).json({ error: { message: String(error?.message ?? error) } }));
  };
  app.get("/api/memory", (_req, res) => res.json({ ...store.snapshot(), limits: MEMORY_LIMITS, ...(store.loadError ? { loadError: store.loadError } : {}) }));
  app.post("/api/memory/preferences", (req, res) => send(res, () => store.addPreference(req.body ?? {})));
  app.put("/api/memory/preferences/:id", (req, res) => send(res, () => store.updatePreference(req.params.id, req.body ?? {})));
  app.delete("/api/memory/preferences/:id", (req, res) => send(res, () => store.removePreference(req.params.id)));
  app.post("/api/memory/scenarios", (req, res) => send(res, () => store.saveScenario(req.body ?? {})));
  app.delete("/api/memory/scenarios/:id", (req, res) => send(res, () => store.removeScenario(req.params.id)));
}
