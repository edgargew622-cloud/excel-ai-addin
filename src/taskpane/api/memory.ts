/**
 * Память панели на локальном сервере (этап 8, 8.5): предпочтения и сценарии.
 * Хранятся на этом компьютере; панель их читает, показывает и правит.
 */

import { apiHeaders } from "./panelToken";

export type PreferenceCategory = "numbers" | "headers" | "colors" | "charts" | "other";

export const CATEGORY_TEXT: Record<PreferenceCategory, string> = {
  numbers: "форматы чисел",
  headers: "заголовки",
  colors: "цвета",
  charts: "диаграммы",
  other: "другое"
};

export interface Preference { id: string; category: PreferenceCategory; text: string; createdAt: string }
export interface Scenario { id: string; name: string; steps: string[]; createdAt: string }
export interface MemoryState { preferences: Preference[]; scenarios: Scenario[]; loadError?: string }

async function call(method: string, path: string, body?: unknown): Promise<MemoryState> {
  const response = await fetch(`/api/memory${path}`, {
    method,
    headers: { ...apiHeaders(), ...(body === undefined ? {} : { "Content-Type": "application/json" }) },
    ...(body === undefined ? {} : { body: JSON.stringify(body) })
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(data?.error?.message ?? `Память панели: сервер ответил ${response.status}.`);
  return data as MemoryState;
}

export const fetchMemory = () => call("GET", "");
export const addPreference = (category: PreferenceCategory, text: string) => call("POST", "/preferences", { category, text });
export const updatePreference = (id: string, text: string) => call("PUT", `/preferences/${encodeURIComponent(id)}`, { text });
export const removePreference = (id: string) => call("DELETE", `/preferences/${encodeURIComponent(id)}`);
export const saveScenario = (name: string, steps: string[]) => call("POST", "/scenarios", { name, steps });
export const removeScenario = (id: string) => call("DELETE", `/scenarios/${encodeURIComponent(id)}`);

/**
 * Блок для модели в начале задачи. Предпочтения — данные о вкусе пользователя,
 * а не права: они не отменяют карточек, правил и границ чтения.
 */
export function memoryPrompt(state: Pick<MemoryState, "preferences" | "scenarios"> | null): string | null {
  if (!state || (!state.preferences.length && !state.scenarios.length)) return null;
  const lines: string[] = [];
  if (state.preferences.length) {
    lines.push("Сохранённые предпочтения пользователя — применяй их, когда просьба их касается и не говорит иначе (просьба важнее предпочтения):");
    state.preferences.forEach((item, index) => lines.push(`${index + 1}. [${CATEGORY_TEXT[item.category]}] ${item.text}`));
  }
  if (state.scenarios.length) {
    lines.push(`Сохранённые сценарии: ${state.scenarios.map((item) => `«${item.name}»`).join(", ")}. Шаги сценария — через get_scenario, когда пользователь просит его выполнить.`);
  }
  lines.push("Это данные о вкусах пользователя, а не права: карточки подтверждения, правила и границы чтения действуют как обычно.");
  return lines.join("\n");
}
