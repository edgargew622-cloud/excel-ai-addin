/**
 * Инструменты памяти для модели (этап 8, 8.5). Книгу они не трогают: пишут
 * в память панели на этом компьютере. Сохранение идёт через карточку
 * (`destructive`), недоступно в «Только анализ» (`sideEffect: "memory"`) и
 * только по прямой просьбе пользователя — это проверяет цикл агента.
 */

import { addPreference, fetchMemory, saveScenario, type PreferenceCategory } from "../taskpane/api/memory";
import { ToolError } from "./excelTools";

export async function remember_preference(a: { category: PreferenceCategory; text: string }) {
  const state = await addPreference(a.category, a.text);
  const saved = state.preferences.find((item) => item.text.toLowerCase() === a.text.trim().replace(/\s+/g, " ").toLowerCase());
  if (!saved) throw new ToolError("Сервер не подтвердил сохранение предпочтения. Проверьте окно «Память».");
  return {
    ok: true,
    saved: saved.text,
    preferences: state.preferences.length,
    note: "Сохранено на этом компьютере. Со следующей задачи оно придёт тебе в начале. Изменить или удалить — окно «Память» в панели."
  };
}

export async function save_scenario(a: { name: string; steps: string[] }) {
  const state = await saveScenario(a.name, a.steps);
  const saved = state.scenarios.find((item) => item.name.toLowerCase() === a.name.trim().replace(/\s+/g, " ").toLowerCase());
  if (!saved) throw new ToolError("Сервер не подтвердил сохранение сценария. Проверьте окно «Память».");
  return {
    ok: true,
    scenario: saved.name,
    steps: saved.steps,
    note: "Сценарий сохранён на этом компьютере. Запуск — кнопкой в окне «Память» или просьбой «выполни сценарий …». Каждое изменение внутри по-прежнему через карточку."
  };
}

export async function get_scenario(a: { name: string }) {
  const state = await fetchMemory();
  const wanted = a.name.trim().toLowerCase();
  const found = state.scenarios.find((item) => item.name.toLowerCase() === wanted)
    ?? state.scenarios.find((item) => item.name.toLowerCase().includes(wanted));
  if (!found) {
    throw new ToolError(`Сценария «${a.name}» нет. Сохранены: ${state.scenarios.map((item) => `«${item.name}»`).join(", ") || "ни одного"}.`);
  }
  return {
    name: found.name,
    steps: found.steps,
    note: "Выполняй шаги по порядку как обычные просьбы: каждое изменение — через свою карточку. Если шаг не подходит к этой книге, остановись и спроси."
  };
}
