import { useEffect, useState } from "react";
import {
  addPreference,
  CATEGORY_TEXT,
  fetchMemory,
  removePreference,
  removeScenario,
  updatePreference,
  type MemoryState,
  type PreferenceCategory,
  type Scenario
} from "./api/memory";

/**
 * Окно «Память» (этап 8, 8.5): всё, что панель помнит о пользователе, видно
 * и правится здесь. Ничего не запоминается молча — модель может только
 * предложить сохранить, и это проходит через карточку.
 */
export default function MemoryPanel({ onClose, onRun }: { onClose: () => void; onRun: (scenario: Scenario) => void }) {
  const [state, setState] = useState<MemoryState | null>(null);
  const [error, setError] = useState("");
  const [draft, setDraft] = useState("");
  const [category, setCategory] = useState<PreferenceCategory>("other");
  const [editing, setEditing] = useState<{ id: string; text: string } | null>(null);

  const apply = (work: Promise<MemoryState>) => work.then((next) => { setState(next); setError(""); }, (err) => setError(String(err?.message ?? err)));

  useEffect(() => { void apply(fetchMemory()); }, []);

  return (
    <section className="keys" aria-label="Память панели">
      <div className="keys-head">
        <strong>Память</strong>
        <span className="spacer" />
        <button className="ghost" onClick={onClose}>Готово</button>
      </div>
      <p className="key-note">
        Хранится только на этом компьютере. Предпочтения модель получает в начале каждой задачи; просьба важнее предпочтения.
        Агент сам ничего не запоминает — только предлагает, и вы подтверждаете карточкой.
      </p>
      {error && <div className="key-error" role="alert">{error}</div>}
      {state?.loadError && <div className="warn-note">{state.loadError}</div>}

      <strong>Предпочтения</strong>
      {state && !state.preferences.length && <p className="key-note">Пока пусто. Скажите агенту, например: «запомни: суммы всегда с разделителем тысяч».</p>}
      <ul className="key-list">
        {state?.preferences.map((item) => (
          <li key={item.id}>
            {editing?.id === item.id ? (
              <>
                <input value={editing.text} onChange={(event) => setEditing({ id: item.id, text: event.target.value })} maxLength={200} />
                <button onClick={() => { void apply(updatePreference(item.id, editing.text)); setEditing(null); }}>Сохранить</button>
                <button className="ghost" onClick={() => setEditing(null)}>Отмена</button>
              </>
            ) : (
              <>
                <span>[{CATEGORY_TEXT[item.category]}] {item.text}</span>
                <span className="spacer" />
                <button className="ghost" onClick={() => setEditing({ id: item.id, text: item.text })}>Изменить</button>
                <button className="ghost" onClick={() => void apply(removePreference(item.id))}>Удалить</button>
              </>
            )}
          </li>
        ))}
      </ul>
      <div className="row">
        <select value={category} onChange={(event) => setCategory(event.target.value as PreferenceCategory)}>
          {(Object.keys(CATEGORY_TEXT) as PreferenceCategory[]).map((key) => <option key={key} value={key}>{CATEGORY_TEXT[key]}</option>)}
        </select>
        <input value={draft} onChange={(event) => setDraft(event.target.value)} placeholder="Новое предпочтение" maxLength={200} />
        <button disabled={!draft.trim()} onClick={() => { void apply(addPreference(category, draft)); setDraft(""); }}>Добавить</button>
      </div>

      <strong>Сценарии</strong>
      {state && !state.scenarios.length && <p className="key-note">Пока пусто. После задачи скажите агенту: «сохрани это как сценарий „Месячный отчёт“».</p>}
      <ul className="key-list">
        {state?.scenarios.map((item) => (
          <li key={item.id}>
            <details>
              <summary>{item.name} — шагов: {item.steps.length}</summary>
              <ol>{item.steps.map((step, index) => <li key={index}>{step}</li>)}</ol>
            </details>
            <span className="spacer" />
            <button onClick={() => onRun(item)}>Запустить</button>
            <button className="ghost" onClick={() => void apply(removeScenario(item.id))}>Удалить</button>
          </li>
        ))}
      </ul>
    </section>
  );
}
