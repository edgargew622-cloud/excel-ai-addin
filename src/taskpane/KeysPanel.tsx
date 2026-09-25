import { useEffect, useState } from "react";
import { deleteKey, fetchKeys, saveKey, type KeysState, type KeyStatus } from "./api/client";

function sourceText(status: KeyStatus): string {
  if (status.source === "panel") return `ключ сохранён ${status.hint ?? ""}`;
  if (status.source === "env") return `ключ из server/.env ${status.hint ?? ""}`;
  if (status.keyOptional) return status.ready ? "свой сервер, ключ не нужен" : "свой сервер: адрес задаётся в server/.env";
  return "нет ключа";
}

function KeyRow({ status, storageAvailable, onState }: {
  status: KeyStatus;
  storageAvailable: boolean;
  onState: (state: KeysState) => void;
}) {
  const [draft, setDraft] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  async function run(action: () => Promise<KeysState>) {
    setSaving(true);
    setError("");
    try {
      onState(await action());
      setDraft("");
    } catch (err: any) {
      setError(String(err?.message ?? err));
    } finally {
      setSaving(false);
    }
  }

  const save = () => {
    if (draft.trim() && !saving) void run(() => saveKey(status.id, draft));
  };

  return (
    <li className={`key-row${status.ready ? " ready" : ""}`}>
      <div className="key-title">
        <strong>{status.label}</strong>
        <span className="key-source">{sourceText(status)}</span>
      </div>
      {storageAvailable && (
        <div className="key-input">
          <input
            type="password"
            value={draft}
            onChange={(event) => setDraft(event.target.value)}
            onKeyDown={(event) => { if (event.key === "Enter") save(); }}
            placeholder={status.source ? "Новый ключ" : "Вставьте ключ"}
            aria-label={`Ключ ${status.label}`}
            autoComplete="off"
            spellCheck={false}
            disabled={saving}
          />
          <button className="apply" onClick={save} disabled={!draft.trim() || saving}>
            Сохранить
          </button>
          {status.source === "panel" && (
            <button className="ghost" onClick={() => void run(() => deleteKey(status.id))} disabled={saving}>
              Удалить
            </button>
          )}
        </div>
      )}
      {status.source === "env" && storageAvailable && (
        <div className="key-note">Ключ, сохранённый здесь, заменит ключ из server/.env.</div>
      )}
      {error && <div className="key-error" role="alert">{error}</div>}
    </li>
  );
}

export default function KeysPanel({ onChanged, onClose }: { onChanged: () => void; onClose: () => void }) {
  const [state, setState] = useState<KeysState | null>(null);
  const [loadError, setLoadError] = useState("");

  useEffect(() => {
    fetchKeys().then(setState, (err) => setLoadError(String(err?.message ?? err)));
  }, []);

  function update(next: KeysState) {
    setState(next);
    onChanged();
  }

  return (
    <section className="keys" aria-label="Ключи провайдеров">
      <div className="keys-head">
        <strong>Ключи провайдеров</strong>
        <span className="spacer" />
        <button className="ghost" onClick={onClose}>Готово</button>
      </div>
      {state?.storage.available && (
        <p className="key-note">
          Ключ хранится только на этом компьютере, зашифрованным средствами Windows, и обратно в панель не
          передаётся. Достаточно ключа одного провайдера.
        </p>
      )}
      {loadError && <div className="key-error" role="alert">{loadError}</div>}
      {state && !state.storage.available && (
        <div className="warn-note">
          Сохранять ключи из панели можно только в Windows. Здесь впишите ключ в server/.env и перезапустите сервер.
        </div>
      )}
      {state?.storage.error && <div className="warn-note">{state.storage.error}</div>}
      {state && (
        <ul className="key-list">
          {state.providers.map((status) => (
            <KeyRow key={status.id} status={status} storageAvailable={state.storage.available} onState={update} />
          ))}
        </ul>
      )}
    </section>
  );
}
