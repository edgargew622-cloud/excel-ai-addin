/**
 * Номер беседы внутри самой книги (28.09.2026, найдено пользователем: после
 * первого сохранения книги в OneDrive чат исчезал).
 *
 * Адрес книги меняется при первом сохранении и «Сохранить как», а при
 * сохранении в облако Excel ещё и перезапускает панель — привязка беседы к
 * адресу её теряла. Поэтому в настройки документа (Office Settings — они
 * сохраняются в файле вместе с книгой) пишется случайный номер беседы.
 * Сама переписка по-прежнему хранится только на этом компьютере; в файле —
 * лишь номер, по которому её найти. Чужая книга с чужим номером ничего не
 * откроет: переписки с таким номером на этом компьютере нет.
 */

const SETTING = "amai.conversationId";
const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function settings(): Office.Settings | null {
  try {
    return (globalThis as any).Office?.context?.document?.settings ?? null;
  } catch {
    return null;
  }
}

export function documentConversationId(): string | null {
  const value = settings()?.get(SETTING);
  return typeof value === "string" && UUID.test(value) ? value.toLowerCase() : null;
}

/** Номер беседы книги: прежний или новый, записанный в настройки документа. */
export function ensureDocumentConversationId(): Promise<string | null> {
  const existing = documentConversationId();
  if (existing) return Promise.resolve(existing);
  const store = settings();
  if (!store) return Promise.resolve(null);
  const id = (globalThis.crypto?.randomUUID?.() ?? "").toLowerCase();
  if (!UUID.test(id)) return Promise.resolve(null);
  return new Promise((resolve) => {
    try {
      store.set(SETTING, id);
      store.saveAsync((result) => resolve(String(result.status).toLowerCase() === "succeeded" ? id : null));
    } catch {
      resolve(null);
    }
  });
}

export const documentConversationKey = (id: string) => `doc:${id}`;
