/**
 * Беседа в файл (этап 10, 10.1). В панели Windows файл пишет локальный
 * сервер в «Документы\am.AI\Беседы» и называет путь. В веб-панели (Mac,
 * Excel в браузере) сервера нет — там беседа копируется текстом.
 */

import { apiHeaders } from "./panelToken";

export async function exportConversationFile(name: string, text: string): Promise<string> {
  const res = await fetch("/api/conversations/export", {
    method: "POST",
    headers: apiHeaders(),
    body: JSON.stringify({ name, text })
  });
  const data = await res.json().catch(() => null);
  if (!res.ok) throw new Error(data?.error?.message ?? `Локальный сервер вернул ${res.status}.`);
  return String(data.path);
}

/** Копирование текста: буфер обмена, а если панель его не дала — выделение для ⌘C / Ctrl+C. */
export async function copyText(text: string, fallbackTarget?: HTMLTextAreaElement | null): Promise<"copied" | "selected"> {
  try {
    await navigator.clipboard.writeText(text);
    return "copied";
  } catch {
    if (fallbackTarget) {
      fallbackTarget.value = text;
      fallbackTarget.focus();
      fallbackTarget.select();
      try {
        if (document.execCommand("copy")) return "copied";
      } catch {
        /* старый способ тоже недоступен */
      }
    }
    return "selected";
  }
}
