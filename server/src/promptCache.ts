/**
 * Кэш запроса для Claude (01.10.2026).
 *
 * Каждый шаг агента несёт ~35–45 тыс. токенов: описания инструментов, правила
 * и беседу. DeepSeek и OpenAI сами берут меньше за повторяющееся начало
 * запроса, а Claude — только если начало помечено cache_control. Без отметок
 * беседа пользователя на Sonnet 5.5 стоила $1,30 вместо ~$0,25.
 *
 * Отметки (Anthropic кэширует всё начало запроса до отметки, в порядке
 * инструменты → системные сообщения → беседа):
 * 1. на первом системном сообщении — правила агента: с ними в кэш попадают и
 *    описания инструментов; одинаковы во всех задачах;
 * 2. на последнем сообщении — так беседа кэшируется по ходу, и каждый шаг
 *    платит полную цену только за новое.
 * Кэш живёт 5 минут с последнего обращения. OpenRouter передаёт отметки в
 * Claude как есть; у остальных моделей сообщения не меняются.
 */

const EPHEMERAL = { type: "ephemeral" } as const;

export function cachesByMarkers(providerId: string, model: string): boolean {
  return (providerId === "openrouter" && /^anthropic\//.test(model)) || providerId === "anthropic";
}

function marked(message: Record<string, unknown>): Record<string, unknown> {
  const content = message.content;
  if (typeof content === "string" && content.length) {
    return { ...message, content: [{ type: "text", text: content, cache_control: EPHEMERAL }] };
  }
  if (Array.isArray(content) && content.length) {
    const parts = content.map((part) => ({ ...part }));
    parts[parts.length - 1] = { ...parts[parts.length - 1], cache_control: EPHEMERAL };
    return { ...message, content: parts };
  }
  return message;
}

export function withPromptCache(messages: readonly unknown[], providerId: string, model: string): unknown[] {
  if (!cachesByMarkers(providerId, model)) return [...messages];
  const out = messages.map((message) => ({ ...(message as Record<string, unknown>) }));
  const firstSystem = out.findIndex((message) => message.role === "system");
  if (firstSystem >= 0) out[firstSystem] = marked(out[firstSystem]);
  // Последнее сообщение с текстом: обычно вопрос пользователя или ответ инструмента.
  for (let i = out.length - 1; i > firstSystem; i--) {
    const content = out[i].content;
    if ((typeof content === "string" && content.length) || (Array.isArray(content) && content.length)) {
      out[i] = marked(out[i]);
      break;
    }
  }
  return out;
}
