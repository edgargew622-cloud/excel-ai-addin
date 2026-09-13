export type InternalToolCall = {
  id: string;
  name: string;
  arguments: string;
};

export type InternalMessage =
  | { role: "system" | "user"; content: string }
  | {
      role: "assistant";
      content: string | null;
      provider?: string;
      reasoning_content?: string | null;
      tool_calls?: InternalToolCall[];
    }
  | { role: "tool"; tool_call_id: string; content: string };

/**
 * Панель хранит единый внутренний формат tool call: {id,name,arguments}.
 * OpenAI-compatible Chat Completions ожидает wire-формат
 * {id,type:"function",function:{name,arguments}}. Сериализация живёт только
 * на границе с провайдером, чтобы внутренний agent loop не зависел от API.
 */
export function serializeMessages(messages: InternalMessage[], providerId: string): unknown[] {
  return messages.map((message) => {
    if (message.role !== "assistant") return message;

    const out: Record<string, unknown> = {
      role: "assistant",
      content: message.content ?? ""
    };

    if (message.tool_calls?.length) {
      out.tool_calls = message.tool_calls.map((call) => ({
        id: call.id,
        type: "function",
        function: {
          name: call.name,
          arguments: call.arguments
        }
      }));
    }

    // reasoning_content — расширение DeepSeek. Не отправляем его OpenAI,
    // xAI, Moonshot и др. При DeepSeek сохраняем даже пустую строку для
    // собственных assistant-ходов этого провайдера, чтобы tool-chain replay
    // соответствовал требованиям thinking mode.
    if (providerId === "deepseek" && message.provider === "deepseek") {
      out.reasoning_content = typeof message.reasoning_content === "string" ? message.reasoning_content : "";
    }

    return out;
  });
}
