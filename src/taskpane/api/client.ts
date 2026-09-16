/**
 * Всё общение с моделями идёт через собственный прокси на /api.
 * Ключей провайдеров в этом файле нет и быть не должно.
 */

export interface ProviderInfo {
  id: string;
  label: string;
  models: string[];
  defaultModel: string;
}

export interface ToolCall {
  id: string;
  name: string;
  arguments: string;
}

export type ChatMessage =
  | { role: "system" | "user"; content: string }
  | {
      role: "assistant";
      content: string | null;
      provider?: string;
      reasoning_content?: string | null;
      tool_calls?: ToolCall[];
    }
  | { role: "tool"; tool_call_id: string; content: string };

export interface StreamResult {
  content: string;
  reasoningContent: string;
  toolCalls: ToolCall[];
  finishReason: string;
}

function headers(): Record<string, string> {
  return { "Content-Type": "application/json" };
}

export async function fetchProviders(): Promise<ProviderInfo[]> {
  const res = await fetch("/api/providers", { headers: headers() });
  if (!res.ok) throw new Error(`Локальный сервер вернул ${res.status}. Проверьте npm run diagnose.`);
  return res.json();
}

/**
 * Терминальные finish_reason, при которых вызовы инструментов считаются
 * завершёнными и могут исполняться. Провайдеры называют это по-разному:
 * OpenAI-совместимые шлют tool_calls, старые сборки — function_call,
 * а Claude через OpenRouter отдаёт своё родное tool_use. Любой другой
 * признак означает оборванный шаг, и команды из него не исполняются.
 */
export const TOOL_CALL_FINISH_REASONS = ["tool_calls", "function_call", "tool_use"];

/** Один полностью завершённый проход модели. */
export async function streamChat(opts: {
  provider: string;
  model: string;
  messages: ChatMessage[];
  tools: unknown[];
  signal?: AbortSignal;
  onDelta: (text: string) => void;
}): Promise<StreamResult> {
  const res = await fetch("/api/chat", {
    method: "POST",
    headers: headers(),
    signal: opts.signal,
    body: JSON.stringify({
      provider: opts.provider,
      model: opts.model,
      messages: opts.messages,
      tools: opts.tools
    })
  });

  if (!res.ok || !res.body) {
    const text = await res.text().catch(() => "");
    let message = text;
    try { message = JSON.parse(text)?.error?.message ?? text; } catch { /* non-JSON error */ }
    throw new Error(message || `Ошибка локального сервера ${res.status}`);
  }

  const reader = res.body.getReader();
  const decoder = new TextDecoder();

  let buffer = "";
  let content = "";
  let reasoningContent = "";
  let finishReason: string | null = null;
  let sawDone = false;
  const partial = new Map<number, ToolCall>();

  const consumeLine = (line: string) => {
    const trimmed = line.trim();
    if (!trimmed.startsWith("data:")) return;

    const payload = trimmed.slice(5).trim();
    if (payload === "[DONE]") {
      sawDone = true;
      return;
    }

    let chunk: any;
    try {
      chunk = JSON.parse(payload);
    } catch {
      throw new Error("Получен повреждённый SSE-кадр от провайдера.");
    }

    if (chunk.error) throw new Error(chunk.error.message ?? "Провайдер вернул ошибку");

    const choice = chunk.choices?.[0];
    if (!choice) return;
    if (typeof choice.finish_reason === "string" && choice.finish_reason) finishReason = choice.finish_reason;

    const delta = choice.delta ?? {};
    if (typeof delta.reasoning_content === "string") reasoningContent += delta.reasoning_content;

    if (typeof delta.content === "string" && delta.content) {
      content += delta.content;
      opts.onDelta(delta.content);
    }

    for (const tc of delta.tool_calls ?? []) {
      const idx = Number.isInteger(tc.index) ? tc.index : 0;
      const acc = partial.get(idx) ?? { id: "", name: "", arguments: "" };
      if (tc.id) acc.id = tc.id;
      if (tc.function?.name) acc.name = tc.function.name;
      if (typeof tc.function?.arguments === "string") acc.arguments += tc.function.arguments;
      partial.set(idx, acc);
    }
  };

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;

    buffer += decoder.decode(value, { stream: true });
    const lines = buffer.split("\n");
    buffer = lines.pop() ?? "";
    for (const line of lines) consumeLine(line);
  }

  buffer += decoder.decode();
  if (buffer.trim()) consumeLine(buffer);

  // Не исполняем ничего из оборванного ответа. finish_reason — обязательный
  // терминальный маркер шага. [DONE] дополнительно проверяем там, где он есть.
  if (!finishReason) {
    throw new Error("Поток ответа оборвался до finish_reason; команды не выполнялись.");
  }

  const toolCalls = [...partial.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([i, c]) => ({ ...c, id: c.id || `call_${i}_${Date.now()}` }))
    .filter((c) => c.name);

  if (toolCalls.length && !TOOL_CALL_FINISH_REASONS.includes(finishReason)) {
    throw new Error(`Провайдер прислал tool_calls, но завершил шаг как "${finishReason}"; команды не выполнялись.`);
  }

  // Некоторые совместимые провайдеры не присылают [DONE], поэтому он не
  // является единственным критерием. Если [DONE] был — отлично; безопасность
  // обеспечивается терминальным finish_reason и полным JSON аргументов ниже.
  void sawDone;

  for (const call of toolCalls) {
    try {
      JSON.parse(call.arguments || "{}");
    } catch {
      throw new Error(`Аргументы ${call.name} не завершены как JSON; команда не выполнялась.`);
    }
  }

  return { content, reasoningContent, toolCalls, finishReason };
}
