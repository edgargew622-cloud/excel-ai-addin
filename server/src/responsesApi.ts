/**
 * Мост к /v1/responses.
 *
 * Панель и внутренний цикл агента говорят на формате Chat Completions.
 * Некоторые модели — например `gpt-6-astra` — доступны с функциями только
 * через /v1/responses, поэтому запрос и поток переводятся здесь, на границе
 * с провайдером. Всё остальное приложение об этом интерфейсе не знает.
 */

import type { InternalMessage } from "./protocol.js";

export interface ChatTool {
  type: "function";
  function: { name: string; description?: string; parameters?: unknown };
}

/** Сообщения Chat Completions → элементы input для /v1/responses.
 * Вызовы инструментов и их результаты там отдельные элементы, а не поля
 * сообщения, и связываются между собой по call_id. */
export function toResponsesInput(messages: InternalMessage[]): unknown[] {
  const input: unknown[] = [];
  for (const message of messages) {
    if (message.role === "tool") {
      input.push({ type: "function_call_output", call_id: message.tool_call_id, output: message.content });
      continue;
    }
    if (message.role === "assistant") {
      if (message.content) input.push({ role: "assistant", content: message.content });
      for (const call of message.tool_calls ?? []) {
        input.push({ type: "function_call", call_id: call.id, name: call.name, arguments: call.arguments });
      }
      continue;
    }
    input.push({ role: message.role, content: message.content });
  }
  return input;
}

/** Инструменты в /v1/responses описываются плоско, без вложенного function. */
export function toResponsesTools(tools: ChatTool[]): unknown[] {
  return tools.map((tool) => ({
    type: "function",
    name: tool.function.name,
    description: tool.function.description,
    parameters: tool.function.parameters
  }));
}

export function buildResponsesBody(model: string, messages: InternalMessage[], tools: ChatTool[]): Record<string, unknown> {
  return {
    model,
    input: toResponsesInput(messages),
    tools: toResponsesTools(tools),
    tool_choice: "auto",
    stream: true
  };
}

/** Кусок SSE в формате Chat Completions — именно его ждёт панель. */
function chunk(payload: Record<string, unknown>): string {
  return `data: ${JSON.stringify(payload)}\n\n`;
}

const textDelta = (text: string) => chunk({ choices: [{ index: 0, delta: { content: text } }] });

const toolCallStart = (index: number, id: string, name: string) =>
  chunk({ choices: [{ index: 0, delta: { tool_calls: [{ index, id, type: "function", function: { name, arguments: "" } }] } }] });

const toolCallArgs = (index: number, args: string) =>
  chunk({ choices: [{ index: 0, delta: { tool_calls: [{ index, function: { arguments: args } }] } }] });

const finish = (reason: string) => chunk({ choices: [{ index: 0, delta: {}, finish_reason: reason }] });

/**
 * Переводит одно событие /v1/responses в ноль или несколько кусков
 * Chat Completions. Состояние нужно, чтобы нумеровать вызовы инструментов
 * подряд: в /v1/responses у них свои output_index вперемешку с сообщениями.
 */
export class ResponsesTranslator {
  private readonly indexByItem = new Map<string, number>();
  private toolCount = 0;
  private finished = false;

  /** Вернулись ли вызовы инструментов: от этого зависит finish_reason. */
  get sawToolCalls(): boolean {
    return this.toolCount > 0;
  }

  translate(event: Record<string, any>): string[] {
    const type = String(event?.type ?? "");

    if (type === "response.output_text.delta" && typeof event.delta === "string" && event.delta) {
      return [textDelta(event.delta)];
    }

    // Сводка размышлений показывается отдельно от ответа и в историю не идёт.
    if (type === "response.reasoning_summary_text.delta" && typeof event.delta === "string" && event.delta) {
      return [chunk({ choices: [{ index: 0, delta: { reasoning_content: event.delta } }] })];
    }

    if (type === "response.output_item.added" && event.item?.type === "function_call") {
      const itemId = String(event.item.id ?? event.item.call_id ?? "");
      const index = this.toolCount++;
      this.indexByItem.set(itemId, index);
      return [toolCallStart(index, String(event.item.call_id ?? itemId), String(event.item.name ?? ""))];
    }

    if (type === "response.function_call_arguments.delta" && typeof event.delta === "string") {
      const index = this.indexByItem.get(String(event.item_id ?? "")) ?? 0;
      return event.delta ? [toolCallArgs(index, event.delta)] : [];
    }

    if (type === "response.completed" || type === "response.incomplete") {
      if (this.finished) return [];
      this.finished = true;
      // incomplete означает обрыв по лимиту: терминальным «stop» его выдавать
      // нельзя, иначе оборванный шаг выглядел бы завершённым.
      const reason = type === "response.incomplete" ? "length" : this.sawToolCalls ? "tool_calls" : "stop";
      return [finish(reason), "data: [DONE]\n\n"];
    }

    if (type === "response.failed" || type === "error") {
      const message = event.response?.error?.message ?? event.message ?? "Провайдер прервал ответ";
      return [chunk({ error: { message: String(message) } })];
    }

    return [];
  }

  /** Поток кончился без терминального события: шаг не завершён, и молчать
   * об этом нельзя — панель обязана отвергнуть оборванный ответ. */
  finalizeIfUnfinished(): string[] {
    if (this.finished) return [];
    this.finished = true;
    return [chunk({ error: { message: "Поток /v1/responses оборвался до завершения ответа." } })];
  }
}

/** Разбирает SSE-строки провайдера и отдаёт готовые куски для панели. */
export function translateResponsesChunk(translator: ResponsesTranslator, line: string): string[] {
  const trimmed = line.trim();
  if (!trimmed.startsWith("data:")) return [];
  const payload = trimmed.slice(5).trim();
  if (!payload || payload === "[DONE]") return [];
  try {
    return translator.translate(JSON.parse(payload));
  } catch {
    return [];
  }
}
