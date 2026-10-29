/**
 * Мост к Claude (Anthropic Messages API, 28.10.2026).
 *
 * Панель и цикл агента говорят на формате Chat Completions. Claude отвечает
 * через Messages API: запрос строится здесь официальным SDK, а его поток
 * переводится в куски Chat Completions — остальное приложение о другом
 * интерфейсе не знает (так же устроен мост к /v1/responses).
 *
 * Особенности Claude, которые здесь учтены:
 * - системные сообщения — отдельное поле system, а не сообщения беседы;
 * - вызов инструмента — блок tool_use, результат — tool_result в сообщении
 *   пользователя;
 * - блоки размышлений (thinking) возвращаются Claude без изменений: панель
 *   получает их отдельным полем дельты thinking_block и хранит у хода;
 * - панель сокращает старые результаты и начало длинной беседы, а Claude
 *   привязывает размышления к неизменной истории. Чтобы такая правка не
 *   давала отказ 400, устаревшие блоки он отбрасывает сам (drop_block).
 */

import Anthropic from "@anthropic-ai/sdk";
import type { InternalMessage } from "./protocol.js";
import type { ChatTool } from "./responsesApi.js";

export const ANTHROPIC_BETAS = ["thinking-binding-controls-2026-08-01", "server-side-fallback-2026-07-01"];

/** Модели, у которых есть резерв при отказе фильтров безопасности. */
const FALLBACK_MODELS = new Set(["claude-fable-5-1", "claude-opus-5-5", "claude-sonnet-5-5"]);

/** Идентификатор вызова у Claude — только латиница, цифры, «_» и «-». */
export function toolUseId(id: string): string {
  const clean = id.replace(/[^A-Za-z0-9_-]/g, "_");
  return clean || "call";
}

function toolInput(args: string): Record<string, unknown> {
  try {
    const parsed = JSON.parse(args || "{}");
    return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed : {};
  } catch {
    return {};
  }
}

type AssistantWithThinking = Extract<InternalMessage, { role: "assistant" }> & { thinking_blocks?: unknown[] };

export function buildAnthropicParams(model: string, messages: InternalMessage[], tools: ChatTool[]) {
  // Системные сообщения в начале — правила панели и контекст задачи. Первое
  // (правила) не меняется от задачи к задаче, на нём точка кэша.
  const systemTexts: string[] = [];
  const out: Anthropic.Beta.BetaMessageParam[] = [];
  let pendingResults: Anthropic.Beta.BetaToolResultBlockParam[] = [];
  const flushResults = () => {
    if (pendingResults.length) out.push({ role: "user", content: pendingResults });
    pendingResults = [];
  };

  for (const message of messages) {
    if (message.role === "tool") {
      pendingResults.push({ type: "tool_result", tool_use_id: toolUseId(message.tool_call_id), content: message.content || "(пусто)" });
      continue;
    }
    flushResults();
    if (message.role === "system") {
      // Заметки панели посреди беседы — тоже в system: вставка между вызовом
      // и его результатом нарушила бы порядок, которого требует Claude.
      if (message.content) systemTexts.push(message.content);
      continue;
    }
    if (message.role === "user") {
      out.push({ role: "user", content: message.content || "(пусто)" });
      continue;
    }
    const assistant = message as AssistantWithThinking;
    const blocks: Anthropic.Beta.BetaContentBlockParam[] = [];
    // Размышления — первыми и без изменений: так их требует Claude.
    if (assistant.provider === "anthropic") {
      for (const block of assistant.thinking_blocks ?? []) blocks.push(block as Anthropic.Beta.BetaContentBlockParam);
    }
    if (assistant.content) blocks.push({ type: "text", text: assistant.content });
    for (const call of assistant.tool_calls ?? []) {
      blocks.push({ type: "tool_use", id: toolUseId(call.id), name: call.name, input: toolInput(call.arguments) });
    }
    if (blocks.length) out.push({ role: "assistant", content: blocks });
  }
  flushResults();
  if (!out.length || out[0].role !== "user") out.unshift({ role: "user", content: "(начало беседы)" });

  const [rules, ...rest] = systemTexts;
  const system: Anthropic.Beta.BetaTextBlockParam[] = [
    ...(rules ? [{ type: "text" as const, text: rules, cache_control: { type: "ephemeral" as const } }] : []),
    ...(rest.length ? [{ type: "text" as const, text: rest.join("\n\n") }] : [])
  ];

  return {
    model,
    max_tokens: 64000,
    system,
    messages: out,
    tools: tools.map((tool) => ({
      name: tool.function.name,
      description: tool.function.description ?? "",
      input_schema: (tool.function.parameters ?? { type: "object", properties: {} }) as Anthropic.Beta.BetaTool.InputSchema,
      // Большие записи (сотни ячеек) приходят по мере генерации; аргументы
      // панель всё равно разбирает и проверяет сама.
      eager_input_streaming: true
    })),
    tool_choice: { type: "auto" as const },
    // Кэш на конце беседы: следующий шаг задачи читает историю из кэша.
    cache_control: { type: "ephemeral" as const },
    thinking: { type: "adaptive" as const, block_binding: { prefix_mismatch_behavior: "drop_block" as const } },
    output_config: { effort: "medium" as const },
    betas: ANTHROPIC_BETAS,
    // Отказ фильтров безопасности — ответ резервной моделью, а не тупик.
    ...(FALLBACK_MODELS.has(model) ? { fallbacks: "default" as const } : {})
  };
}

/** Кусок SSE в формате Chat Completions — именно его ждёт панель. */
function chunk(payload: Record<string, unknown>): string {
  return `data: ${JSON.stringify(payload)}\n\n`;
}

const STOP_REASONS: Record<string, string> = {
  tool_use: "tool_calls",
  end_turn: "stop",
  stop_sequence: "stop",
  max_tokens: "length",
  pause_turn: "stop",
  refusal: "stop"
};

/** Поток событий Messages API → куски Chat Completions. */
export class AnthropicTranslator {
  private toolIndex = new Map<number, number>();
  private nextTool = 0;
  private thinking = new Map<number, Record<string, unknown>>();
  private usage = { input: 0, cacheRead: 0, cacheWrite: 0, output: 0 };
  private stopReason: string | null = null;
  private refusalCategory: string | null = null;
  private finished = false;

  event(event: any): string[] {
    switch (event?.type) {
      case "message_start": {
        const u = event.message?.usage ?? {};
        this.usage.input = u.input_tokens ?? 0;
        this.usage.cacheRead = u.cache_read_input_tokens ?? 0;
        this.usage.cacheWrite = u.cache_creation_input_tokens ?? 0;
        this.usage.output = u.output_tokens ?? 0;
        return [];
      }
      case "content_block_start": {
        const block = event.content_block ?? {};
        if (block.type === "tool_use") {
          const index = this.nextTool++;
          this.toolIndex.set(event.index, index);
          return [chunk({ choices: [{ index: 0, delta: { tool_calls: [{ index, id: block.id, type: "function", function: { name: block.name, arguments: "" } }] } }] })];
        }
        if (block.type === "thinking") this.thinking.set(event.index, { type: "thinking", thinking: block.thinking ?? "", signature: block.signature ?? "" });
        if (block.type === "redacted_thinking") this.thinking.set(event.index, { type: "redacted_thinking", data: block.data });
        if (block.type === "text" && block.text) return [chunk({ choices: [{ index: 0, delta: { content: block.text } }] })];
        return [];
      }
      case "content_block_delta": {
        const delta = event.delta ?? {};
        if (delta.type === "text_delta") return [chunk({ choices: [{ index: 0, delta: { content: delta.text } }] })];
        if (delta.type === "input_json_delta") {
          const index = this.toolIndex.get(event.index);
          if (index === undefined) return [];
          return [chunk({ choices: [{ index: 0, delta: { tool_calls: [{ index, function: { arguments: delta.partial_json ?? "" } }] } }] })];
        }
        const block = this.thinking.get(event.index);
        if (block && delta.type === "thinking_delta") block.thinking = String(block.thinking ?? "") + (delta.thinking ?? "");
        if (block && delta.type === "signature_delta") block.signature = delta.signature ?? "";
        return [];
      }
      case "content_block_stop": {
        const block = this.thinking.get(event.index);
        if (!block) return [];
        this.thinking.delete(event.index);
        return [chunk({ choices: [{ index: 0, delta: { thinking_block: block } }] })];
      }
      case "message_delta": {
        if (event.delta?.stop_reason) this.stopReason = event.delta.stop_reason;
        if (event.delta?.stop_details?.category) this.refusalCategory = event.delta.stop_details.category;
        const u = event.usage ?? {};
        if (typeof u.output_tokens === "number") this.usage.output = u.output_tokens;
        if (typeof u.input_tokens === "number") this.usage.input = u.input_tokens;
        if (typeof u.cache_read_input_tokens === "number") this.usage.cacheRead = u.cache_read_input_tokens;
        if (typeof u.cache_creation_input_tokens === "number") this.usage.cacheWrite = u.cache_creation_input_tokens;
        return [];
      }
      case "message_stop":
        return this.finish();
      default:
        return [];
    }
  }

  /** Расход в полях Chat Completions: кэш — часть входа, как у OpenAI. */
  usageChunk(): Record<string, unknown> {
    const { input, cacheRead, cacheWrite, output } = this.usage;
    return {
      prompt_tokens: input + cacheRead + cacheWrite,
      completion_tokens: output,
      total_tokens: input + cacheRead + cacheWrite + output,
      prompt_tokens_details: { cached_tokens: cacheRead }
    };
  }

  /** Расход для журнала сервера (usageMetrics.ts). */
  metricUsage() {
    const { input, cacheRead, cacheWrite, output } = this.usage;
    return { promptTokens: input + cacheRead + cacheWrite, completionTokens: output, totalTokens: input + cacheRead + cacheWrite + output, cachedTokens: cacheRead, cacheWriteTokens: cacheWrite };
  }

  finish(): string[] {
    if (this.finished) return [];
    this.finished = true;
    const pieces: string[] = [];
    if (this.stopReason === "refusal") {
      pieces.push(chunk({ choices: [{ index: 0, delta: { content: `\n\nClaude отказался продолжать этот ответ${this.refusalCategory ? ` (фильтр «${this.refusalCategory}»)` : ""}. Переформулируйте просьбу или выберите другую модель.` } }] }));
    }
    pieces.push(chunk({ choices: [{ index: 0, delta: {}, finish_reason: STOP_REASONS[this.stopReason ?? ""] ?? "stop" }] }));
    pieces.push(chunk({ choices: [], usage: this.usageChunk() }));
    pieces.push("data: [DONE]\n\n");
    return pieces;
  }

  get done(): boolean {
    return this.finished;
  }
}

/** Текст ошибки Anthropic для панели. */
export function anthropicErrorText(error: unknown): { status: number; message: string } {
  if (error instanceof Anthropic.AuthenticationError) return { status: 401, message: "Claude: ключ API не принят. Проверьте его в «Ключах»." };
  if (error instanceof Anthropic.PermissionDeniedError) return { status: 403, message: `Claude: нет доступа к модели. ${error.message}` };
  if (error instanceof Anthropic.RateLimitError) return { status: 429, message: "Claude: превышен лимит запросов или кончились средства на счёте. Подождите минуту или пополните баланс в консоли Anthropic." };
  if (error instanceof Anthropic.BadRequestError) return { status: 400, message: `Claude отклонил запрос: ${error.message}` };
  if (error instanceof Anthropic.APIError) return { status: typeof error.status === "number" ? error.status : 502, message: `Claude вернул ${error.status ?? "ошибку"}. ${error.message}` };
  return { status: 500, message: String((error as any)?.message ?? error) };
}
