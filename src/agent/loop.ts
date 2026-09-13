import { streamChat, type ChatMessage, type ToolCall } from "../taskpane/api/client";
import { getActiveSheetName, resolveToolArgs, runTool, ToolError } from "../excel/excelTools";
import { TOOL_BY_NAME, toolsForApi, SYSTEM_PROMPT } from "../excel/toolSchemas";
import { createTaskBudget, type TaskBudget } from "./budget";
import { byteLength, type OpMetric } from "./metrics";

export interface ToolEvent {
  id: string;
  name: string;
  args: unknown;
  status: "running" | "done" | "error" | "rejected" | "cancelled";
  result?: string;
  /** Длительность выполнения инструмента, миллисекунды. */
  ms?: number;
  /** Размер результата, переданного модели, в байтах. */
  bytes?: number;
  undoable?: boolean;
  undoNote?: string;
}

export interface AgentHooks {
  onDelta: (text: string) => void;
  /** Модель закончила текстовую часть шага — можно зафиксировать пузырь в UI. */
  onStepEnd: (text: string) => void;
  onToolEvent: (e: ToolEvent) => void;
  /** Длительность и размер одной операции. Данных книги здесь нет. */
  onMetric?: (m: OpMetric) => void;
  /** Вернуть true, если пользователь разрешил разрушительную операцию. */
  confirm: (name: string, args: unknown) => Promise<boolean>;
}

function parseArgs(raw: string): unknown {
  if (!raw || !raw.trim()) return {};
  try {
    return JSON.parse(raw);
  } catch {
    // Некоторые модели оборачивают JSON в ```json ... ```
    const stripped = raw.replace(/^```(?:json)?/i, "").replace(/```$/, "").trim();
    return JSON.parse(stripped);
  }
}


async function confirmWithAbort(
  confirm: () => Promise<boolean>,
  signal?: AbortSignal
): Promise<boolean> {
  if (!signal) return confirm();
  if (signal.aborted) throw new DOMException("Остановлено пользователем", "AbortError");

  return new Promise<boolean>((resolve, reject) => {
    const onAbort = () => {
      cleanup();
      reject(new DOMException("Остановлено пользователем", "AbortError"));
    };
    const cleanup = () => signal.removeEventListener("abort", onAbort);
    signal.addEventListener("abort", onAbort, { once: true });
    confirm().then(
      (value) => {
        cleanup();
        resolve(value);
      },
      (error) => {
        cleanup();
        reject(error);
      }
    );
  });
}

/** Результат инструмента для модели: всегда строка, всегда с признаком успеха. */
function toolResult(ok: boolean, payload: unknown): string {
  return JSON.stringify(ok ? { ok: true, result: payload } : { ok: false, error: String(payload) });
}

async function executeCall(
  call: ToolCall,
  hooks: AgentHooks,
  taskSheet: string,
  budget: TaskBudget,
  signal?: AbortSignal
): Promise<string> {
  const spec = TOOL_BY_NAME.get(call.name);

  if (!spec) {
    hooks.onToolEvent({
      id: call.id,
      name: call.name,
      args: call.arguments,
      status: "error",
      result: "неизвестный инструмент"
    });
    return toolResult(false, `Инструмента "${call.name}" не существует. Используй только объявленные функции.`);
  }

  let args: unknown;
  try {
    args = parseArgs(call.arguments);
  } catch {
    hooks.onToolEvent({ id: call.id, name: call.name, args: call.arguments, status: "error", result: "битый JSON" });
    return toolResult(false, "Аргументы не разобрались как JSON. Пришли корректный объект.");
  }

  // Фиксируем активный лист до подтверждения/исполнения. Это исключает
  // гонку, когда пользователь переключает вкладку между чтением и записью.
  args = await resolveToolArgs(call.name, args, taskSheet);

  if (signal?.aborted) throw new DOMException("Остановлено пользователем", "AbortError");

  // Бюджет проверяется до подтверждения: незачем спрашивать разрешение на
  // операцию, которую мы всё равно не выполним. Отказ возвращается моделью
  // как обычный результат инструмента, поэтому пара вызов-ответ остаётся целой
  // и модель может корректно завершить задачу.
  const verdict = budget.tryToolCall(spec.mutating);
  if (!verdict.ok) {
    hooks.onToolEvent({ id: call.id, name: call.name, args, status: "error", result: verdict.reason });
    return toolResult(false, verdict.reason);
  }

  if (spec.destructive) {
    const allowed = await confirmWithAbort(() => hooks.confirm(call.name, args), signal);
    if (signal?.aborted) throw new DOMException("Остановлено пользователем", "AbortError");
    if (!allowed) {
      hooks.onToolEvent({ id: call.id, name: call.name, args, status: "rejected" });
      return toolResult(false, "Пользователь отклонил операцию. Не повторяй её, предложи другой путь или спроси уточнение.");
    }
  }

  hooks.onToolEvent({ id: call.id, name: call.name, args, status: "running" });

  const startedAt = Date.now();
  try {
    if (signal?.aborted) throw new DOMException("Остановлено пользователем", "AbortError");
    const result = await runTool(call.name, args);
    const payload = toolResult(true, result);
    const ms = Date.now() - startedAt;
    const bytes = byteLength(payload);
    const meta = result && typeof result === "object" ? (result as Record<string, unknown>) : null;
    hooks.onToolEvent({
      id: call.id,
      name: call.name,
      args,
      status: "done",
      ms,
      bytes,
      ...(typeof meta?.undoable === "boolean" ? { undoable: meta.undoable } : {}),
      ...(typeof meta?.undoNote === "string" ? { undoNote: meta.undoNote } : {})
    });
    hooks.onMetric?.({ kind: "tool", name: call.name, ms, bytes, ok: true });
    return toolResult(true, result);
  } catch (e: any) {
    if (e?.name === "AbortError") throw e;
    const msg = e instanceof ToolError ? e.message : e?.message ?? String(e);
    const ms = Date.now() - startedAt;
    hooks.onToolEvent({ id: call.id, name: call.name, args, status: "error", result: msg, ms });
    hooks.onMetric?.({ kind: "tool", name: call.name, ms, bytes: byteLength(msg), ok: false });
    return toolResult(false, msg);
  }
}

export function cancellationToolMessage(call: ToolCall): ChatMessage {
  return {
    role: "tool",
    tool_call_id: call.id,
    content: toolResult(false, "Операция отменена пользователем до выполнения.")
  };
}

export function cancellationToolMessages(calls: ToolCall[], startIndex = 0): ChatMessage[] {
  return calls.slice(startIndex).map(cancellationToolMessage);
}

function closeCancelledCalls(
  calls: ToolCall[],
  startIndex: number,
  messages: ChatMessage[],
  history: ChatMessage[],
  hooks: AgentHooks
) {
  for (let i = startIndex; i < calls.length; i++) {
    const call = calls[i];
    const msg = cancellationToolMessage(call);
    messages.push(msg);
    history.push(msg);
    hooks.onToolEvent({
      id: call.id,
      name: call.name,
      args: call.arguments,
      status: "cancelled",
      result: "отменено пользователем"
    });
  }
}

/**
 * Гоняет модель до финального текстового ответа.
 * history мутируется на месте, чтобы UI видел ту же ленту сообщений.
 */
export async function runAgent(opts: {
  provider: string;
  model: string;
  history: ChatMessage[];
  hooks: AgentHooks;
  signal?: AbortSignal;
  /** Бюджеты задачи. Подставляется в тестах; по умолчанию — DEFAULT_BUDGETS. */
  budget?: TaskBudget;
  /** Разовые системные указания, например о восстановлении беседы. */
  notices?: string[];
}): Promise<void> {
  const tools = toolsForApi();
  // Разовые системные указания идут отдельными сообщениями роли system:
  // подделывать ход пользователя или модели для этого нельзя.
  const notices: ChatMessage[] = (opts.notices ?? []).map((content) => ({ role: "system", content }));
  const messages: ChatMessage[] = [
    { role: "system", content: SYSTEM_PROMPT },
    ...notices,
    ...opts.history
  ];
  // Лист фиксируется на всю пользовательскую задачу, а не на отдельный tool call.
  // Явный sheet в аргументах модели всё равно имеет приоритет.
  const taskSheet = await getActiveSheetName();
  const budget = opts.budget ?? createTaskBudget();

  while (true) {
    const stepVerdict = budget.tryModelStep();
    if (!stepVerdict.ok) {
      // Сообщение о бюджете видно пользователю, но в историю не попадает:
      // модель этого не говорила, и провайдеру такой ход отправлять нельзя.
      opts.hooks.onStepEnd(stepVerdict.reason);
      return;
    }

    const requestStartedAt = Date.now();
    const step = await streamChat({
      provider: opts.provider,
      model: opts.model,
      messages,
      tools,
      signal: opts.signal,
      onDelta: opts.hooks.onDelta
    });
    opts.hooks.onMetric?.({
      kind: "model",
      name: `${opts.provider}/${opts.model}`,
      ms: Date.now() - requestStartedAt,
      bytes: byteLength(step.content) + byteLength(step.reasoningContent),
      ok: true
    });

    const assistantMsg: ChatMessage = {
      role: "assistant",
      content: step.content || "",
      provider: opts.provider,
      ...(opts.provider === "deepseek" ? { reasoning_content: step.reasoningContent } : {}),
      ...(step.toolCalls.length ? { tool_calls: step.toolCalls } : {})
    };
    messages.push(assistantMsg);
    opts.history.push(assistantMsg);
    opts.hooks.onStepEnd(step.content);

    if (!step.toolCalls.length) return; // Финальный ответ.

    for (let callIndex = 0; callIndex < step.toolCalls.length; callIndex++) {
      const call = step.toolCalls[callIndex];
      if (opts.signal?.aborted) {
        closeCancelledCalls(step.toolCalls, callIndex, messages, opts.history, opts.hooks);
        throw new DOMException("Остановлено пользователем", "AbortError");
      }

      try {
        const content = await executeCall(call, opts.hooks, taskSheet, budget, opts.signal);
        const toolMsg: ChatMessage = { role: "tool", tool_call_id: call.id, content };
        messages.push(toolMsg);
        opts.history.push(toolMsg);
      } catch (error: any) {
        if (error?.name === "AbortError") {
          // assistant с tool_calls уже находится в history. Закрываем каждый
          // ещё не получивший ответа call синтетическим cancelled-result, чтобы
          // следующий запрос не содержал оборванную tool-пару.
          closeCancelledCalls(step.toolCalls, callIndex, messages, opts.history, opts.hooks);
        }
        throw error;
      }
    }
  }

}
