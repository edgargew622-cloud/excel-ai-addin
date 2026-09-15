import { streamChat, type ChatMessage, type ToolCall } from "../taskpane/api/client";
import {
  executeSetRangePlan,
  prepareSetRangePlan,
  preflightToolArgs,
  releaseSetRangePlanSnapshot,
  resolveToolArgs,
  runTool,
  ToolError,
  ToolExecutionError,
  type ExecutionState,
  type SetRangePlan
} from "../excel/excelTools";
import { TOOL_BY_NAME, toolsForApi, SYSTEM_PROMPT, writableAtCurrentStage } from "../excel/toolSchemas";
import { getActiveContext } from "../excel/workbookContext";

export const MAX_ITERATIONS = 20;
export const MAX_READ_CALLS = 30;
export const MAX_MUTATING_CALLS = 8;
export const MAX_TASK_ACTIVE_MS = 5 * 60_000;
export const MAX_TOOL_RESULT_BYTES = 128 * 1024;
export const MAX_REQUEST_BYTES = 1024 * 1024;

const byteLength = (value: string) => new TextEncoder().encode(value).length;

export interface ToolEvent {
  id: string;
  name: string;
  args: unknown;
  status: "running" | "done" | "error" | "rejected" | "cancelled" | "uncertain";
  result?: string;
  executionState?: ExecutionState;
  undoable?: boolean;
  undoNote?: string;
}

export interface AgentHooks {
  onDelta: (text: string) => void;
  /** Модель закончила текстовую часть шага — можно зафиксировать пузырь в UI. */
  onStepEnd: (text: string) => void;
  onToolEvent: (e: ToolEvent) => void;
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
function toolResult(ok: boolean, payload: unknown, executionState?: ExecutionState): string {
  return JSON.stringify(ok
    ? { ok: true, result: payload, executionState: executionState ?? "verified" }
    : { ok: false, error: String(payload), executionState: executionState ?? "failed_before_write" });
}

interface CallOutcome { content: string; stop: boolean }

function failedCall(call: ToolCall, hooks: AgentHooks, args: unknown, message: string): CallOutcome {
  hooks.onToolEvent({ id: call.id, name: call.name, args, status: "error", result: message, executionState: "failed_before_write" });
  return { content: toolResult(false, message, "failed_before_write"), stop: false };
}

async function executeCall(
  call: ToolCall,
  hooks: AgentHooks,
  taskSheet: string,
  onConfirmationWait: (milliseconds: number) => void,
  analysisOnly: boolean,
  signal?: AbortSignal,
  deadlineAt?: number
): Promise<CallOutcome> {
  const spec = TOOL_BY_NAME.get(call.name);

  if (!spec) {
    hooks.onToolEvent({
      id: call.id,
      name: call.name,
      args: call.arguments,
      status: "error",
      result: "неизвестный инструмент"
    });
    return { content: toolResult(false, `Инструмента "${call.name}" не существует. Используй только объявленные функции.`, "not_started"), stop: false };
  }

  let args: unknown;
  try {
    args = parseArgs(call.arguments);
  } catch {
    hooks.onToolEvent({ id: call.id, name: call.name, args: call.arguments, status: "error", result: "битый JSON" });
    return { content: toolResult(false, "Аргументы не разобрались как JSON. Пришли корректный объект.", "not_started"), stop: false };
  }

  // Фиксируем активный лист до подтверждения/исполнения. Это исключает
  // гонку, когда пользователь переключает вкладку между чтением и записью.
  try {
    args = await resolveToolArgs(call.name, args, taskSheet);
    preflightToolArgs(call.name, args);
  } catch (error: any) {
    return failedCall(call, hooks, args, error?.message ?? String(error));
  }

  if (signal?.aborted) throw new DOMException("Остановлено пользователем", "AbortError");

  if (analysisOnly && spec.mutating) {
    return failedCall(call, hooks, args, `Режим «Только анализ» запрещает инструмент ${call.name}. Операция не выполнялась.`);
  }
  if (spec.mutating && !writableAtCurrentStage(spec)) {
    return failedCall(call, hooks, args, `Инструмент ${call.name} ещё не подключён к проверяемому контуру этапа 3.`);
  }

  let preparedPlan: SetRangePlan | null = null;
  if (call.name === "set_range_values") {
    try { preparedPlan = await prepareSetRangePlan(args); }
    catch (error: any) { return failedCall(call, hooks, args, error?.message ?? String(error)); }
  }

  if (spec.destructive) {
    const confirmationStarted = Date.now();
    let allowed: boolean;
    try {
      allowed = await confirmWithAbort(() => hooks.confirm(call.name, preparedPlan ?? args), signal);
    } catch (error) {
      if (preparedPlan) releaseSetRangePlanSnapshot(preparedPlan);
      throw error;
    } finally {
      onConfirmationWait(Date.now() - confirmationStarted);
    }
    if (signal?.aborted) {
      if (preparedPlan) releaseSetRangePlanSnapshot(preparedPlan);
      throw new DOMException("Остановлено пользователем", "AbortError");
    }
    if (!allowed) {
      if (preparedPlan) releaseSetRangePlanSnapshot(preparedPlan);
      hooks.onToolEvent({ id: call.id, name: call.name, args, status: "rejected", executionState: "not_started" });
      return { content: toolResult(false, "Пользователь отклонил операцию. Не повторяй её, предложи другой путь или спроси уточнение.", "not_started"), stop: false };
    }
  }

  hooks.onToolEvent({ id: call.id, name: call.name, args, status: "running" });

  try {
    if (signal?.aborted) {
      if (preparedPlan) releaseSetRangePlanSnapshot(preparedPlan);
      throw new DOMException("Остановлено пользователем", "AbortError");
    }
    const result = preparedPlan
      ? await executeSetRangePlan(preparedPlan)
      : await runTool(call.name, args, { analysisOnly, signal, deadlineAt });
    const meta = result && typeof result === "object" ? (result as Record<string, unknown>) : null;
    const executionState = spec.mutating
      ? meta?.executionState === "verified" ? "verified" : "applied"
      : "not_started";
    const serialized = toolResult(true, result, executionState);
    if (byteLength(serialized) > MAX_TOOL_RESULT_BYTES) {
      if (spec.mutating) {
        const warning = "Операция могла изменить книгу, но её ответ превысил лимит. Не повторяйте её без проверки текущего состояния.";
        hooks.onToolEvent({ id: call.id, name: call.name, args, status: "uncertain", result: warning, executionState: "unknown" });
        return { content: toolResult(false, warning, "unknown"), stop: true };
      }
      const warning = `Ответ чтения превысил ${MAX_TOOL_RESULT_BYTES} байт. Прочитайте меньший диапазон.`;
      hooks.onToolEvent({ id: call.id, name: call.name, args, status: "error", result: warning, executionState: "not_started" });
      return {
        content: JSON.stringify({ ok: false, error: warning, incomplete: true, continuation: "read_smaller_range", executionState: "not_started" }),
        stop: false
      };
    }
    hooks.onToolEvent({
      id: call.id,
      name: call.name,
      args,
      status: "done",
      executionState,
      ...(typeof meta?.undoable === "boolean" ? { undoable: meta.undoable } : {}),
      ...(typeof meta?.undoNote === "string" ? { undoNote: meta.undoNote } : {})
    });
    return { content: serialized, stop: false };
  } catch (e: any) {
    if (e?.name === "AbortError" && !spec.mutating) throw e;
    const msg = e instanceof ToolError ? e.message : e?.message ?? String(e);
    const state = e instanceof ToolExecutionError ? e.executionState : spec.mutating ? "unknown" : "failed_before_write";
    const stop = state === "applied" || state === "unknown";
    hooks.onToolEvent({ id: call.id, name: call.name, args, status: stop ? "uncertain" : "error", result: msg, executionState: state });
    return { content: toolResult(false, msg, state), stop };
  }
}

export function cancellationToolMessage(call: ToolCall): ChatMessage {
  return {
    role: "tool",
    tool_call_id: call.id,
    content: toolResult(false, "Операция отменена пользователем до выполнения.", "not_started")
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
      result: "отменено пользователем",
      executionState: "not_started"
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
  analysisOnly?: boolean;
  initialContext?: Awaited<ReturnType<typeof getActiveContext>>;
}): Promise<void> {
  const analysisOnly = opts.analysisOnly === true;
  const tools = toolsForApi(analysisOnly);
  const activeContext = opts.initialContext ?? await getActiveContext();
  const messages: ChatMessage[] = [
    { role: "system", content: SYSTEM_PROMPT },
    {
      role: "system",
      content: `Минимальный контекст задачи (прочитан ${activeContext.readAt}): ${JSON.stringify(activeContext)}. ` +
        (analysisOnly ? "Режим «Только анализ»: любые изменения книги запрещены." : "Режим: анализ и подтверждаемые изменения.")
    },
    ...opts.history
  ];
  // Лист фиксируется на всю пользовательскую задачу, а не на отдельный tool call.
  // Явный sheet в аргументах модели всё равно имеет приоритет.
  const taskSheet = activeContext.activeSheet.name;
  const startedAt = Date.now();
  let confirmationWaitMs = 0;
  let readCalls = 0;
  let mutatingCalls = 0;
  const activeTime = () => Date.now() - startedAt - confirmationWaitMs;
  const stopWithNotice = (notice: string) => { opts.hooks.onStepEnd(notice); };

  for (let i = 0; i < MAX_ITERATIONS; i++) {
    if (activeTime() >= MAX_TASK_ACTIVE_MS) {
      stopWithNotice("Достигнут предел времени задачи. Выполненная часть сохранена; проверьте книгу перед продолжением.");
      return;
    }
    if (byteLength(JSON.stringify({ provider: opts.provider, model: opts.model, messages, tools })) > MAX_REQUEST_BYTES) {
      stopWithNotice("История и инструменты превысили предел размера запроса. Начните новую задачу или сократите историю.");
      return;
    }
    const timeout = AbortSignal.timeout(Math.max(1, MAX_TASK_ACTIVE_MS - activeTime()));
    const requestSignal = opts.signal ? AbortSignal.any([opts.signal, timeout]) : timeout;
    let step;
    try {
      step = await streamChat({
        provider: opts.provider,
        model: opts.model,
        messages,
        tools,
        signal: requestSignal,
        onDelta: opts.hooks.onDelta
      });
    } catch (error) {
      if (timeout.aborted && !opts.signal?.aborted) {
        stopWithNotice("Время ожидания модели истекло. Выполненная часть сохранена; продолжите отдельной задачей.");
        return;
      }
      throw error;
    }

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
        const spec = TOOL_BY_NAME.get(call.name);
        if (spec?.mutating) mutatingCalls += 1;
        else readCalls += 1;
        const budgetExceeded = activeTime() >= MAX_TASK_ACTIVE_MS || readCalls > MAX_READ_CALLS || mutatingCalls > MAX_MUTATING_CALLS;
        const outcome = budgetExceeded
          ? failedCall(call, opts.hooks, call.arguments, "Предел времени или числа вызовов достигнут; операция не выполнялась.")
          : await executeCall(
              call,
              opts.hooks,
              taskSheet,
              (ms) => { confirmationWaitMs += ms; },
              analysisOnly,
              opts.signal,
              Date.now() + Math.max(1, MAX_TASK_ACTIVE_MS - activeTime())
            );
        const toolMsg: ChatMessage = { role: "tool", tool_call_id: call.id, content: outcome.content };
        messages.push(toolMsg);
        opts.history.push(toolMsg);
        if (budgetExceeded || outcome.stop) {
          closeCancelledCalls(step.toolCalls, callIndex + 1, messages, opts.history, opts.hooks);
          stopWithNotice(outcome.stop
            ? "Выполнение остановлено: состояние книги после операции требует проверки. Не повторяйте правку автоматически."
            : "Лимит задачи достигнут. Выполненная часть сохранена; продолжите отдельной задачей.");
          return;
        }
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

  stopWithNotice(`Достигнут предел в ${MAX_ITERATIONS} ответов модели. Выполненная часть сохранена; продолжите отдельной задачей.`);
}
