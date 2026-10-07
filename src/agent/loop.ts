import { streamChat, type ChatMessage, type StepUsage, type ToolCall } from "../taskpane/api/client";
import {
  preflightToolArgs,
  resolveToolArgs,
  runTool,
  ToolError,
  ToolExecutionError,
  type ExecutionState
} from "../excel/excelTools";
import { planDriverFor, type OperationPlan, type PlanDriver } from "../excel/plans";
import { lastUserRequest, READ_PERMISSION, ReadScope, sheetsReadBy, type ScopeIO } from "./readScope";
import { excelScopeIO } from "../excel/excelTools";
import { TOOL_BY_NAME, toolsForApi, SYSTEM_PROMPT, writableAtCurrentStage } from "../excel/toolSchemas";
import { getActiveContext } from "../excel/workbookContext";
import { WEB_PANEL } from "../taskpane/panelMode";

export const MAX_ITERATIONS = 20;
export const MAX_READ_CALLS = 30;
export const MAX_MUTATING_CALLS = 8;
/**
 * Оформление и интернет — свои пределы (05.10.2026, «Книга18»). Прежде рамки,
 * цвет шапки и формат чисел считались наравне с записью данных: таблица на
 * три блока упиралась в «8 изменений», и пользователь четырежды писал
 * «продолжай». Поиск по Грузии съедал предел чтений листа. Защита от массовой
 * правки данных осталась прежней — 8 изменений; оформление обратимо и данных
 * не трогает.
 */
export const MAX_FORMAT_CALLS = 20;
export const MAX_WEB_CALLS = 24;
// Диаграммы и условное оформление ячеек не меняют («Книга20»: подсветка строк
// стала девятым «изменением» и потребовала «продолжай»).
const FORMAT_TOOLS: ReadonlySet<string> = new Set([
  "format_range", "format_chart", "freeze_panes", "set_page_layout",
  "create_chart", "add_conditional_format", "move_conditional_format", "apply_color_convention",
  // Новый пустой лист данных не меняет («Книга15»: два листа — четверть предела).
  "create_sheet",
  // Срезы и вид листа (10.7): кнопки и сетка, данные не меняются.
  "add_slicer", "set_sheet_view",
  // Дашборд (10.8): вид диаграмм, их место и отбор в сводных — данные не меняются.
  "edit_chart", "arrange_charts", "filter_pivots",
  // Удаление диаграммы и сводной (08.10): данные источника не меняются.
  "delete_chart", "delete_pivot", "delete_slicer"
]);
const WEB_TOOLS: ReadonlySet<string> = new Set(["web_search", "read_web_page"]);

/** На какой предел задачи идёт вызов. */
export function callBudget(name: string, mutating: boolean): "write" | "format" | "web" | "read" {
  if (FORMAT_TOOLS.has(name)) return "format";
  if (WEB_TOOLS.has(name)) return "web";
  return mutating ? "write" : "read";
}

export const LIMITS_TEXT = `не больше ${MAX_MUTATING_CALLS} изменений данных, ${MAX_FORMAT_CALLS} действий оформления, диаграмм и новых листов, ` +
  `${MAX_READ_CALLS} чтений книги, ${MAX_WEB_CALLS} обращений к интернету`;
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
  /** Расход каждого обращения к модели — для итога задачи в панели. */
  onUsage?: (usage: StepUsage) => void;
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

/** Аргументы для события и отчёта: битый JSON не должен ломать разбор дубликата. */
function parseArgsSafely(raw: string): unknown {
  try { return parseArgs(raw); } catch { return raw; }
}

/** Одинаковый ли это изменяющий вызов внутри одного шага модели.
 * Чтения не считаются: повторное чтение безвредно и иногда осмысленно. */
export function duplicateMutatingCall(call: ToolCall, seen: Set<string>): boolean {
  const spec = TOOL_BY_NAME.get(call.name);
  if (!spec?.mutating) return false;
  let args: unknown;
  try { args = parseArgs(call.arguments); } catch { args = call.arguments; }
  const signature = `${call.name}:${JSON.stringify(args)}`;
  if (seen.has(signature)) return true;
  seen.add(signature);
  return false;
}

/** Слова, которыми пользователь просит что-то запомнить (8.5). */
export const MEMORY_INTENT = /запомн|всегда|по умолчанию|впредь|в дальнейшем|сохрани[\s\S]{0,40}сценари|сценари[\s\S]{0,40}сохран|remember|always|by default|save[\s\S]{0,40}scenario/i;

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
  deadlineAt?: number,
  readAccess?: { scope: ReadScope; io: ScopeIO; request?: string; webEnabled?: boolean }
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

  if (analysisOnly && (spec.mutating || spec.sideEffect)) {
    return failedCall(call, hooks, args, `Режим «Только анализ» запрещает инструмент ${call.name}. Операция не выполнялась.`);
  }
  // Память (8.5): ничего не запоминается без прямой просьбы пользователя. Подсказка
  // в ячейке «запомни: …» — не просьба: проверяется его собственное сообщение.
  if (spec.sideEffect === "memory" && !MEMORY_INTENT.test(readAccess?.request ?? "")) {
    return failedCall(call, hooks, args,
      "Сохранять в память можно только по прямой просьбе пользователя в его сообщении («запомни», «всегда», «по умолчанию», «сохрани сценарий»). " +
      "Он этого не просил — не сохраняй; если считаешь полезным, предложи ему словами.");
  }
  // Интернет (8.7) — только если пользователь его включил: модель могла вызвать
  // инструмент по памяти из прошлой задачи, когда он был включён.
  if (spec.needsWeb && !readAccess?.webEnabled) {
    return failedCall(call, hooks, args, "Поиск в интернете выключен: пользователь включает его галочкой «Интернет» в панели. Скажи ему об этом, если без интернета не обойтись.");
  }
  // Панель из интернета (Mac, Excel в браузере): локального сервера нет.
  if (spec.needsLocal && WEB_PANEL) {
    return failedCall(call, hooks, args, "Этой панели недоступны копия книги, память, файлы и поиск в интернете: они работают только в версии для Windows с программой на компьютере. Скажи пользователю об этом, если без них не обойтись.");
  }
  if (spec.mutating && !writableAtCurrentStage(spec)) {
    return failedCall(call, hooks, args, `Инструмент ${call.name} ещё не переведён на проверяемый путь с предпросмотром и сверкой результата, поэтому модели не выдаётся.`);
  }

  // Границы чтения (8.0.6): лист вне просьбы читается только с разрешения
  // пользователя. Проверка до Excel: отказ не должен прочитать ни ячейки.
  if (readAccess && !spec.mutating && !spec.sideEffect) {
    let needed: string[];
    try { needed = await sheetsReadBy(call.name, args as Record<string, any>, readAccess.io); }
    catch (error: any) { return failedCall(call, hooks, args, error?.message ?? String(error)); }
    const outside = readAccess.scope.outside(needed);
    if (outside.length) {
      // Лист, которого нет (опечатка модели: «86д» вместо «86d», 01.10.2026), —
      // не повод спрашивать пользователя: сразу назвать настоящие листы.
      try {
        const existing = await readAccess.io.allSheets();
        const same = (p: string, q: string) => p.trim().toLowerCase() === q.trim().toLowerCase();
        const missing = outside.filter((sheet) => !existing.some((name) => same(name, sheet)));
        if (missing.length) {
          return failedCall(call, hooks, args,
            `${missing.length > 1 ? "Листов" : "Листа"} ${missing.map((sheet) => `«${sheet}»`).join(", ")} в книге нет. ` +
            `Листы книги: ${existing.map((name) => `«${name}»`).join(", ")}. Проверь имя — скопируй его отсюда.`);
        }
      } catch { /* список листов не прочитался — спросим как обычно */ }
      const started = Date.now();
      let allowed: boolean;
      try {
        allowed = await confirmWithAbort(() => hooks.confirm(READ_PERMISSION, { sheets: outside, tool: call.name, args }), signal);
      } finally {
        onConfirmationWait(Date.now() - started);
      }
      if (signal?.aborted) throw new DOMException("Остановлено пользователем", "AbortError");
      if (!allowed) {
        hooks.onToolEvent({ id: call.id, name: call.name, args, status: "rejected", executionState: "not_started" });
        return {
          content: toolResult(false,
            `Пользователь не разрешил читать ${outside.length > 1 ? "листы" : "лист"} ${outside.map((sheet) => `«${sheet}»`).join(", ")}. ` +
            "Их данные в этой задаче недоступны: не пытайся прочитать их другим инструментом или через имя диапазона. " +
            "Продолжай с разрешёнными листами или спроси пользователя.",
            "not_started"),
          stop: false
        };
      }
      readAccess.scope.allow(outside);
    }
  }

  // Какие операции проходят через план, знает реестр, а не этот цикл: иначе
  // каждый новый инструмент с предпросмотром требовал бы править цикл.
  const driver: PlanDriver | undefined = planDriverFor(call.name);
  let preparedPlan: OperationPlan | null = null;
  if (driver) {
    try { preparedPlan = await driver.prepare(args); }
    catch (error: any) { return failedCall(call, hooks, args, error?.message ?? String(error)); }
  }
  const releasePlanSnapshots = (plan: OperationPlan) => driver?.release(plan);

  if (spec.destructive) {
    const confirmationStarted = Date.now();
    let allowed: boolean;
    try {
      allowed = await confirmWithAbort(() => hooks.confirm(call.name, preparedPlan ?? args), signal);
    } catch (error) {
      if (preparedPlan) releasePlanSnapshots(preparedPlan);
      throw error;
    } finally {
      onConfirmationWait(Date.now() - confirmationStarted);
    }
    if (signal?.aborted) {
      if (preparedPlan) releasePlanSnapshots(preparedPlan);
      throw new DOMException("Остановлено пользователем", "AbortError");
    }
    if (!allowed) {
      if (preparedPlan) releasePlanSnapshots(preparedPlan);
      hooks.onToolEvent({ id: call.id, name: call.name, args, status: "rejected", executionState: "not_started" });
      return { content: toolResult(false, "Пользователь видел предпросмотр этой операции в панели и отклонил её; книга не менялась. Сам предпросмотр тебе не передаётся — не утверждай, что его не было. Не повторяй операцию, предложи другой путь или спроси уточнение.", "not_started"), stop: false };
    }
  }

  // Лист, изменение которого пользователь подтвердил, он видит: читать его дальше можно.
  if (readAccess && spec.mutating) {
    const a = args as Record<string, unknown>;
    // Лист, который создаёт сама операция, — из плана: у import_file_layout имя
    // выбирает панель (живая проверка 01.10.2026: агент не мог проверить свою копию).
    const plan = preparedPlan as unknown as Record<string, unknown> | null;
    readAccess.scope.allow([a.sheet, a.destSheet, a.newSheet, plan?.sheetName, plan?.destSheet].filter((value): value is string => typeof value === "string"));
  }

  hooks.onToolEvent({ id: call.id, name: call.name, args, status: "running" });

  try {
    if (signal?.aborted) {
      if (preparedPlan) releasePlanSnapshots(preparedPlan);
      throw new DOMException("Остановлено пользователем", "AbortError");
    }
    const result = preparedPlan && driver
      ? await driver.execute(preparedPlan, signal)
      : await runTool(call.name, args, { analysisOnly, signal, deadlineAt });
    const meta = result && typeof result === "object" ? (result as Record<string, unknown>) : null;
    // Групповая запись сообщает итог сама и может вернуть «unknown»: часть
    // операций выполнена, итог одной неизвестен. Сводить это к «applied»
    // нельзя — тогда цикл продолжит работу так, будто всё ясно.
    const reported = typeof meta?.executionState === "string" ? meta.executionState : null;
    const executionState: ExecutionState = spec.mutating
      ? reported === "verified" || reported === "unknown" || reported === "failed_before_write"
        ? reported
        : "applied"
      : "not_started";
    // Операция может вернуть неуспех, не бросая исключения: группа, которая
    // остановилась на полпути, отвечает ok: false и applied. Проверка
    // в плане стабилизации (S4.5): такой ответ обёртка называла ok: true,
    // панель показывала успехом, а цикл продолжал задачу — остановку
    // вызывал только возвращённый unknown.
    const reportedFailure = spec.mutating && meta?.ok === false;
    const serialized = reportedFailure
      ? JSON.stringify({ ok: false, result, executionState })
      : toolResult(true, result, executionState);
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
    // То же правило, что для отказа исключением: изменение с неподтверждённым
    // итогом останавливает задачу, следующий шаг начинается с чтения книги.
    const stopAfter = spec.mutating && (executionState === "unknown" || executionState === "applied" && (reportedFailure || reported === "applied"));
    hooks.onToolEvent({
      id: call.id,
      name: call.name,
      args,
      status: stopAfter ? "uncertain" : reportedFailure ? "error" : "done",
      executionState,
      ...(typeof meta?.undoable === "boolean" ? { undoable: meta.undoable } : {}),
      ...(typeof meta?.undoNote === "string" ? { undoNote: meta.undoNote } : {})
    });
    // Неопределённый или неподтверждённый итог останавливает задачу:
    // следующий шаг должен начинаться с чтения книги, а не с продолжения
    // по предположению.
    return { content: serialized, stop: stopAfter };
  } catch (e: any) {
    if (e?.name === "AbortError" && !spec.mutating) throw e;
    const msg = e instanceof ToolError ? e.message : e?.message ?? String(e);
    const state = e instanceof ToolExecutionError ? e.executionState : spec.mutating ? "unknown" : "failed_before_write";
    const stop = state === "applied" || state === "unknown";
    hooks.onToolEvent({ id: call.id, name: call.name, args, status: stop ? "uncertain" : "error", result: msg, executionState: state });
    return { content: toolResult(false, msg, state), stop };
  }
}

export function cancellationToolMessage(call: ToolCall, reason = "Операция отменена пользователем до выполнения."): ChatMessage {
  return {
    role: "tool",
    tool_call_id: call.id,
    content: toolResult(false, reason, "not_started")
  };
}

/** Что осталось невыполненным — для сообщения о пределе: «format_range Лист1!A42». */
export function pendingCallsSummary(calls: readonly ToolCall[]): string {
  return calls.map((call) => {
    const args = parseArgsSafely(call.arguments) as Record<string, unknown> | null;
    const sheet = typeof args?.sheet === "string" && args.sheet ? `${args.sheet}!` : "";
    const place = [args?.address, args?.sourceAddress, args?.destAddress].find((value) => typeof value === "string" && value) as string | undefined;
    return `${call.name}${place ? ` ${sheet}${place}` : ""}`;
  }).join("; ");
}

export function cancellationToolMessages(calls: ToolCall[], startIndex = 0): ChatMessage[] {
  return calls.slice(startIndex).map((call) => cancellationToolMessage(call));
}

function closeCancelledCalls(
  calls: ToolCall[],
  startIndex: number,
  messages: ChatMessage[],
  history: ChatMessage[],
  hooks: AgentHooks,
  /** Не выполнено из-за предела задачи, а не отменено пользователем. */
  byLimit = false
) {
  for (let i = startIndex; i < calls.length; i++) {
    const call = calls[i];
    const msg = cancellationToolMessage(call, byLimit ? "Не выполнялось: достигнут предел задачи." : undefined);
    messages.push(msg);
    history.push(msg);
    hooks.onToolEvent({
      id: call.id,
      name: call.name,
      args: call.arguments,
      status: "cancelled",
      result: byLimit ? "не выполнено: предел задачи" : "отменено пользователем",
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
  /** Время на задачу; по умолчанию MAX_TASK_ACTIVE_MS. Медленным моделям — больше. */
  taskBudgetMs?: number;
  /** Как узнать листы книги и лист имени диапазона; по умолчанию — из Excel. Для тестов. */
  scopeIO?: ScopeIO;
  /** Сохранённые предпочтения и сценарии (8.5) — блоком в начале задачи. */
  memoryPrompt?: string | null;
  /** Прикреплённые файлы (8.6): какие есть; содержимое модель читает сама, частями. */
  filesPrompt?: string | null;
  /** Поиск в интернете (8.7) включён пользователем в панели. */
  webEnabled?: boolean;
}): Promise<void> {
  const analysisOnly = opts.analysisOnly === true;
  const taskBudgetMs = opts.taskBudgetMs && opts.taskBudgetMs > 0 ? opts.taskBudgetMs : MAX_TASK_ACTIVE_MS;
  const tools = toolsForApi(analysisOnly, opts.webEnabled === true, !WEB_PANEL);
  const activeContext = opts.initialContext ?? await getActiveContext();
  const messages: ChatMessage[] = [
    { role: "system", content: SYSTEM_PROMPT },
    {
      role: "system",
      content: `Минимальный контекст задачи (прочитан ${activeContext.readAt}): ${JSON.stringify(activeContext)}. ` +
        (analysisOnly
          // «Книга19», 03.10.2026: модель отсылала к «переключателю в панели»,
          // и три просьбы подряд ушли впустую. Где он — говорим точно.
          ? "Режим «Только анализ»: любые изменения книги запрещены. Если просят изменить книгу, ответьте коротко: " +
            "«Снимите галочку „Только анализ“ над чатом и повторите просьбу» — и что именно будет сделано. " +
            "Режим проверяется заново при каждой просьбе: не утверждайте, что галочка «не применилась»."
          : "Режим: анализ и подтверждаемые изменения.")
    },
    ...(opts.memoryPrompt ? [{ role: "system" as const, content: opts.memoryPrompt }] : []),
    ...(opts.filesPrompt ? [{ role: "system" as const, content: opts.filesPrompt }] : []),
    ...opts.history
  ];
  // Лист фиксируется на всю пользовательскую задачу, а не на отдельный tool call.
  // Явный sheet в аргументах модели всё равно имеет приоритет.
  const taskSheet = activeContext.activeSheet.name;
  const readAccess = { scope: new ReadScope(taskSheet, lastUserRequest(opts.history)), io: opts.scopeIO ?? excelScopeIO, request: lastUserRequest(opts.history), webEnabled: opts.webEnabled === true };
  const startedAt = Date.now();
  let confirmationWaitMs = 0;
  const used = { write: 0, format: 0, web: 0, read: 0 };
  const caps = { write: MAX_MUTATING_CALLS, format: MAX_FORMAT_CALLS, web: MAX_WEB_CALLS, read: MAX_READ_CALLS };
  const activeTime = () => Date.now() - startedAt - confirmationWaitMs;
  const stopWithNotice = (notice: string) => { opts.hooks.onStepEnd(notice); };

  for (let i = 0; i < MAX_ITERATIONS; i++) {
    if (activeTime() >= taskBudgetMs) {
      stopWithNotice("Достигнут предел времени задачи. Выполненная часть сохранена; проверьте книгу перед продолжением.");
      return;
    }
    if (byteLength(JSON.stringify({ provider: opts.provider, model: opts.model, messages, tools })) > MAX_REQUEST_BYTES) {
      stopWithNotice("История и инструменты превысили предел размера запроса. Начните новую задачу или сократите историю.");
      return;
    }
    const timeout = AbortSignal.timeout(Math.max(1, taskBudgetMs - activeTime()));
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
      if (step.usage) opts.hooks.onUsage?.(step.usage);
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

    // Один шаг модели не должен дважды изменить книгу одним и тем же вызовом.
    // Проверка в Excel: на повторённое сообщение модель выдала три одинаковых
    // вызова подряд. Для чтения это безвредно, для записи — нет.
    const seenMutating = new Set<string>();
    for (let callIndex = 0; callIndex < step.toolCalls.length; callIndex++) {
      const call = step.toolCalls[callIndex];
      const duplicate = duplicateMutatingCall(call, seenMutating);
      if (duplicate) {
        const message = `Этот же вызов ${call.name} уже есть в текущем шаге с теми же аргументами. ` +
          "Повторный вызов не выполнялся: одна просьба не должна менять книгу дважды. " +
          "Если изменение действительно нужно повторить, дождись результата первого вызова и объясни, зачем повтор.";
        const outcome = failedCall(call, opts.hooks, parseArgsSafely(call.arguments), message);
        const toolMsg: ChatMessage = { role: "tool", tool_call_id: call.id, content: outcome.content };
        messages.push(toolMsg);
        opts.history.push(toolMsg);
        continue;
      }
      if (opts.signal?.aborted) {
        closeCancelledCalls(step.toolCalls, callIndex, messages, opts.history, opts.hooks);
        throw new DOMException("Остановлено пользователем", "AbortError");
      }

      try {
        const spec = TOOL_BY_NAME.get(call.name);
        const budget = callBudget(call.name, spec?.mutating === true);
        used[budget] += 1;
        const budgetExceeded = activeTime() >= taskBudgetMs || used[budget] > caps[budget];
        const outcome = budgetExceeded
          ? failedCall(call, opts.hooks, call.arguments, "Предел времени или числа вызовов достигнут; операция не выполнялась.")
          : await executeCall(
              call,
              opts.hooks,
              taskSheet,
              (ms) => { confirmationWaitMs += ms; },
              analysisOnly,
              opts.signal,
              Date.now() + Math.max(1, taskBudgetMs - activeTime()),
              readAccess
            );
        const toolMsg: ChatMessage = { role: "tool", tool_call_id: call.id, content: outcome.content };
        messages.push(toolMsg);
        opts.history.push(toolMsg);
        if (budgetExceeded || outcome.stop) {
          closeCancelledCalls(step.toolCalls, callIndex + 1, messages, opts.history, opts.hooks, budgetExceeded && !outcome.stop);
          // Живая беседа 01.10.2026: сообщение о пределе не говорило, что осталось, —
          // четыре заголовка тихо остались без оформления. Теперь — перечень и как продолжить.
          const left = pendingCallsSummary(step.toolCalls.slice(callIndex));
          stopWithNotice(outcome.stop
            ? "Выполнение остановлено: состояние книги после операции требует проверки. Не повторяйте правку автоматически."
            : `Достигнут предел одной задачи (${LIMITS_TEXT} и ${Math.round(taskBudgetMs / 60_000)} минут). ` +
              `Выполненная часть сохранена. Не выполнено: ${left}. Чтобы доделать, напишите «продолжай».`);
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
