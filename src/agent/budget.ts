/**
 * Раздельные бюджеты на одну пользовательскую задачу.
 *
 * Один общий счётчик шагов наказывал за аккуратность: агент, который сначала
 * посмотрел карту книги, получал меньше попыток на саму работу. Читающие вызовы
 * дёшевы для книги и ограничиваются объёмом контекста, а не числом шагов;
 * изменяющие дороги по последствиям, и их предел должен быть отдельным и малым.
 *
 * Значения начальные и пересматриваются по результатам измерений.
 */

export interface BudgetLimits {
  /** Ответы модели за задачу. */
  modelSteps: number;
  /** Вызовы читающих инструментов. */
  readCalls: number;
  /** Вызовы изменяющих инструментов. */
  writeCalls: number;
  /** Полное время задачи, миллисекунды. */
  taskMs: number;
}

export const DEFAULT_BUDGETS: BudgetLimits = {
  modelSteps: 20,
  readCalls: 30,
  writeCalls: 8,
  taskMs: 5 * 60_000
};

export type BudgetKind = "modelSteps" | "readCalls" | "writeCalls" | "taskMs";

export interface BudgetDenial {
  ok: false;
  kind: BudgetKind;
  /** Текст для модели или для панели. */
  reason: string;
}

export type BudgetVerdict = { ok: true } | BudgetDenial;

const OK: BudgetVerdict = { ok: true };

export interface TaskBudget {
  /** Разрешает очередной запрос к модели. */
  tryModelStep(): BudgetVerdict;
  /** Разрешает очередной вызов инструмента. */
  tryToolCall(mutating: boolean): BudgetVerdict;
  used(): { modelSteps: number; readCalls: number; writeCalls: number; elapsedMs: number };
}

export function createTaskBudget(
  limits: BudgetLimits = DEFAULT_BUDGETS,
  now: () => number = Date.now
): TaskBudget {
  const startedAt = now();
  let modelSteps = 0;
  let readCalls = 0;
  let writeCalls = 0;

  const timeVerdict = (): BudgetVerdict => {
    const elapsed = now() - startedAt;
    if (elapsed < limits.taskMs) return OK;
    return {
      ok: false,
      kind: "taskMs",
      reason:
        `Исчерпан предел времени на задачу (${Math.round(limits.taskMs / 1000)} с). ` +
        "Работа остановлена. Сформулируйте следующий шаг отдельным сообщением."
    };
  };

  return {
    tryModelStep() {
      const time = timeVerdict();
      if (!time.ok) return time;
      if (modelSteps >= limits.modelSteps) {
        return {
          ok: false,
          kind: "modelSteps",
          reason:
            `Достигнут предел в ${limits.modelSteps} шагов модели. Работа остановлена, ` +
            "чтобы не зациклиться. Сформулируйте задачу мельче."
        };
      }
      modelSteps += 1;
      return OK;
    },

    tryToolCall(mutating: boolean) {
      const time = timeVerdict();
      if (!time.ok) return time;

      if (mutating) {
        if (writeCalls >= limits.writeCalls) {
          return {
            ok: false,
            kind: "writeCalls",
            reason:
              `Исчерпан предел изменяющих операций за задачу (${limits.writeCalls}). ` +
              "Опиши пользователю, что уже сделано и что осталось, не вызывая больше изменяющих инструментов."
          };
        }
        writeCalls += 1;
        return OK;
      }

      if (readCalls >= limits.readCalls) {
        return {
          ok: false,
          kind: "readCalls",
          reason:
            `Исчерпан предел чтений за задачу (${limits.readCalls}). ` +
            "Ответь по тому, что уже прочитано, и скажи, каких данных не хватает."
        };
      }
      readCalls += 1;
      return OK;
    },

    used() {
      return { modelSteps, readCalls, writeCalls, elapsedMs: now() - startedAt };
    }
  };
}

/** Исчерпание какого бюджета останавливает всю задачу, а какое — только вызов. */
export function stopsTask(kind: BudgetKind): boolean {
  return kind === "modelSteps" || kind === "taskMs";
}
