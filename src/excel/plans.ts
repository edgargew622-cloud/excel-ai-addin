/**
 * Реестр операций, которые проходят через план с предпросмотром.
 *
 * До этапа 6 цикл агента перечислял такие операции по именам прямо в коде,
 * и каждый новый инструмент требовал править сам цикл. Теперь инструмент
 * объявляет здесь, как готовить план, как его исполнять и как отпустить
 * удержанные снимки при отказе или остановке.
 *
 * Смысл плана один для всех операций: подготовка ничего не меняет и её можно
 * показать пользователю, исполнение проверяет, что книга не изменилась
 * с момента предпросмотра, и сверяет результат.
 */

import {
  executeFormatRangePlan,
  executeSetRangePlan,
  executeSetRangesPlan,
  prepareFormatRangePlan,
  prepareSetRangePlan,
  prepareSetRangesPlan,
  releaseSetRangePlanSnapshot,
  releaseSetRangesPlanSnapshots,
  type FormatRangePlan,
  type SetRangePlan,
  type SetRangesPlan
} from "./excelTools";

export type OperationPlan = SetRangePlan | SetRangesPlan | FormatRangePlan;

export interface PlanDriver<P extends OperationPlan = OperationPlan> {
  prepare(args: unknown): Promise<P>;
  execute(plan: P): Promise<unknown>;
  /** Отказ, остановка и сбой освобождают закреплённые снимки. */
  release(plan: P): void;
}

const drivers: Record<string, PlanDriver<any>> = {
  set_range_values: {
    prepare: prepareSetRangePlan,
    execute: executeSetRangePlan,
    release: releaseSetRangePlanSnapshot
  },
  set_ranges_values: {
    prepare: prepareSetRangesPlan,
    execute: executeSetRangesPlan,
    release: releaseSetRangesPlanSnapshots
  },
  format_range: {
    prepare: prepareFormatRangePlan,
    execute: executeFormatRangePlan,
    // Оформление снимков не закрепляет: состояние до операции хранится
    // в самом плане, а он живёт не дольше подтверждения.
    release: () => undefined
  }
};

export function planDriverFor(toolName: string): PlanDriver | undefined {
  return drivers[toolName];
}

/** Имена операций с предпросмотром — для тестов и проверок полноты. */
export const PLANNED_TOOLS = Object.keys(drivers);
