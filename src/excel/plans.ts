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
  executeApplyFilterPlan,
  executeRowOpPlan,
  prepareDeleteRowsPlan,
  prepareInsertRowsPlan,
  executeFillRangePlan,
  executeFormatRangePlan,
  executeSortRangePlan,
  prepareApplyFilterPlan,
  prepareFillRangePlan,
  prepareSortRangePlan,
  executeSetRangePlan,
  executeSetRangesPlan,
  prepareFormatRangePlan,
  prepareSetRangePlan,
  prepareSetRangesPlan,
  releaseSetRangePlanSnapshot,
  releaseSetRangesPlanSnapshots,
  type ApplyFilterPlan,
  type RowOpPlan,
  type FillRangePlan,
  type FormatRangePlan,
  type SortRangePlan,
  type SetRangePlan,
  type SetRangesPlan
} from "./excelTools";
import {
  executeConditionalFormatPlan,
  executeCreateTablePlan,
  executeFreezePanesPlan,
  prepareConditionalFormatPlan,
  prepareCreateTablePlan,
  prepareFreezePanesPlan,
  type ConditionalFormatPlan,
  type CreateTablePlan,
  type FreezePanesPlan
} from "./sheetFormatPlans";
import { executeCreateChartPlan, prepareCreateChartPlan, type CreateChartPlan } from "./chartPlans";
import { executeCreatePivotPlan, prepareCreatePivotPlan, type CreatePivotPlan } from "./pivotPlans";
import {
  executeCreateSheetPlan,
  executeDeleteSheetPlan,
  executeRenameSheetPlan,
  prepareCreateSheetPlan,
  prepareDeleteSheetPlan,
  prepareRenameSheetPlan,
  type CreateSheetPlan,
  type DeleteSheetPlan,
  type RenameSheetPlan
} from "./sheetPlans";
import {
  executeCleanPlan,
  executeRemoveDuplicatesPlan,
  prepareConvertValuesPlan,
  prepareRemoveDuplicatesPlan,
  prepareTrimTextPlan,
  type CleanValuesPlan,
  type RemoveDuplicatesPlan
} from "./dataCleaning";

export type OperationPlan =
  | CleanValuesPlan
  | RemoveDuplicatesPlan
  | RenameSheetPlan
  | DeleteSheetPlan
  | SetRangePlan
  | SetRangesPlan
  | FillRangePlan
  | FormatRangePlan
  | SortRangePlan
  | ApplyFilterPlan
  | RowOpPlan
  | FreezePanesPlan
  | ConditionalFormatPlan
  | CreateTablePlan
  | CreateChartPlan
  | CreatePivotPlan
  | CreateSheetPlan;

export interface PlanDriver<P extends OperationPlan = OperationPlan> {
  prepare(args: unknown): Promise<P>;
  /** Сигнал остановки задачи: длинная операция проверяет его на безопасных границах. */
  execute(plan: P, signal?: AbortSignal): Promise<unknown>;
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
  fill_range: {
    prepare: prepareFillRangePlan,
    execute: executeFillRangePlan,
    release: () => undefined
  },
  format_range: {
    prepare: prepareFormatRangePlan,
    execute: executeFormatRangePlan,
    // Оформление снимков не закрепляет: состояние до операции хранится
    // в самом плане, а он живёт не дольше подтверждения.
    release: () => undefined
  },
  sort_range: {
    prepare: prepareSortRangePlan,
    execute: executeSortRangePlan,
    release: () => undefined
  },
  apply_filter: {
    prepare: prepareApplyFilterPlan,
    execute: executeApplyFilterPlan,
    release: () => undefined
  },
  // У вставки и удаления строк отката нет вовсе, поэтому снимков они
  // не закрепляют: всё, что можно обещать, живёт в самом плане.
  insert_rows: {
    prepare: prepareInsertRowsPlan,
    execute: executeRowOpPlan,
    release: () => undefined
  },
  delete_rows: {
    prepare: prepareDeleteRowsPlan,
    execute: executeRowOpPlan,
    release: () => undefined
  },
  freeze_panes: {
    prepare: prepareFreezePanesPlan,
    execute: executeFreezePanesPlan,
    release: () => undefined
  },
  add_conditional_format: {
    prepare: prepareConditionalFormatPlan,
    execute: executeConditionalFormatPlan,
    release: () => undefined
  },
  create_table: {
    prepare: prepareCreateTablePlan,
    execute: executeCreateTablePlan,
    release: () => undefined
  },
  create_chart: {
    prepare: prepareCreateChartPlan,
    execute: executeCreateChartPlan,
    release: () => undefined
  },
  create_pivot_table: {
    prepare: prepareCreatePivotPlan,
    execute: executeCreatePivotPlan,
    release: () => undefined
  },
  create_sheet: {
    prepare: prepareCreateSheetPlan,
    execute: executeCreateSheetPlan,
    release: () => undefined
  },
  // Очистка данных (этап 7, 7.2): общий исполнитель, у каждой — своя подготовка.
  trim_text: {
    prepare: prepareTrimTextPlan,
    execute: executeCleanPlan,
    release: () => undefined
  },
  convert_values: {
    prepare: prepareConvertValuesPlan,
    execute: executeCleanPlan,
    release: () => undefined
  },
  // Отмены нет: план ничего не удерживает.
  remove_duplicates: {
    prepare: prepareRemoveDuplicatesPlan,
    execute: executeRemoveDuplicatesPlan,
    release: () => undefined
  },
  // Листы (этап 7, 7.3.1).
  rename_sheet: {
    prepare: prepareRenameSheetPlan,
    execute: executeRenameSheetPlan,
    release: () => undefined
  },
  delete_sheet: {
    prepare: prepareDeleteSheetPlan,
    execute: executeDeleteSheetPlan,
    release: () => undefined
  }
};

export function planDriverFor(toolName: string): PlanDriver | undefined {
  return drivers[toolName];
}

/** Имена операций с предпросмотром — для тестов и проверок полноты. */
export const PLANNED_TOOLS = Object.keys(drivers);
