/**
 * Общие гарантии изменяющих операций (план стабилизации, S4).
 *
 * Эти тесты проверяют не отдельный инструмент, а поведение, которое обязано
 * быть одинаковым у всех: чужая книга не получает записи, частичный итог
 * останавливает задачу, остановка не даёт начаться следующим записям,
 * закреплённые снимки освобождаются. Там, где это возможно, тест перебирает
 * весь реестр планов, а не перечисляет инструменты по именам: новый
 * инструмент без проверки иначе прошёл бы незамеченным.
 */

import test from "node:test";
import assert from "node:assert/strict";
import { runAgent } from "./loop";
import { PLANNED_TOOLS, planDriverFor } from "../excel/plans";
import { executeSetRangesPlan, prepareSetRangesPlan } from "../excel/excelTools";
import { resetSnapshotStore, snapshotStoreStats } from "../excel/snapshotStore";

const initialContext: any = {
  workbook: { sessionId: "s", documentUrl: "C:/books/a.xlsx", identityConfirmed: true },
  activeSheet: { id: "sheet-1", name: "Sheet1" },
  activeCell: "A1",
  selectedAreas: ["A1"],
  readAt: "2026-09-24T00:00:00Z",
  revision: 0,
  revisionCoverage: {},
  capabilities: {}
};

/**
 * Лист с ячейками по адресам. `ignore` — ячейки, запись в которые Excel
 * молча не принимает (как неугловая ячейка объединения).
 */
function sheetExcel(options: { ignore?: string[]; onWrite?: (address: string) => void } = {}) {
  const cells = new Map<string, unknown>();
  const writes: string[] = [];
  const rangeAt = (address: string): any => {
    const key = address.replace(/^.*!/, "");
    return {
      address: `Sheet1!${key}`, rowCount: 1, columnCount: 1, rowIndex: 0, columnIndex: 0,
      load: () => undefined,
      get values() { return [[cells.get(key) ?? ""]]; },
      set values(value: unknown[][]) {
        writes.push(key);
        options.onWrite?.(key);
        if (!options.ignore?.includes(key)) cells.set(key, value[0][0]);
      },
      get formulas() { return [[cells.get(key) ?? ""]]; },
      set formulas(value: unknown[][]) { this.values = value; },
      get valueTypes() { return [["Double"]]; },
      getCell() { return this; }
    };
  };
  const sheet = { id: "sheet-1", name: "Sheet1", load: () => undefined, getRange: rangeAt };
  (globalThis as any).Office = { context: { document: { url: "C:/books/a.xlsx" }, requirements: { isSetSupported: () => true } } };
  (globalThis as any).Excel = {
    run: async (fn: any) => fn({
      workbook: {
        application: { calculationMode: "automatic", load: () => undefined },
        worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet }
      },
      sync: async () => undefined
    })
  };
  return { cells, writes };
}

/** Ответ модели: вызовы инструментов одним шагом. */
function modelReplies(calls: { name: string; args: unknown }[]) {
  let requests = 0;
  const previous = globalThis.fetch;
  globalThis.fetch = async () => {
    requests += 1;
    const toolCalls = calls.map((call, index) => ({
      index, id: `call_${index}`, type: "function",
      function: { name: call.name, arguments: JSON.stringify(call.args) }
    }));
    const body = requests === 1
      ? { choices: [{ delta: { tool_calls: toolCalls }, finish_reason: "tool_calls" }] }
      : { choices: [{ delta: { content: "Готово." }, finish_reason: "stop" }] };
    return new Response(`data: ${JSON.stringify(body)}\n\ndata: [DONE]\n\n`, { status: 200 });
  };
  return { requests: () => requests, restore: () => { globalThis.fetch = previous; } };
}

/* --- частичный итог группы -------------------------------------------------- */

test("a group that stopped half way stops the task and is not shown as done", async (t) => {
  // План стабилизации, S4.5: группа возвращала ok: false и applied, а цикл
  // останавливал задачу по возвращённому итогу только при unknown — и шёл
  // дальше. В панели такая операция ещё и выглядела успешной.
  const excel = sheetExcel({ ignore: ["B1"] });
  const model = modelReplies([
    { name: "set_ranges_values", args: { writes: [
      { sheet: "Sheet1", address: "A1", values: [[1]] },
      { sheet: "Sheet1", address: "B1", values: [[2]] }
    ] } },
    // Следующая изменяющая команда того же шага не должна начаться.
    { name: "set_range_values", args: { sheet: "Sheet1", address: "C1", values: [[3]] } }
  ]);
  t.after(model.restore);

  const history: any[] = [{ role: "user", content: "Запиши" }];
  const events: any[] = [];
  const notices: string[] = [];
  await runAgent({
    provider: "deepseek", model: "test", history, initialContext, analysisOnly: false,
    hooks: {
      onDelta: () => undefined,
      onStepEnd: (text) => notices.push(text),
      onToolEvent: (event) => events.push(event),
      confirm: async () => true
    }
  });

  // B1 пишется дважды: панель один раз дописывает расхождение — это штатно.
  assert.ok(!excel.writes.includes("C1"), "C1 не записывалась");
  assert.equal(model.requests(), 1, "модель не спрашивали дальше: задача остановлена");
  const group = events.filter((event) => event.name === "set_ranges_values").at(-1);
  assert.notEqual(group.status, "done", "частичный итог не выглядит успехом");
  const reply = JSON.parse(history.find((message) => message.tool_call_id === "call_0").content);
  assert.equal(reply.ok, false, "модель видит, что группа не удалась");
  assert.equal(reply.executionState, "applied");
  assert.match(notices.at(-1) ?? "", /не повторяйте/i);
});

/* --- идентичность книги ------------------------------------------------------ */

test("no planned operation touches Excel when the workbook is not the one previewed", async () => {
  // План стабилизации, S4.2: идентичность книги сверял только исполнитель
  // одиночной записи. Тест перебирает весь реестр: план, подготовленный
  // для другой книги, не должен дойти до Excel ни у одного инструмента.
  let runs = 0;
  (globalThis as any).Office = { context: { document: { url: "C:/books/другая.xlsx" }, requirements: { isSetSupported: () => true } } };
  (globalThis as any).Excel = { run: async () => { runs += 1; throw new Error("до Excel дойти не должно"); } };

  const foreign = { workbookSessionId: "другая-сессия", documentUrl: "C:/books/a.xlsx", sheetId: "sheet-1", sheetName: "Sheet1" };
  for (const name of PLANNED_TOOLS) {
    const driver = planDriverFor(name)!;
    const plan: any = { kind: name, id: "p", target: foreign, workbook: foreign, items: [{ kind: "set_range_values", id: "i", target: foreign }] };
    runs = 0;
    let state: string | undefined;
    try {
      const result: any = await driver.execute(plan);
      state = result?.executionState;
    } catch (error: any) {
      state = error?.executionState;
    }
    assert.equal(state, "failed_before_write", `${name}: чужая книга — отказ до записи`);
    assert.equal(runs, 0, `${name}: до Excel не дошло`);
  }
});

/* --- остановка посреди группы --------------------------------------------------- */

test("stopping during a group finishes the started write and starts no more", async () => {
  // План стабилизации, S4.6: начатая запись получает честный итог,
  // следующие не начинаются.
  resetSnapshotStore();
  const stop = new AbortController();
  const excel = sheetExcel({ onWrite: (address) => { if (address === "A1") stop.abort(); } });
  const plan = await prepareSetRangesPlan({ writes: [
    { sheet: "Sheet1", address: "A1", values: [[1]] },
    { sheet: "Sheet1", address: "B1", values: [[2]] },
    { sheet: "Sheet1", address: "C1", values: [[3]] }
  ] });

  const result: any = await executeSetRangesPlan(plan, stop.signal);
  assert.deepEqual(excel.writes, ["A1"], "после остановки записей нет");
  assert.equal(result.ok, false);
  assert.equal(result.operations[0].executionState, "verified", "начатая запись доведена и проверена");
  assert.deepEqual(result.operations.slice(1).map((op: any) => op.executionState), ["not_started", "not_started"]);
  assert.match(result.operations[1].note, /остановил пользователь/i);
  assert.equal(snapshotStoreStats().pinned, 0, "закреплённые снимки освобождены");
});

/* --- снимки ------------------------------------------------------------------- */

test("pinned snapshots are released on success, refusal and failure", async (t) => {
  // План стабилизации, S4.7.
  for (const scenario of ["success", "refusal", "failure"] as const) {
    resetSnapshotStore();
    sheetExcel({ ignore: scenario === "failure" ? ["A1"] : [] });
    const model = modelReplies([{ name: "set_range_values", args: { sheet: "Sheet1", address: "A1", values: [[5]] } }]);
    await runAgent({
      provider: "deepseek", model: "test", history: [{ role: "user", content: "Запиши" }], initialContext, analysisOnly: false,
      hooks: {
        onDelta: () => undefined,
        onStepEnd: () => undefined,
        onToolEvent: () => undefined,
        confirm: async () => scenario !== "refusal"
      }
    });
    model.restore();
    assert.equal(snapshotStoreStats().pinned, 0, `${scenario}: закреплённых снимков не осталось`);
  }
  t.after(() => resetSnapshotStore());
});
