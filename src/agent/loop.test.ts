import test from "node:test";
import assert from "node:assert/strict";
import { cancellationToolMessages, runAgent } from "./loop";
import { executeSetRangePlan, prepareSetRangePlan, resolveToolArgs, runTool, ToolExecutionError, valuesForLiteralWrite } from "../excel/excelTools";

// Панель работает только внутри Excel: инструмент выдаётся, если Excel
// поддерживает его набор API (7.1.6). В тестах — Excel с ExcelApi 1.14,
// как на проверочной машине; тесты, которым нужен другой Excel, ставят свой.
(globalThis as any).Office ??= {
  context: { requirements: { isSetSupported: (_: string, version: string) => Number(version.split(".")[1]) <= 14 } }
};


const initialContext: any = {
  workbook: { sessionId: "test", documentUrl: "", identityConfirmed: true },
  activeSheet: { id: "sheet-1", name: "Sheet1" },
  activeCell: "Sheet1!A1",
  selectedAreas: ["Sheet1!A1"],
  readAt: "2026-09-15T00:00:00.000Z",
  capabilities: {}
};

test("cancelled tool calls are closed with matching tool_call_id", () => {
  const calls = [
    { id: "a", name: "set_range_values", arguments: "{}" },
    { id: "b", name: "insert_rows", arguments: "{}" }
  ];
  const messages = cancellationToolMessages(calls, 1);
  assert.equal(messages.length, 1);
  assert.equal(messages[0].role, "tool");
  assert.equal(messages[0].tool_call_id, "b");
  assert.match(String(messages[0].content), /отменена/i);
});

test("task sheet is stable unless the model explicitly names another sheet", async () => {
  const implicit = (await resolveToolArgs("set_range_values", { address: "A1", values: [[1]] }, "Sheet1")) as any;
  assert.equal(implicit.sheet, "Sheet1");

  const explicit = (await resolveToolArgs(
    "set_range_values",
    { sheet: "Sheet2", address: "A1", values: [[1]] },
    "Sheet1"
  )) as any;
  assert.equal(explicit.sheet, "Sheet2");
});

test("context tools with empty schemas do not receive an injected sheet", async () => {
  assert.deepEqual(await resolveToolArgs("get_active_context", {}, "Sheet1"), {});
  assert.deepEqual(await resolveToolArgs("list_sheets", {}, "Sheet1"), {});
});

function failingVerificationExcel(failureAt: number) {
  let syncCount = 0;
  let writeCount = 0;
  let currentValues: unknown[][] = [[0]];
  const range = {
    address: "Sheet1!A1",
    rowCount: 1,
    columnCount: 1,
    load: () => undefined,
    get values() { return currentValues; },
    set values(value: unknown[][]) { writeCount += 1; currentValues = value; },
    get formulas() { return currentValues; },
    set formulas(value: unknown[][]) { writeCount += 1; currentValues = value; }
  };
  const sheet = { id: "sheet-1", name: "Sheet1", load: () => undefined, getRange: () => range };
  const ctx = {
    workbook: {
      application: { calculationMode: "automatic", load: () => undefined },
      worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet }
    },
    sync: async () => {
      syncCount += 1;
      if (syncCount === failureAt) throw new Error("read failed");
    }
  };
  return { ctx, get writeCount() { return writeCount; } };
}

test("a successful write followed by a failed verification is reported as applied", async (t) => {
  const fake = failingVerificationExcel(5);
  const previousExcel = (globalThis as any).Excel;
  (globalThis as any).Excel = { run: async (fn: (ctx: unknown) => Promise<unknown>) => fn(fake.ctx) };
  t.after(() => { (globalThis as any).Excel = previousExcel; });

  await assert.rejects(
    runTool("set_range_values", { sheet: "Sheet1", address: "A1", values: [[42]] }),
    (error: unknown) => error instanceof ToolExecutionError && error.executionState === "applied"
  );
  assert.equal(fake.writeCount, 1);
});

test("agent stops after uncertain write and closes later tool calls", async (t) => {
  const fake = failingVerificationExcel(5);
  const previousExcel = (globalThis as any).Excel;
  const previousFetch = globalThis.fetch;
  (globalThis as any).Excel = { run: async (fn: (ctx: unknown) => Promise<unknown>) => fn(fake.ctx) };
  let fetchCount = 0;
  globalThis.fetch = async () => {
    fetchCount += 1;
    const calls = [
      { index: 0, id: "write_one", type: "function", function: { name: "set_range_values", arguments: '{"sheet":"Sheet1","address":"A1","values":[[42]]}' } },
      { index: 1, id: "write_two", type: "function", function: { name: "set_range_values", arguments: '{"sheet":"Sheet1","address":"B1","values":[[43]]}' } }
    ];
    const sse = `data: ${JSON.stringify({ choices: [{ delta: { tool_calls: calls }, finish_reason: "tool_calls" }] })}\n\ndata: [DONE]\n\n`;
    return new Response(sse, { status: 200 });
  };
  t.after(() => { (globalThis as any).Excel = previousExcel; globalThis.fetch = previousFetch; });

  const history: any[] = [{ role: "user", content: "Запиши числа" }];
  const notices: string[] = [];
  await runAgent({
    provider: "deepseek", model: "test", history, initialContext,
    hooks: {
      onDelta: () => undefined,
      onStepEnd: (text) => notices.push(text),
      onToolEvent: () => undefined,
      confirm: async () => true
    }
  });

  assert.equal(fake.writeCount, 1);
  assert.equal(fetchCount, 1);
  assert.equal(history.at(-2)?.tool_call_id, "write_one");
  assert.equal(JSON.parse(history.at(-2)?.content).executionState, "applied");
  assert.equal(history.at(-1)?.tool_call_id, "write_two");
  assert.match(notices.at(-1) ?? "", /не повторяйте/i);
});

test("all calls in one model response count toward the read budget", async (t) => {
  const previousExcel = (globalThis as any).Excel;
  const previousFetch = globalThis.fetch;
  let reads = 0;
  const range = {
    address: "Sheet1!A1", rowCount: 1, columnCount: 1,
    values: [[1]], formulas: [[1]], numberFormat: [["General"]],
    load: () => undefined
  };
  const sheet = { id: "sheet-1", name: "Sheet1", load: () => undefined, getRange: () => { reads += 1; return range; } };
  const ctx = {
    workbook: {
      application: { calculationMode: "automatic", load: () => undefined },
      worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet }
    },
    sync: async () => undefined
  };
  (globalThis as any).Excel = { run: async (fn: (context: unknown) => Promise<unknown>) => fn(ctx) };
  globalThis.fetch = async () => {
    const calls = Array.from({ length: 31 }, (_, index) => ({
      index, id: `read_${index}`, type: "function",
      function: { name: "get_range_values", arguments: '{"sheet":"Sheet1","address":"A1"}' }
    }));
    return new Response(`data: ${JSON.stringify({ choices: [{ delta: { tool_calls: calls }, finish_reason: "tool_calls" }] })}\n\ndata: [DONE]\n\n`);
  };
  t.after(() => { (globalThis as any).Excel = previousExcel; globalThis.fetch = previousFetch; });

  const history: any[] = [{ role: "user", content: "Прочитай много раз" }];
  const notices: string[] = [];
  await runAgent({
    provider: "deepseek", model: "test", history, initialContext,
    hooks: {
      onDelta: () => undefined,
      onStepEnd: (text) => notices.push(text),
      onToolEvent: () => undefined,
      confirm: async () => true
    }
  });
  assert.equal(reads, 30);
  assert.equal(history.filter((message) => message.role === "tool").length, 31);
  assert.match(notices.at(-1) ?? "", /лимит/i);
});

test("invalid address is rejected before asking for confirmation", async (t) => {
  const previousExcel = (globalThis as any).Excel;
  const previousFetch = globalThis.fetch;
  const sheet = { name: "Sheet1", load: () => undefined };
  const ctx = { workbook: { worksheets: { getActiveWorksheet: () => sheet } }, sync: async () => undefined };
  (globalThis as any).Excel = { run: async (fn: (context: unknown) => Promise<unknown>) => fn(ctx) };
  globalThis.fetch = async () => {
    const calls = [{ index: 0, id: "invalid", type: "function", function: {
      name: "set_range_values", arguments: '{"sheet":"Sheet1","address":"not-a-range","values":[[1]]}'
    } }];
    return new Response(`data: ${JSON.stringify({ choices: [{ delta: { tool_calls: calls }, finish_reason: "tool_calls" }] })}\n\ndata: [DONE]\n\n`);
  };
  t.after(() => { (globalThis as any).Excel = previousExcel; globalThis.fetch = previousFetch; });

  let confirmations = 0;
  const history: any[] = [{ role: "user", content: "Запиши" }];
  // The model repeats the invalid call until the response budget is reached;
  // none of those attempts should reach the approval dialog or Excel writer.
  await runAgent({
    provider: "deepseek", model: "test", history, initialContext,
    hooks: {
      onDelta: () => undefined, onStepEnd: () => undefined, onToolEvent: () => undefined,
      confirm: async () => { confirmations += 1; return true; }
    }
  });
  assert.equal(confirmations, 0);
  assert.equal(JSON.parse(history[2].content).executionState, "failed_before_write");
});

test("analysis-only mode rejects a mutating call in the executor", async (t) => {
  const previousFetch = globalThis.fetch;
  let fetchCount = 0;
  globalThis.fetch = async () => {
    fetchCount += 1;
    if (fetchCount === 1) {
      const calls = [{ index: 0, id: "forbidden", type: "function", function: {
        name: "set_range_values", arguments: '{"sheet":"Sheet1","address":"A1","values":[[9]]}'
      } }];
      return new Response(`data: ${JSON.stringify({ choices: [{ delta: { tool_calls: calls }, finish_reason: "tool_calls" }] })}\n\ndata: [DONE]\n\n`);
    }
    return new Response(`data: ${JSON.stringify({ choices: [{ delta: { content: "Запись запрещена." }, finish_reason: "stop" }] })}\n\ndata: [DONE]\n\n`);
  };
  t.after(() => { globalThis.fetch = previousFetch; });
  let confirmations = 0;
  const history: any[] = [{ role: "user", content: "Запиши" }];
  await runAgent({
    provider: "deepseek", model: "test", history, initialContext, analysisOnly: true,
    hooks: {
      onDelta: () => undefined, onStepEnd: () => undefined, onToolEvent: () => undefined,
      confirm: async () => { confirmations += 1; return true; }
    }
  });
  assert.equal(confirmations, 0);
  assert.match(history.find((message) => message.role === "tool")?.content ?? "", /Только анализ/);
});

test("set-range plan is immutable and detects a manual edit before writing", async (t) => {
  let current: unknown[][] = [[1]];
  let writes = 0;
  const range = {
    address: "Sheet1!A1", rowCount: 1, columnCount: 1,
    get values() { return current; },
    set values(value: unknown[][]) { writes += 1; current = value; },
    get formulas() { return current; },
    set formulas(value: unknown[][]) { writes += 1; current = value; },
    load: () => undefined
  };
  const sheet = { id: "sheet-1", name: "Sheet1", load: () => undefined, getRange: () => range };
  const ctx = {
    workbook: {
      application: { calculationMode: "automatic", load: () => undefined },
      worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet }
    },
    sync: async () => undefined
  };
  const previousExcel = (globalThis as any).Excel;
  (globalThis as any).Excel = { run: async (fn: (context: unknown) => Promise<unknown>) => fn(ctx) };
  t.after(() => { (globalThis as any).Excel = previousExcel; });

  const plan = await prepareSetRangePlan({ sheet: "Sheet1", address: "A1", values: [[2]] });
  assert.equal(Object.isFrozen(plan), true);
  assert.equal(Object.isFrozen(plan.before), true);
  current = [[7]];
  await assert.rejects(
    () => executeSetRangePlan(plan),
    (error: unknown) => error instanceof ToolExecutionError && error.executionState === "failed_before_write"
  );
  assert.equal(writes, 0);
});

test("set-range plan follows the sheet id after rename instead of its old name", async (t) => {
  let current: unknown[][] = [[1]];
  let requestedItems: string[] = [];
  const sheet = {
    id: "sheet-stable-id",
    name: "BeforeRename",
    load: () => undefined,
    getRange: () => range
  };
  const range = {
    get address() { return `${sheet.name}!A1`; },
    rowCount: 1,
    columnCount: 1,
    get values() { return current; },
    set values(value: unknown[][]) { current = value; },
    get formulas() { return current; },
    set formulas(value: unknown[][]) { current = value; },
    load: () => undefined
  };
  const ctx = {
    workbook: {
      application: { calculationMode: "automatic", load: () => undefined },
      worksheets: {
        getActiveWorksheet: () => sheet,
        getItem: (key: string) => { requestedItems.push(key); return sheet; }
      }
    },
    sync: async () => undefined
  };
  const previousExcel = (globalThis as any).Excel;
  const previousOffice = (globalThis as any).Office;
  (globalThis as any).Office = { context: { document: { url: "C:/books/a.xlsx" } } };
  (globalThis as any).Excel = { run: async (fn: (context: unknown) => Promise<unknown>) => fn(ctx) };
  t.after(() => {
    (globalThis as any).Excel = previousExcel;
    (globalThis as any).Office = previousOffice;
  });

  const plan = await prepareSetRangePlan({ sheet: "BeforeRename", address: "A1", values: [["=1+1"]], isFormula: true });
  sheet.name = "AfterRename";
  requestedItems = [];
  const result = await executeSetRangePlan(plan) as { sheet: string; executionState: string };
  assert.equal(result.sheet, "AfterRename");
  assert.equal(result.executionState, "verified");
  assert.equal(requestedItems[0], "sheet-stable-id", "execution must resolve the immutable sheet id");
});

test("set-range plan refuses execution after the workbook changes", async (t) => {
  let writes = 0;
  let current: unknown[][] = [[1]];
  const range = {
    address: "Sheet1!A1", rowCount: 1, columnCount: 1,
    get values() { return current; },
    set values(value: unknown[][]) { writes += 1; current = value; },
    get formulas() { return current; },
    set formulas(value: unknown[][]) { writes += 1; current = value; },
    load: () => undefined
  };
  const sheet = { id: "sheet-1", name: "Sheet1", load: () => undefined, getRange: () => range };
  const ctx = {
    workbook: {
      application: { calculationMode: "automatic", load: () => undefined },
      worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet }
    },
    sync: async () => undefined
  };
  const previousExcel = (globalThis as any).Excel;
  const previousOffice = (globalThis as any).Office;
  const document = { url: "C:/books/a.xlsx" };
  (globalThis as any).Office = { context: { document } };
  (globalThis as any).Excel = { run: async (fn: (context: unknown) => Promise<unknown>) => fn(ctx) };
  t.after(() => {
    (globalThis as any).Excel = previousExcel;
    (globalThis as any).Office = previousOffice;
  });

  const plan = await prepareSetRangePlan({ sheet: "Sheet1", address: "A1", values: [[2]] });
  document.url = "C:/books/b.xlsx";
  await assert.rejects(() => executeSetRangePlan(plan), /книга изменилась/i);
  assert.equal(writes, 0);
});

test("literal text beginning with equals is escaped before Excel assignment", () => {
  assert.deepEqual(valuesForLiteralWrite([["=SUM(1,2)", "text", 3]]), [["'=SUM(1,2)", "'text", 3]]);
  // Текст, похожий на число или дату, остаётся текстом; пустая строка очищает ячейку.
  assert.deepEqual(valuesForLiteralWrite([["00123", "04.09.2026", "", true, null]]), [["'00123", "'04.09.2026", "", true, null]]);
});

/** Макет с поддержкой опроса объединений. Excel на замеренной сборке отдаёт
 * объединение одним углом без границ, поэтому угол здесь — одна ячейка. */
function excelWithMergeAnchor(anchorAddress: string | null) {
  const range = {
    address: "Sheet1!B2",
    rowCount: 1,
    columnCount: 1,
    rowIndex: 1,
    columnIndex: 1,
    load: () => undefined,
    values: [[0]],
    formulas: [[0]],
    getMergedAreasOrNullObject: () => merged
  };
  const merged = {
    isNullObject: anchorAddress === null,
    address: anchorAddress ?? "",
    areaCount: anchorAddress ? 1 : 0,
    areas: { items: [] as { address: string }[], load: () => undefined },
    load: () => undefined
  };
  const probe = {
    address: "Sheet1!A1:V22",
    load: () => undefined,
    getMergedAreasOrNullObject: () => merged
  };
  const sheet = {
    id: "sheet-1",
    name: "Sheet1",
    load: () => undefined,
    getRange: () => range,
    getRangeByIndexes: () => probe
  };
  return {
    workbook: {
      application: { calculationMode: "automatic", load: () => undefined },
      worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet }
    },
    sync: async () => undefined
  };
}

test("a write plan warns when an unresolved merge anchor may cover the target", async () => {
  (globalThis as any).Excel = { run: async (fn: any) => fn(excelWithMergeAnchor("Sheet1!A1")) };
  const plan = (await prepareSetRangePlan({ sheet: "Sheet1", address: "B2", values: [[7]] })) as any;

  // Угол A1 стоит выше и левее цели B2, значит может её накрывать.
  assert.deepEqual(plan.mergedAnchorsUnresolved, ["Sheet1!A1"]);
  assert.match(plan.mergeWarning, /Sheet1!A1/);
  assert.match(plan.mergeWarning, /ведёт себя не так, как в обычную/);
  assert.match(plan.mergeWarning, /перед подтверждением/);
});

test("a write plan stays quiet when no anchor can reach the target", async () => {
  (globalThis as any).Excel = { run: async (fn: any) => fn(excelWithMergeAnchor("Sheet1!D9")) };
  const plan = (await prepareSetRangePlan({ sheet: "Sheet1", address: "B2", values: [[7]] })) as any;

  // Угол D9 правее и ниже B2: объединение влево и вверх не растёт.
  assert.equal(plan.mergedAnchorsUnresolved, undefined);
  assert.equal(plan.mergeWarning, undefined);
});

test("a write plan survives builds without merge probing", async () => {
  const excel = excelWithMergeAnchor(null) as any;
  delete excel.workbook.worksheets.getItem().getRangeByIndexes;
  (globalThis as any).Excel = { run: async (fn: any) => fn(excel) };

  // Опрос вспомогательный: его отсутствие не должно ронять подготовку записи.
  const plan = (await prepareSetRangePlan({ sheet: "Sheet1", address: "B2", values: [[7]] })) as any;
  assert.equal(plan.cellCount, 1);
  assert.equal(plan.mergeWarning, undefined);
});

test("the same write twice in one step is refused, reads are not", async () => {
  const { duplicateMutatingCall } = await import("./loop");
  const seen = new Set<string>();
  const write = { id: "a", name: "set_range_values", arguments: '{"address":"A1","values":[[1]]}' };

  assert.equal(duplicateMutatingCall(write, seen), false, "первый вызов проходит");
  assert.equal(duplicateMutatingCall({ ...write, id: "b" }, seen), true, "тот же вызов повторно — нет");
  // Другие аргументы — другая операция.
  assert.equal(duplicateMutatingCall({ id: "c", name: "set_range_values", arguments: '{"address":"A2","values":[[1]]}' }, seen), false);
  // Повторное чтение безвредно и иногда осмысленно.
  const read = { id: "d", name: "get_range_values", arguments: '{"address":"A1"}' };
  assert.equal(duplicateMutatingCall(read, seen), false);
  assert.equal(duplicateMutatingCall({ ...read, id: "e" }, seen), false);
});

test("broken arguments do not break duplicate detection", async () => {
  const { duplicateMutatingCall } = await import("./loop");
  const seen = new Set<string>();
  const broken = { id: "a", name: "set_range_values", arguments: "{не json" };
  assert.equal(duplicateMutatingCall(broken, seen), false);
  assert.equal(duplicateMutatingCall({ ...broken, id: "b" }, seen), true);
});
