import test from "node:test";
import assert from "node:assert/strict";
import { createTaskBudget, stopsTask, type BudgetLimits } from "./budget";

const limits: BudgetLimits = { modelSteps: 3, readCalls: 2, writeCalls: 1, taskMs: 1000 };

test("read and write budgets are independent", () => {
  const budget = createTaskBudget(limits, () => 0);

  // Разведка не должна отнимать бюджет у правки: два чтения исчерпывают свой
  // счётчик, но запись после них всё равно разрешена.
  assert.equal(budget.tryToolCall(false).ok, true);
  assert.equal(budget.tryToolCall(false).ok, true);
  const thirdRead = budget.tryToolCall(false);
  assert.equal(thirdRead.ok, false);
  assert.equal(thirdRead.ok === false && thirdRead.kind, "readCalls");

  assert.equal(budget.tryToolCall(true).ok, true);
});

test("write budget is separate and small", () => {
  const budget = createTaskBudget(limits, () => 0);
  assert.equal(budget.tryToolCall(true).ok, true);
  const second = budget.tryToolCall(true);
  assert.equal(second.ok, false);
  assert.equal(second.ok === false && second.kind, "writeCalls");
  // Чтения при исчерпанном бюджете записи остаются доступны: агент должен
  // иметь возможность перечитать книгу и отчитаться о сделанном.
  assert.equal(budget.tryToolCall(false).ok, true);
});

test("model steps are counted and then refused", () => {
  const budget = createTaskBudget(limits, () => 0);
  for (let i = 0; i < limits.modelSteps; i++) {
    assert.equal(budget.tryModelStep().ok, true, `step ${i}`);
  }
  const over = budget.tryModelStep();
  assert.equal(over.ok, false);
  assert.equal(over.ok === false && over.kind, "modelSteps");
  assert.equal(budget.used().modelSteps, limits.modelSteps);
});

test("time limit refuses both model steps and tool calls", () => {
  let clock = 0;
  const budget = createTaskBudget(limits, () => clock);
  assert.equal(budget.tryModelStep().ok, true);

  clock = limits.taskMs;
  const step = budget.tryModelStep();
  assert.equal(step.ok, false);
  assert.equal(step.ok === false && step.kind, "taskMs");

  const call = budget.tryToolCall(false);
  assert.equal(call.ok, false);
  assert.equal(call.ok === false && call.kind, "taskMs");
});

test("time exhaustion does not consume the tool counters", () => {
  let clock = 0;
  const budget = createTaskBudget(limits, () => clock);
  clock = limits.taskMs;
  budget.tryToolCall(false);
  budget.tryToolCall(true);
  // Отказ по времени не должен списывать чтения и записи: иначе счётчики
  // в отчёте покажут работу, которой не было.
  assert.equal(budget.used().readCalls, 0);
  assert.equal(budget.used().writeCalls, 0);
});

test("only step and time exhaustion stop the whole task", () => {
  assert.equal(stopsTask("modelSteps"), true);
  assert.equal(stopsTask("taskMs"), true);
  // Исчерпание чтений или записей возвращается модели как результат
  // инструмента: она должна успеть отчитаться, а не оборваться на середине.
  assert.equal(stopsTask("readCalls"), false);
  assert.equal(stopsTask("writeCalls"), false);
});

test("refusal text tells the model what to do next", () => {
  const budget = createTaskBudget({ ...limits, writeCalls: 0 }, () => 0);
  const denied = budget.tryToolCall(true);
  assert.equal(denied.ok, false);
  if (denied.ok === false) {
    assert.match(denied.reason, /не вызывая больше изменяющих/i);
  }
});
