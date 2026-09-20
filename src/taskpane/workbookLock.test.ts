import test from "node:test";
import assert from "node:assert/strict";
import { BUSY, busyReason, createWorkbookLock, withWorkbookLock } from "./workbookLock";

/** Работа, которая ждёт, пока её не отпустят: так ловятся два нажатия подряд. */
function pending<T>(value: T) {
  let finish: () => void;
  const gate = new Promise<void>((resolve) => { finish = resolve; });
  return {
    work: async () => { await gate; return value; },
    finish: () => finish!()
  };
}

test("two clicks in one tick start one task, not two", async () => {
  const lock = createWorkbookLock();
  const first = pending("сделано");
  let runs = 0;
  const run = () => withWorkbookLock(lock, "task", () => { runs += 1; return first.work(); });

  // Оба нажатия проходят до первого await: React перерисовать панель не успел.
  const a = run();
  const b = run();
  assert.notEqual(a, BUSY, "первое нажатие запускает задачу");
  assert.equal(b, BUSY, "второе — нет");
  assert.equal(runs, 1);

  first.finish();
  assert.equal(await (a as Promise<string>), "сделано");
  assert.equal(lock.owner(), null, "после конца замок свободен");
  assert.notEqual(run(), BUSY, "новую задачу запустить можно");
});

test("undo and a task exclude each other in both directions", async () => {
  const lock = createWorkbookLock();
  const task = pending("задача");
  const started = withWorkbookLock(lock, "task", task.work);
  assert.equal(withWorkbookLock(lock, "undo", async () => "отмена"), BUSY, "отмена ждёт задачу");
  assert.match(busyReason(lock.owner()), /выполняет задачу/);

  task.finish();
  await started;

  const undoing = pending("отмена");
  const undoStarted = withWorkbookLock(lock, "undo", undoing.work);
  assert.equal(withWorkbookLock(lock, "task", async () => "задача"), BUSY, "задача ждёт отмену");
  assert.match(busyReason(lock.owner()), /отмена предыдущей операции/);
  undoing.finish();
  await undoStarted;
  assert.equal(lock.owner(), null);
});

test("a failure releases the lock, both async and synchronous", async () => {
  const lock = createWorkbookLock();
  const failed = withWorkbookLock(lock, "task", async () => { throw new Error("сбой сети"); });
  await assert.rejects(() => failed as Promise<unknown>, /сбой сети/);
  assert.equal(lock.owner(), null, "панель не остаётся заблокированной навсегда");

  assert.throws(() => withWorkbookLock(lock, "task", () => { throw new Error("сразу"); }), /сразу/);
  assert.equal(lock.owner(), null);
});

test("a late release from a finished task does not unlock the next one", async () => {
  const lock = createWorkbookLock();
  const stale = lock.tryAcquire("task")!;
  stale();
  const undoRelease = lock.tryAcquire("undo")!;
  // Опоздавший обработчик прежней задачи не должен снимать чужой замок.
  stale();
  assert.equal(lock.owner(), "undo");
  undoRelease();
  assert.equal(lock.owner(), null);
});

test("the panel is told who holds the lock, so it can redraw", () => {
  const lock = createWorkbookLock();
  const seen: (string | null)[] = [];
  const unsubscribe = lock.subscribe((owner) => seen.push(owner));
  const release = lock.tryAcquire("task")!;
  release();
  unsubscribe();
  lock.tryAcquire("undo");
  assert.deepEqual(seen, ["task", null], "после отписки уведомлений нет");
});

test("the panel really goes through this lock, not through a flag of its own", async () => {
  // План требует проверять настоящий путь панели: тест замка, которым панель
  // не пользуется, ничего не доказывает.
  const source = await import("node:fs/promises").then((fs) =>
    fs.readFile(new URL("./Taskpane.tsx", import.meta.url), "utf8"));

  assert.match(source, /withWorkbookLock\(lock\.current, "task"/, "отправка идёт через замок");
  assert.match(source, /withWorkbookLock\(lock\.current, "undo"/, "отмена идёт через замок");
  assert.doesNotMatch(source, /setBusy\(/, "прежнего состояния busy больше нет");
  assert.match(source, /const busy = lockOwner !== null/, "интерфейс блокируется по владельцу замка");
});
