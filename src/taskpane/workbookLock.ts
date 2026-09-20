/**
 * Один владелец изменений книги.
 *
 * Панель защищалась состоянием React: пока идёт задача, поле ввода и кнопки
 * выключены. Но состояние обновляется не мгновенно, а между нажатием и первым
 * `await` в обработчике успевает пройти второе нажатие — тогда запускаются два
 * цикла изменения книги сразу. То же и с отменой: она шла отдельно от задачи
 * и могла выполниться посреди записи.
 *
 * Замок захватывается синхронно, до первого `await`, и освобождается в
 * `finally`. Владелец один: задача и отмена исключают друг друга.
 */

export type LockOwner = "task" | "undo";

export interface WorkbookLock {
  /** Захват без ожидания: null — занято кем-то другим. */
  tryAcquire(owner: LockOwner): (() => void) | null;
  owner(): LockOwner | null;
  /** Подписка на смену владельца — для перерисовки панели. */
  subscribe(listener: (owner: LockOwner | null) => void): () => void;
}

export function createWorkbookLock(): WorkbookLock {
  let owner: LockOwner | null = null;
  let ticket = 0;
  const listeners = new Set<(owner: LockOwner | null) => void>();
  const announce = () => { for (const listener of listeners) listener(owner); };

  return {
    tryAcquire(next) {
      if (owner !== null) return null;
      owner = next;
      const mine = ++ticket;
      announce();
      // Освобождение по своему билету: повторный вызов и опоздавший
      // обработчик прежней операции не снимут чужой замок.
      return () => {
        if (owner === null || mine !== ticket) return;
        owner = null;
        announce();
      };
    },
    owner: () => owner,
    subscribe(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    }
  };
}

/** Признак того, что работа не запускалась: замком владеет кто-то другой. */
export const BUSY = Symbol("панель уже выполняет операцию");

/**
 * Выполняет работу под замком. Захват синхронный: до первого `await`
 * внутри этой функции ничего не ждём, иначе второе нажатие успеет пройти.
 */
export function withWorkbookLock<T>(
  lock: WorkbookLock,
  owner: LockOwner,
  work: () => Promise<T>
): Promise<T> | typeof BUSY {
  const release = lock.tryAcquire(owner);
  if (!release) return BUSY;
  let started: Promise<T>;
  try {
    started = work();
  } catch (error) {
    // Синхронный отказ работы — тоже её конец.
    release();
    throw error;
  }
  return started.finally(release);
}

/** Понятное объяснение отказа: чем именно занята панель. */
export function busyReason(owner: LockOwner | null): string {
  if (owner === "task") return "Панель уже выполняет задачу. Дождитесь её конца или нажмите «Остановить».";
  if (owner === "undo") return "Идёт отмена предыдущей операции. Дождитесь её конца.";
  return "Панель занята.";
}
