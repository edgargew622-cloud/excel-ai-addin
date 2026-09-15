import type { RevisionCoverage, WorkbookRevision } from "./workbookRevision";

export type SnapshotKind = "content" | "details" | "search" | "plan";
export type SnapshotState = "fresh" | "stale" | "evicted" | "missing";

export interface SnapshotIdentity {
  workbookSessionId: string;
  documentUrl: string;
}

export interface WorkbookSnapshot<T = unknown> {
  id: string;
  kind: SnapshotKind;
  source: string;
  workbook: SnapshotIdentity;
  sheetId?: string;
  sheetName?: string;
  address?: string;
  capturedAt: string;
  revision: WorkbookRevision;
  coverage: RevisionCoverage;
  payload: T;
  bytes: number;
  pinned: boolean;
}

export interface SnapshotRecall<T = unknown> {
  id: string;
  state: SnapshotState;
  historical: boolean;
  reason?: string;
  snapshot?: Omit<WorkbookSnapshot<T>, "pinned">;
}

export const MAX_SNAPSHOT_STORE_BYTES = 2 * 1024 * 1024;
export const MAX_SNAPSHOTS = 50;
const MAX_TOMBSTONES = 100;

const snapshots = new Map<string, WorkbookSnapshot>();
const order: string[] = [];
const tombstones = new Map<string, string>();
let totalBytes = 0;

const byteLength = (value: string) => new TextEncoder().encode(value).length;
const id = () => globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`;

function rememberEviction(snapshotId: string, reason: string) {
  tombstones.set(snapshotId, reason);
  while (tombstones.size > MAX_TOMBSTONES) tombstones.delete(tombstones.keys().next().value as string);
}

function remove(snapshotId: string, reason: string) {
  const snapshot = snapshots.get(snapshotId);
  if (!snapshot) return;
  snapshots.delete(snapshotId);
  totalBytes -= snapshot.bytes;
  const index = order.indexOf(snapshotId);
  if (index >= 0) order.splice(index, 1);
  rememberEviction(snapshotId, reason);
}

function evictFor(incomingBytes: number) {
  while (order.length && (totalBytes + incomingBytes > MAX_SNAPSHOT_STORE_BYTES || snapshots.size >= MAX_SNAPSHOTS)) {
    const candidate = order.find((snapshotId) => !snapshots.get(snapshotId)?.pinned);
    if (!candidate) break;
    remove(candidate, "Снимок вытеснен ограничением памяти.");
  }
}

export function recordSnapshot<T>(input: {
  kind: SnapshotKind;
  source: string;
  workbook: SnapshotIdentity;
  sheetId?: string;
  sheetName?: string;
  address?: string;
  revision: WorkbookRevision;
  coverage: RevisionCoverage;
  payload: T;
  pin?: boolean;
}): WorkbookSnapshot<T> | null {
  const snapshotId = id();
  const { pin = false, ...data } = input;
  const base = { ...data, id: snapshotId, capturedAt: new Date().toISOString(), pinned: pin };
  let bytes = byteLength(JSON.stringify({ ...base, bytes: 0 }));
  bytes = byteLength(JSON.stringify({ ...base, bytes }));
  if (bytes > MAX_SNAPSHOT_STORE_BYTES) {
    rememberEviction(snapshotId, "Снимок превышает общий лимит памяти и не был сохранён.");
    return null;
  }
  evictFor(bytes);
  if (totalBytes + bytes > MAX_SNAPSHOT_STORE_BYTES || snapshots.size >= MAX_SNAPSHOTS) {
    rememberEviction(snapshotId, "Снимок не сохранён: всю доступную память удерживают активные планы.");
    return null;
  }
  const snapshot: WorkbookSnapshot<T> = { ...base, bytes };
  snapshots.set(snapshotId, snapshot);
  order.push(snapshotId);
  totalBytes += bytes;
  return snapshot;
}

export function setSnapshotPinned(snapshotId: string, pinned: boolean): boolean {
  const snapshot = snapshots.get(snapshotId);
  if (!snapshot) return false;
  snapshot.pinned = pinned;
  return true;
}

function coverageGap(kind: SnapshotKind, coverage: RevisionCoverage): string | null {
  if (!coverage.structure) return "События структуры книги не отслеживались полностью.";
  if ((kind === "content" || kind === "search" || kind === "plan") && !coverage.content) {
    return "Изменения значений и формул не отслеживались полностью.";
  }
  if (kind === "details" && (!coverage.format || !coverage.protection)) {
    return "Изменения оформления или защиты не отслеживались полностью.";
  }
  if (!coverage.sheetNames) return "Переименование листов не отслеживалось этой версией Excel.";
  return null;
}

export function recallSnapshot<T>(
  snapshotId: string,
  workbook: SnapshotIdentity,
  revision: WorkbookRevision,
  currentCoverage: RevisionCoverage
): SnapshotRecall<T> {
  const snapshot = snapshots.get(snapshotId) as WorkbookSnapshot<T> | undefined;
  if (!snapshot) {
    const reason = tombstones.get(snapshotId);
    return reason
      ? { id: snapshotId, state: "evicted", historical: true, reason }
      : { id: snapshotId, state: "missing", historical: true, reason: "Снимок не найден в этой сессии панели." };
  }
  const publicSnapshot = { ...snapshot };
  delete (publicSnapshot as Partial<WorkbookSnapshot>).pinned;
  const mismatch = snapshot.workbook.workbookSessionId !== workbook.workbookSessionId ||
    snapshot.workbook.documentUrl !== workbook.documentUrl;
  const gap = coverageGap(snapshot.kind, snapshot.coverage) ?? coverageGap(snapshot.kind, currentCoverage);
  const changed = snapshot.kind === "details"
    ? snapshot.revision.structure !== revision.structure || snapshot.revision.format !== revision.format
    : snapshot.revision.structure !== revision.structure || snapshot.revision.content !== revision.content;
  const reason = mismatch
    ? "Снимок относится к другой книге или сессии панели."
    : gap ?? (changed ? "Книга изменилась после создания снимка." : undefined);
  return {
    id: snapshotId,
    state: reason ? "stale" : "fresh",
    historical: Boolean(reason),
    ...(reason ? { reason } : {}),
    snapshot: publicSnapshot as Omit<WorkbookSnapshot<T>, "pinned">
  };
}

export function snapshotStoreStats() {
  return { count: snapshots.size, bytes: totalBytes, pinned: [...snapshots.values()].filter((item) => item.pinned).length };
}

export function resetSnapshotStore() {
  snapshots.clear();
  order.length = 0;
  tombstones.clear();
  totalBytes = 0;
}
