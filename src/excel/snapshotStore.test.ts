import test from "node:test";
import assert from "node:assert/strict";
import {
  MAX_SNAPSHOTS,
  recallSnapshot,
  recordSnapshot,
  resetSnapshotStore,
  setSnapshotPinned,
  snapshotStoreStats
} from "./snapshotStore";
import { bumpWorkbookRevision, getRevisionCoverage, getWorkbookRevision, resetWorkbookRevision, setRevisionCoverage } from "./workbookRevision";

const workbook = { workbookSessionId: "session-a", documentUrl: "C:/book.xlsx" };

function freshStore() {
  resetSnapshotStore();
  resetWorkbookRevision();
  setRevisionCoverage({ content: true, structure: true, format: true, protection: true, sheetNames: true });
}

test("snapshot becomes stale after the relevant workbook revision", () => {
  freshStore();
  const snapshot = recordSnapshot({
    kind: "content", source: "get_range_values", workbook, revision: getWorkbookRevision(),
    coverage: { content: true, structure: true, format: true, protection: true, sheetNames: true }, payload: [[1]]
  });
  assert.ok(snapshot);
  assert.equal(recallSnapshot(snapshot.id, workbook, getWorkbookRevision(), getRevisionCoverage()).state, "fresh");
  bumpWorkbookRevision("format");
  assert.equal(recallSnapshot(snapshot.id, workbook, getWorkbookRevision(), getRevisionCoverage()).state, "fresh");
  bumpWorkbookRevision("content");
  assert.equal(recallSnapshot(snapshot.id, workbook, getWorkbookRevision(), getRevisionCoverage()).state, "stale");
});

test("incomplete event coverage never presents a snapshot as fresh", () => {
  freshStore();
  const snapshot = recordSnapshot({
    kind: "details", source: "get_range_details", workbook, revision: getWorkbookRevision(),
    coverage: { content: true, structure: true, format: true, protection: false, sheetNames: true }, payload: {}
  });
  assert.ok(snapshot);
  const recalled = recallSnapshot(snapshot.id, workbook, getWorkbookRevision(), getRevisionCoverage());
  assert.equal(recalled.state, "stale");
  assert.match(recalled.reason ?? "", /защит/i);
});

test("pinned plan survives eviction while old unpinned snapshots get tombstones", () => {
  freshStore();
  const pinned = recordSnapshot({
    kind: "plan", source: "set_range_values:before", workbook, revision: getWorkbookRevision(),
    coverage: { content: true, structure: true, format: true, protection: true, sheetNames: true }, payload: [[1]], pin: true
  });
  assert.ok(pinned);
  let firstUnpinned = "";
  for (let index = 0; index < MAX_SNAPSHOTS + 5; index++) {
    const item = recordSnapshot({
      kind: "content", source: "read", workbook, revision: getWorkbookRevision(),
      coverage: { content: true, structure: true, format: true, protection: true, sheetNames: true }, payload: index
    });
    if (!firstUnpinned && item) firstUnpinned = item.id;
  }
  assert.equal(snapshotStoreStats().count, MAX_SNAPSHOTS);
  assert.equal(recallSnapshot(pinned.id, workbook, getWorkbookRevision(), getRevisionCoverage()).state, "fresh");
  assert.equal(recallSnapshot(firstUnpinned, workbook, getWorkbookRevision(), getRevisionCoverage()).state, "evicted");
  assert.equal(setSnapshotPinned(pinned.id, false), true);
});

test("snapshot becomes stale when event monitoring is lost after capture", () => {
  freshStore();
  const snapshot = recordSnapshot({
    kind: "content", source: "read", workbook, revision: getWorkbookRevision(),
    coverage: getRevisionCoverage(), payload: [[1]]
  });
  assert.ok(snapshot);
  setRevisionCoverage({ content: false });
  const recalled = recallSnapshot(snapshot.id, workbook, getWorkbookRevision(), getRevisionCoverage());
  assert.equal(recalled.state, "stale");
  assert.match(recalled.reason ?? "", /значений/i);
});

test("untracked sheet renames never block freshness on Excel below 1.17", () => {
  freshStore();
  // Событие onNameChanged требует ExcelApi 1.17. На замеренной среде потолок
  // 1.14, поэтому sheetNames здесь ложно — как на реальной машине.
  const coverage = { content: true, structure: true, format: true, protection: true, sheetNames: false };
  setRevisionCoverage(coverage);

  const values = recordSnapshot({
    kind: "content", source: "get_range_values", workbook, revision: getWorkbookRevision(),
    sheetId: "{id-1}", sheetName: "Продажи", address: "D2:D6", coverage, payload: [[2], [3]]
  });
  assert.ok(values);
  const recalled = recallSnapshot(values.id, workbook, getWorkbookRevision(), getRevisionCoverage());
  assert.equal(recalled.state, "fresh");
  assert.equal(recalled.historical, false);
  assert.equal(recalled.reason, undefined);
  // Имя листа могло устареть, и об этом сказано отдельно от свежести данных.
  assert.equal(recalled.sheetNameUnverified, true);

  // Настоящее изменение содержимого по-прежнему делает снимок устаревшим.
  bumpWorkbookRevision("content");
  assert.equal(recallSnapshot(values.id, workbook, getWorkbookRevision(), getRevisionCoverage()).state, "stale");
});

test("details snapshot stays fresh without name tracking but reacts to format changes", () => {
  freshStore();
  const coverage = { content: true, structure: true, format: true, protection: true, sheetNames: false };
  setRevisionCoverage(coverage);

  const details = recordSnapshot({
    kind: "details", source: "get_range_details", workbook, revision: getWorkbookRevision(),
    sheetId: "{id-1}", sheetName: "Продажи", address: "A1:H1", coverage, payload: {}
  });
  assert.ok(details);
  assert.equal(recallSnapshot(details.id, workbook, getWorkbookRevision(), getRevisionCoverage()).state, "fresh");
  bumpWorkbookRevision("format");
  assert.equal(recallSnapshot(details.id, workbook, getWorkbookRevision(), getRevisionCoverage()).state, "stale");
});

test("snapshot without a sheet name carries no name warning", () => {
  freshStore();
  const coverage = { content: true, structure: true, format: true, protection: true, sheetNames: false };
  setRevisionCoverage(coverage);
  const search = recordSnapshot({
    kind: "search", source: "search_workbook", workbook, revision: getWorkbookRevision(), coverage, payload: []
  });
  assert.ok(search);
  const recalled = recallSnapshot(search.id, workbook, getWorkbookRevision(), getRevisionCoverage());
  assert.equal(recalled.state, "fresh");
  assert.equal(recalled.sheetNameUnverified, undefined);
});
