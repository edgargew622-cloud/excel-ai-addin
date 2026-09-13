import test from "node:test";
import assert from "node:assert/strict";
import {
  compareVersions,
  formatCapabilityLog,
  probeVersions,
  resolveCeiling,
  resolveFeatures,
  type CapabilityReport
} from "./capabilities";

test("versions compare numerically, not as strings", () => {
  // Строковое сравнение поставило бы 1.10 ниже 1.9 — и замер соврал бы
  // про доступность комментариев на сборке с потолком 1.14.
  assert.equal(compareVersions("1.10", "1.9") > 0, true);
  assert.equal(compareVersions("1.9", "1.10") < 0, true);
  assert.equal(compareVersions("1.14", "1.14"), 0);
  assert.equal(compareVersions("1.20", "1.3") > 0, true);
});

test("ceiling is the highest supported set, not the first gap", () => {
  assert.equal(resolveCeiling(["1.1", "1.9", "1.10", "1.14"]), "1.14");
  assert.equal(resolveCeiling([]), null);
});

test("probe tolerates a throwing isSetSupported", () => {
  const supported = probeVersions((_set, version) => {
    if (version === "1.3") throw new Error("host error");
    return version === "1.1" || version === "1.2";
  }, ["1.1", "1.2", "1.3"]);
  assert.deepEqual(supported, ["1.1", "1.2"]);
});

test("features unavailable on this build carry a human-readable reason", () => {
  const statuses = resolveFeatures(["1.9", "1.10"], [
    { id: "comments", label: "Комментарии", requires: "1.10" },
    { id: "notes", label: "Примечания", requires: "1.18" }
  ]);

  const comments = statuses.find((f) => f.id === "comments");
  assert.equal(comments?.available, true);
  assert.equal(comments?.reason, "");

  const notes = statuses.find((f) => f.id === "notes");
  assert.equal(notes?.available, false);
  assert.match(String(notes?.reason), /1\.18/);
  assert.match(String(notes?.reason), /до 1\.10/);
});

test("an available feature with an unconfirmed version says so", () => {
  const statuses = resolveFeatures(["1.15"], [
    { id: "direct_dependents", label: "Зависимые", requires: "1.15", versionUnconfirmed: true }
  ]);
  assert.equal(statuses[0].available, true);
  assert.match(statuses[0].reason, /требует подтверждения/);
});

test("an unavailable feature with an unconfirmed version does not present a guess as fact", () => {
  const statuses = resolveFeatures(["1.14"], [
    { id: "guessed", label: "Неизвестная", requires: "1.15", versionUnconfirmed: true }
  ]);
  assert.equal(statuses[0].available, false);
  assert.match(statuses[0].reason, /доступно до 1\.14/);
  // Догадка о требуемой версии не должна звучать как установленное требование:
  // иначе возможность спишут как недоступную, не проверив документацию.
  assert.match(statuses[0].reason, /не подтверждена по документации/);
});

test("capability log line carries no workbook content", () => {
  const report: CapabilityReport = {
    ceiling: "1.14",
    supported: ["1.9", "1.14"],
    features: [
      { id: "comments", label: "Комментарии", requires: "1.10", available: true, reason: "" },
      { id: "notes", label: "Примечания", requires: "1.18", available: false, reason: "нет" }
    ],
    host: "Excel",
    platform: "PC",
    officeVersion: "16.0.14334.20906"
  };
  const line = formatCapabilityLog(report);
  assert.match(line, /excelApiCeiling=1\.14/);
  assert.match(line, /missing=notes/);
  // Журнал состоит только из версий и признаков: подставить туда данные книги
  // неоткуда, и это фиксируется тестом.
  assert.equal(line.includes("Лист"), false);
});
