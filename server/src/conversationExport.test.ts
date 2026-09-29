import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { checkedFileName, exportConversation, ExportError, freePath } from "./conversationExport.js";

test("only a plain .md file name is accepted", () => {
  assert.equal(checkedFileName("Продажи 2026-09-30 14-05.md"), "Продажи 2026-09-30 14-05.md");
  for (const bad of ["..\\..\\Windows\\evil.md", "../x.md", "C:\\x.md", "a/b.md", "x.exe", "x.md.bat", "", "..md", "con?.md", 42]) {
    assert.throws(() => checkedFileName(bad), ExportError, String(bad));
  }
});

test("an existing file is never overwritten: (2), (3) are added", () => {
  const taken = new Set(["F/Книга.md", "F/Книга (2).md"].map((p) => join(...p.split("/"))));
  assert.equal(freePath("F", "Книга.md", (p) => taken.has(p)), join("F", "Книга (3).md"));
});

test("the conversation is written as UTF-8 text into the folder, created if missing", async () => {
  const root = mkdtempSync(join(tmpdir(), "export-"));
  try {
    const folder = join(root, "am.AI", "Беседы");
    const first = await exportConversation(folder, "Продажи.md", "# Беседа\n\nПривет");
    assert.equal(first, join(folder, "Продажи.md"));
    assert.equal(readFileSync(first, "utf8"), "# Беседа\n\nПривет");
    writeFileSync(join(folder, "Продажи (2).md"), "чужой файл");
    const third = await exportConversation(folder, "Продажи.md", "ещё раз");
    assert.equal(third, join(folder, "Продажи (3).md"));
    assert.equal(readFileSync(join(folder, "Продажи (2).md"), "utf8"), "чужой файл", "чужой файл не тронут");
    await assert.rejects(exportConversation(folder, "Пусто.md", "  "), /пуста/);
    await assert.rejects(exportConversation(folder, "Большая.md", "x".repeat(3 * 1024 * 1024 + 1)), /слишком большая/);
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});
