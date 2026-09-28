import test from "node:test";
import assert from "node:assert/strict";
import { documentConversationId, documentConversationKey, ensureDocumentConversationId } from "./documentId";

function officeWithSettings(initial: Record<string, unknown> = {}, saveStatus = "succeeded") {
  const values = new Map(Object.entries(initial));
  let saves = 0;
  (globalThis as any).Office = {
    context: {
      document: {
        settings: {
          get: (name: string) => values.get(name) ?? null,
          set: (name: string, value: unknown) => values.set(name, value),
          saveAsync: (callback: (result: { status: string }) => void) => { saves += 1; callback({ status: saveStatus }); }
        }
      }
    }
  };
  return { values, saves: () => saves };
}

test("a new conversation writes a random id into the workbook once, and the same id is found again", async () => {
  const office = officeWithSettings();
  assert.equal(documentConversationId(), null);
  const id = await ensureDocumentConversationId();
  assert.match(String(id), /^[0-9a-f-]{36}$/);
  assert.equal(documentConversationId(), id, "после перезапуска панели номер читается из книги");
  assert.equal(await ensureDocumentConversationId(), id);
  assert.equal(office.saves(), 1, "повторно номер не пишется");
  assert.equal(documentConversationKey(id!), `doc:${id}`);
});

test("a value that is not our id is ignored, and a failed save gives no id", async () => {
  officeWithSettings({ "amai.conversationId": "../../чужое" });
  assert.equal(documentConversationId(), null);
  officeWithSettings({}, "failed");
  assert.equal(await ensureDocumentConversationId(), null);
});
