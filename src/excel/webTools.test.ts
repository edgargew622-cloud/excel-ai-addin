import test from "node:test";
import assert from "node:assert/strict";
import { web_search } from "./webTools";

test("23.10: several searches go in one call — each query answered on its own, one failure does not sink the rest", async () => {
  const asked: string[] = [];
  (globalThis as any).fetch = async (_path: string, init: any) => {
    const body = JSON.parse(init.body);
    asked.push(body.query);
    if (body.query === "сломано") return { ok: false, status: 400, json: async () => ({ error: { message: "плохой запрос" } }) };
    return { ok: true, status: 200, json: async () => ({ service: "DeepSeek — поиск", results: [{ title: body.query, url: "https://ru.wikipedia.org/wiki/" + body.query, snippet: "" }] }) };
  };
  const out = await web_search({ queries: ["Хаш", "Долма", "сломано"], domains: ["ru.wikipedia.org"] }) as any;
  assert.deepEqual(asked.sort(), ["Долма", "Хаш", "сломано"]);
  assert.equal(out.searches.length, 3);
  assert.equal(out.searches[0].results[0].url, "https://ru.wikipedia.org/wiki/Хаш");
  assert.match(out.searches[2].error, /плохой запрос/);
  assert.equal(out.untrustedContent, true);
  await assert.rejects(() => web_search({}), /query или queries/);
});
