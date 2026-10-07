import test from "node:test";
import assert from "node:assert/strict";
import { networkErrorText } from "./networkError.js";

test("a connect timeout is named in plain words with what to do", () => {
  const error = Object.assign(new TypeError("fetch failed"), { cause: { code: "UND_ERR_CONNECT_TIMEOUT" } });
  assert.match(networkErrorText(error, "xAI")!, /Не удалось связаться с xAI: сервер не ответил.*другого провайдера/);
  assert.match(networkErrorText(Object.assign(new TypeError("fetch failed"), { cause: { code: "ENOTFOUND" } }), "DeepSeek")!, /DNS/);
  assert.equal(networkErrorText(new Error("401 Unauthorized"), "xAI"), null);
});
