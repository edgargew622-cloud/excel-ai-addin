import test from "node:test";
import assert from "node:assert/strict";
import {
  cheapestSupportedEffort,
  parseSupportedEfforts,
  reasoningEffortFor,
  rejectedReasoningEffort,
  rememberEffort,
  resetLearnedEfforts
} from "./reasoningEffort.js";

/** Настоящий ответ OpenAI на gpt-6-astra, из-за которого запрос падал с 400. */
const astraRejection = JSON.stringify({
  error: {
    message: "Unsupported value: 'reasoning_effort' does not support 'none' with this model. Supported values are: 'low', 'medium', 'high', and 'xhigh'.",
    type: "invalid_request_error",
    param: "reasoning_effort",
    code: "unsupported_value"
  }
});

test("the model's own rejection names the value to retry with", () => {
  assert.equal(rejectedReasoningEffort(400, astraRejection), "low");
  assert.deepEqual(parseSupportedEfforts(JSON.parse(astraRejection).error.message), ["low", "medium", "high", "xhigh"]);
});

test("the cheapest supported value wins: the agent loop gains nothing from long thinking", () => {
  assert.equal(cheapestSupportedEffort(["high", "low", "medium"]), "low");
  assert.equal(cheapestSupportedEffort(["none", "low"]), "none");
  assert.equal(cheapestSupportedEffort(["medium", "high"]), "medium");
  // Разобрать не удалось — берём заведомо принимаемое значение, а не none.
  assert.equal(cheapestSupportedEffort([]), "low");
});

test("other failures are not mistaken for this one", () => {
  assert.equal(rejectedReasoningEffort(500, astraRejection), null, "не 400");
  assert.equal(rejectedReasoningEffort(400, "не json"), null, "тело не разбирается");
  assert.equal(rejectedReasoningEffort(400, JSON.stringify({ error: { param: "model", message: "no" } })), null, "другой параметр");
  assert.equal(rejectedReasoningEffort(400, JSON.stringify({})), null, "пустое тело");
  assert.equal(rejectedReasoningEffort(401, JSON.stringify({ error: { message: "bad key" } })), null, "ключ");
});

test("a learned value is reused per model and does not leak to others", () => {
  resetLearnedEfforts();
  assert.equal(reasoningEffortFor("openai", "gpt-6-astra"), "none", "до отказа просим минимум");

  rememberEffort("openai", "gpt-6-astra", "low");
  assert.equal(reasoningEffortFor("openai", "gpt-6-astra"), "low", "после отказа не повторяем ошибку");
  assert.equal(reasoningEffortFor("openai", "gpt-5.6-terra"), "none", "другая модель не затронута");

  resetLearnedEfforts();
  assert.equal(reasoningEffortFor("openai", "gpt-6-astra"), "none");
});
