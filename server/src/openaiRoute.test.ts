import test from "node:test";
import assert from "node:assert/strict";
import {
  DEFAULT_ROUTE,
  cheapestSupportedEffort,
  nextRouteAfterRejection,
  parseSupportedEfforts,
  rememberRoute,
  resetLearnedRoutes,
  routeFor,
  type OpenAiRoute
} from "./openaiRoute.js";

/** Настоящие отказы gpt-6-astra, снятые с OpenAI 16 сентября 2026 года. */
const rejectValue = JSON.stringify({
  error: {
    message: "Unsupported value: 'reasoning_effort' does not support 'none' with this model. Supported values are: 'low', 'medium', 'high', and 'xhigh'.",
    type: "invalid_request_error", param: "reasoning_effort", code: "unsupported_value"
  }
});
const rejectTools = JSON.stringify({
  error: {
    message: "Function tools with reasoning_effort are not supported for gpt-6-astra in /v1/chat/completions. To use function tools, use /v1/responses or set reasoning_effort to 'none'.",
    type: "invalid_request_error", param: "reasoning_effort", code: null
  }
});

test("astra's two contradictory refusals lead to the responses API", () => {
  // Замер показал круг: none отвергается как значение, low — вместе
  // с функциями, без параметра действует умолчание и отвергается так же.
  const first = nextRouteAfterRejection(400, rejectValue, DEFAULT_ROUTE, []);
  assert.deepEqual(first?.route, { api: "chat", effort: "low" });

  const tried: OpenAiRoute[] = [DEFAULT_ROUTE];
  const second = nextRouteAfterRejection(400, rejectTools, first!.route, tried);
  assert.deepEqual(second?.route, { api: "responses", effort: null });
  assert.match(second!.reason, /\/v1\/responses/);
});

test("the loop cannot return to a route already refused", () => {
  const tried: OpenAiRoute[] = [{ api: "chat", effort: "none" }, { api: "chat", effort: "low" }];
  // Совет «поставьте none» бесполезен: это значение уже отвергнуто.
  const retry = nextRouteAfterRejection(400, rejectTools, { api: "chat", effort: null }, tried);
  assert.deepEqual(retry?.route, { api: "responses", effort: null });

  // Когда испробовано всё, пробовать больше нечего.
  const exhausted = nextRouteAfterRejection(400, rejectTools, { api: "responses", effort: null }, [...tried, { api: "chat", effort: null }]);
  assert.equal(exhausted, null);
});

test("the cheapest allowed effort wins: the tool loop gains nothing from long thinking", () => {
  assert.deepEqual(parseSupportedEfforts(JSON.parse(rejectValue).error.message), ["low", "medium", "high", "xhigh"]);
  assert.equal(cheapestSupportedEffort(["high", "low", "medium"]), "low");
  assert.equal(cheapestSupportedEffort(["medium", "high"]), "medium");
  assert.equal(cheapestSupportedEffort([]), null);
});

test("other failures are not mistaken for this one", () => {
  assert.equal(nextRouteAfterRejection(500, rejectValue, DEFAULT_ROUTE), null, "не 400");
  assert.equal(nextRouteAfterRejection(400, "не json", DEFAULT_ROUTE), null, "тело не разбирается");
  assert.equal(nextRouteAfterRejection(400, JSON.stringify({ error: { param: "model", message: "bad model" } }), DEFAULT_ROUTE), null, "другой параметр");
  assert.equal(nextRouteAfterRejection(400, JSON.stringify({}), DEFAULT_ROUTE), null, "пустое тело");
});

test("a learned route is reused per model and does not leak to others", () => {
  resetLearnedRoutes();
  assert.deepEqual(routeFor("openai", "gpt-6-astra"), DEFAULT_ROUTE);

  rememberRoute("openai", "gpt-6-astra", { api: "responses", effort: null });
  assert.deepEqual(routeFor("openai", "gpt-6-astra"), { api: "responses", effort: null });
  assert.deepEqual(routeFor("openai", "gpt-5.6-terra"), DEFAULT_ROUTE, "другая модель не затронута");

  resetLearnedRoutes();
  assert.deepEqual(routeFor("openai", "gpt-6-astra"), DEFAULT_ROUTE);
});
