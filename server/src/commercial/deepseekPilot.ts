import { serializeMessages } from "../protocol.js";
import { normalizeUsage } from "../usageMetrics.js";
import { fail } from "./core.js";
import { selectedTools, type ChatInput, type ProviderAdapter, type ProviderResult, type Quote } from "./providers.js";

export async function boundedJson(response: Response, maxBytes: number): Promise<any> {
  if (!response.body) throw new Error("empty_provider_response");
  const reader = response.body.getReader(); const chunks: Uint8Array[] = []; let size = 0;
  try {
    while (true) {
      const part = await reader.read(); if (part.done) break;
      size += part.value.length;
      if (size > maxBytes) throw new Error("provider_response_limit");
      chunks.push(part.value);
    }
    return JSON.parse(Buffer.concat(chunks).toString("utf8"));
  } finally { await reader.cancel().catch(() => undefined); reader.releaseLock(); }
}

export interface PilotPricing {
  model: string; version: string; inputMicroUsdPerMillion: bigint; outputMicroUsdPerMillion: bigint;
  maxOutputTokens: number; validUntil: number;
}
export interface HostedSecretSource { resolve(provider: string): Promise<string> }
/** Contract-ready adapter, deliberately absent from the runtime registry. No implicit retries/fallbacks. */
export class DeepSeekPilot implements ProviderAdapter {
  readonly id = "deepseek";
  constructor(private price: PilotPricing, private hosted: HostedSecretSource, private liveApproved = false, private fetcher: typeof fetch = fetch) {}
  private cost(input: number, output: number) {
    return (BigInt(input) * this.price.inputMicroUsdPerMillion + BigInt(output) * this.price.outputMicroUsdPerMillion + 999_999n) / 1_000_000n;
  }
  quote(input: ChatInput): Quote {
    const p = this.price;
    if (input.model !== p.model || !p.version || p.validUntil <= Date.now() || p.inputMicroUsdPerMillion <= 0n || p.outputMicroUsdPerMillion <= 0n || !Number.isInteger(p.maxOutputTokens) || p.maxOutputTokens < 1 || p.maxOutputTokens > 4096) return fail(503, "approved_pricing_required");
    // Charge all input at cache-miss rate; count UTF-8 bytes plus framing allowance as conservative token bound.
    const inputBound = Buffer.byteLength(JSON.stringify(serializeMessages(input.messages, this.id))) + Buffer.byteLength(JSON.stringify(selectedTools(input))) + 4096;
    if (inputBound > 160_000) return fail(413, "context_limit");
    return { maxCost: this.cost(inputBound, p.maxOutputTokens), pricingVersion: p.version, maxOutputTokens: p.maxOutputTokens };
  }
  async run(input: ChatInput, options: { signal: AbortSignal; attemptId: string; credential?: string; quote: Quote }): Promise<ProviderResult> {
    if (!this.liveApproved) return fail(503, "real_spend_disabled");
    const current = this.quote(input);
    if (current.maxCost !== options.quote.maxCost || current.pricingVersion !== options.quote.pricingVersion) return fail(503, "pricing_changed");
    const key = input.keyMode === "byok" ? options.credential : await this.hosted.resolve(this.id);
    if (!key) return fail(503, "server_credential_required");
    const response = await this.fetcher("https://api.deepseek.com/chat/completions", { method: "POST", redirect: "error", signal: options.signal,
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${key}` },
      body: JSON.stringify({ model: input.model, messages: serializeMessages(input.messages, this.id), tools: selectedTools(input),
        stream: false, max_tokens: options.quote.maxOutputTokens, thinking: { type: "disabled" } }) });
    if (!response.ok) throw new Error("provider_request_failed");
    const data = await boundedJson(response, 131_072);
    const choice = data.choices?.[0], message = choice?.message;
    if (!message || !["stop", "tool_calls"].includes(choice.finish_reason)) throw new Error("provider_incomplete_response");
    const raw = data.usage;
    const valid = raw && [raw.prompt_tokens, raw.completion_tokens].every(v => Number.isSafeInteger(v) && v >= 0);
    const usage = valid ? normalizeUsage(raw) : null;
    const calls = message.tool_calls ?? [];
    if (!Array.isArray(calls) || calls.length > 8) throw new Error("provider_tool_limit");
    return { text: typeof message.content === "string" ? message.content : "", cost: usage ? this.cost(usage.promptTokens, usage.completionTokens) : null, usage,
      finishReason: choice.finish_reason, toolCalls: calls.map(c => ({ id: c.id, name: c.function?.name, arguments: c.function?.arguments })) };
  }
}
