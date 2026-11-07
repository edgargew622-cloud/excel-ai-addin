import { serializeMessages, type InternalMessage } from "../protocol.js";
import { normalizeUsage, type Usage } from "../usageMetrics.js";
import { fail, uuid } from "./core.js";

export interface ChatInput {
  provider: string; model: string; keyMode: "hosted" | "byok"; credentialId?: string;
  messages: Array<{ role: "user" | "assistant"; content: string }>;
}
export interface Quote { maxCost: bigint; pricingVersion: string; maxOutputTokens: number }
export interface ProviderResult { text: string; cost: bigint | null; usage: Usage | null }
export interface ProviderAdapter {
  readonly id: string;
  quote(input: ChatInput): Quote;
  run(input: ChatInput, options: { signal: AbortSignal; attemptId: string; credential?: string; quote: Quote }): Promise<ProviderResult>;
}

export function parseChat(raw: unknown): ChatInput {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return fail(400, "invalid_chat");
  const r = raw as Record<string, unknown>;
  if (Object.keys(r).some(k => !["provider", "model", "keyMode", "credentialId", "messages"].includes(k))) return fail(400, "unknown_chat_property");
  if (typeof r.provider !== "string" || r.provider.length > 64 || typeof r.model !== "string" || r.model.length > 128) return fail(400, "invalid_model");
  if (r.keyMode !== "hosted" && r.keyMode !== "byok") return fail(400, "invalid_key_mode");
  if (r.keyMode === "byok" ? !uuid(r.credentialId) : r.credentialId !== undefined) return fail(400, "invalid_credential_reference");
  if (!Array.isArray(r.messages) || !r.messages.length || r.messages.length > 16) return fail(400, "invalid_messages");
  const messages = r.messages.map(m => {
    if (!m || typeof m !== "object" || Array.isArray(m) || Object.keys(m).some(k => !["role", "content"].includes(k))) return fail(400, "invalid_message");
    if ((m.role !== "user" && m.role !== "assistant") || typeof m.content !== "string" || !m.content.trim()) return fail(400, "invalid_message");
    return { role: m.role as "user" | "assistant", content: m.content as string };
  });
  if (Buffer.byteLength(JSON.stringify(messages)) > 16_384) return fail(413, "context_limit");
  return { provider: r.provider, model: r.model, keyMode: r.keyMode, ...(typeof r.credentialId === "string" ? { credentialId: r.credentialId } : {}), messages };
}

/** Only this deterministic no-network adapter is enabled in C1. All real models stay in the local edition. */
export class FakeProvider implements ProviderAdapter {
  readonly id = "fake";
  calls = 0;
  quote(input: ChatInput): Quote {
    if (input.model !== "c1-test") return fail(400, "model_not_allowed");
    return { maxCost: 10n, pricingVersion: "fake-v1", maxOutputTokens: 64 };
  }
  async run(input: ChatInput, options: { signal: AbortSignal }): Promise<ProviderResult> {
    options.signal.throwIfAborted();
    this.calls++;
    // Reuse the existing provider wire-format boundary; do not persist the messages.
    serializeMessages(input.messages as InternalMessage[], this.id);
    return { text: "C1 test response", cost: 3n, usage: normalizeUsage({ prompt_tokens: 2, completion_tokens: 2, total_tokens: 4 }) };
  }
}
