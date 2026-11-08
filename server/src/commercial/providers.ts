import { serializeMessages, type InternalMessage } from "../protocol.js";
import { normalizeUsage, type Usage } from "../usageMetrics.js";
import { fail, uuid } from "./core.js";
import catalog from "./toolCatalog.json" with { type: "json" };

export interface ChatInput {
  provider: string; model: string; keyMode: "hosted" | "byok"; credentialId?: string;
  messages: InternalMessage[];
  taskId?: string;
  toolNames?: string[];
}
export interface Quote { maxCost: bigint; pricingVersion: string; maxOutputTokens: number }
export interface ProviderResult { text: string; cost: bigint | null; usage: Usage | null; toolCalls?: Array<{ id: string; name: string; arguments: string }>; finishReason?: string }
export interface ProviderAdapter {
  readonly id: string;
  quote(input: ChatInput): Quote;
  run(input: ChatInput, options: { signal: AbortSignal; attemptId: string; credential?: string; quote: Quote }): Promise<ProviderResult>;
}

export function parseChat(raw: unknown): ChatInput {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return fail(400, "invalid_chat");
  const r = raw as Record<string, unknown>;
  if (Object.keys(r).some(k => !["provider", "model", "keyMode", "credentialId", "messages", "taskId", "toolNames"].includes(k))) return fail(400, "unknown_chat_property");
  if (typeof r.provider !== "string" || r.provider.length > 64 || typeof r.model !== "string" || r.model.length > 128) return fail(400, "invalid_model");
  if (r.keyMode !== "hosted" && r.keyMode !== "byok") return fail(400, "invalid_key_mode");
  if (r.keyMode === "byok" ? !uuid(r.credentialId) : r.credentialId !== undefined) return fail(400, "invalid_credential_reference");
  if (r.taskId !== undefined && !uuid(r.taskId)) return fail(400, "invalid_task_id");
  if (r.toolNames !== undefined && (!Array.isArray(r.toolNames) || r.toolNames.length > 70 || new Set(r.toolNames).size !== r.toolNames.length || r.toolNames.some(n => typeof n !== "string" || !catalog.some(t => t.function.name === n)))) return fail(400, "invalid_tool_catalog");
  if (!Array.isArray(r.messages) || !r.messages.length || r.messages.length > 128) return fail(400, "invalid_messages");
  const messages = r.messages.map(m => {
    if (!m || typeof m !== "object" || Array.isArray(m)) return fail(400, "invalid_message");
    if (m.role === "assistant") {
      if (Object.keys(m).some(k => !["role", "content", "provider", "reasoning_content", "tool_calls"].includes(k)) || m.content !== null && typeof m.content !== "string") return fail(400, "invalid_message");
      if (m.tool_calls !== undefined && (!Array.isArray(m.tool_calls) || m.tool_calls.length > 8 || m.tool_calls.some((c: any) => !c || Object.keys(c).some(k => !["id", "name", "arguments"].includes(k)) || typeof c.id !== "string" || !/^[\w-]{1,128}$/.test(c.id) || !catalog.some(t => t.function.name === c.name) || typeof c.arguments !== "string" || c.arguments.length > 32768))) return fail(400, "invalid_tool_call");
      return { role: "assistant", content: m.content, ...(m.tool_calls ? { tool_calls: m.tool_calls.map((c: any) => ({ id: c.id, name: c.name, arguments: c.arguments })) } : {}) } as InternalMessage;
    }
    if (m.role === "tool") {
      if (Object.keys(m).some(k => !["role", "content", "tool_call_id"].includes(k)) || typeof m.content !== "string" || typeof m.tool_call_id !== "string" || !/^[\w-]{1,128}$/.test(m.tool_call_id)) return fail(400, "invalid_message");
      return { role: "tool", content: m.content, tool_call_id: m.tool_call_id } as InternalMessage;
    }
    if (!["system", "user"].includes(m.role) || Object.keys(m).some(k => !["role", "content"].includes(k)) || typeof m.content !== "string" || !m.content.trim()) return fail(400, "invalid_message");
    return { role: m.role, content: m.content } as InternalMessage;
  });
  if (Buffer.byteLength(JSON.stringify(messages)) > 131_072) return fail(413, "context_limit");
  return { provider: r.provider, model: r.model, keyMode: r.keyMode, ...(typeof r.credentialId === "string" ? { credentialId: r.credentialId } : {}), messages,
    ...(r.taskId ? { taskId: r.taskId as string } : {}), ...(r.toolNames ? { toolNames: r.toolNames as string[] } : {}) };
}

export const selectedTools = (input: ChatInput) => catalog.filter(t => input.toolNames?.includes(t.function.name));

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

/** Deliberately scripted synthetic workbook scenario. Never represents model reasoning or real AI quality. */
export class FakeAgentProvider extends FakeProvider {
  override quote(input: ChatInput): Quote {
    if (input.model !== "c2-pilot") return super.quote(input);
    return { maxCost: 10n, pricingVersion: "fake-v2", maxOutputTokens: 1024 };
  }
  override async run(input: ChatInput, options: { signal: AbortSignal }): Promise<ProviderResult> {
    options.signal.throwIfAborted(); this.calls++;
    const current = input.messages.slice(input.messages.findLastIndex(m => m.role === "user") + 1);
    const results = current.filter(m => m.role === "tool");
    const assistantCalls = current.flatMap(m => m.role === "assistant" ? m.tool_calls ?? [] : []);
    const reconciled = results.find(m => m.role === "tool" && assistantCalls.some(c => c.id === m.tool_call_id && c.name === "reconcile_ranges"));
    const result: ProviderResult = { text: "Synthetic pilot complete. Review the Excel operation result.", cost: 3n, usage: normalizeUsage({ prompt_tokens: 2, completion_tokens: 2 }), finishReason: "stop" };
    if (reconciled) return result;
    const name = input.toolNames?.includes("reconcile_ranges") ? "reconcile_ranges" : "enable_tools";
    if (!input.toolNames?.includes(name)) return { ...result, text: "Synthetic pilot requires the accounting tool group." };
    return { ...result, text: "Synthetic pilot: confirm the existing reconciliation operation.", finishReason: "tool_calls",
      toolCalls: [{ id: `pilot_${this.calls}`, name, arguments: JSON.stringify(name === "enable_tools" ? { groups: ["accounting"] } : {
        leftSheet: "PaymentsA", leftAddress: "A1:D4", rightSheet: "PaymentsB", rightAddress: "A1:D4", resultSheet: "PilotResult"
      }) }] };
  }
}
