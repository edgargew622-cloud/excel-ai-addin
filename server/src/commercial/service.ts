import type { Context } from "./core.js";
import { fail, sha256 } from "./core.js";
import { Ledger } from "./ledger.js";
import { CredentialStore } from "./secrets.js";
import { parseChat, type ProviderAdapter } from "./providers.js";

async function bounded<T>(signal: AbortSignal, run: () => Promise<T>): Promise<T> {
  signal.throwIfAborted();
  let listener: () => void = () => undefined;
  const aborted = new Promise<never>((_, reject) => {
    listener = () => reject(new Error("provider_cancelled_or_timed_out"));
    signal.addEventListener("abort", listener, { once: true });
  });
  try { return await Promise.race([run(), aborted]); }
  finally { signal.removeEventListener("abort", listener); }
}

/** All metering is server-side; client fields cannot set a price, role or quota. */
export class ChatService {
  constructor(readonly ledger: Ledger, readonly credentials: CredentialStore, private readonly adapters: ReadonlyMap<string, ProviderAdapter>) {}
  async chat(ctx: Context, raw: unknown, idempotencyKey: string, signal: AbortSignal) {
    if (!/^[a-zA-Z0-9_-]{8,100}$/.test(idempotencyKey)) return fail(400, "idempotency_key_required");
    const input = parseChat(raw);
    const provider = this.adapters.get(input.provider);
    if (!provider) return fail(400, "provider_not_allowed");
    const quote = provider.quote(input);
    if (!Number.isInteger(quote.maxOutputTokens) || quote.maxOutputTokens < 1 || quote.maxOutputTokens > 4096) return fail(503, "unbounded_provider");
    const reserve = await this.ledger.reserve(ctx, {
      idempotencyKey, bodyHash: sha256(JSON.stringify({ userId: ctx.userId, input })), provider: input.provider, model: input.model,
      mode: input.keyMode, credentialId: input.credentialId, providerBound: quote.maxCost, pricingVersion: quote.pricingVersion
    });
    if (!reserve.created) return { request: reserve.request, replay: true, text: null };
    const requestId = reserve.request.id;
    let dispatched = false;
    try {
      signal.throwIfAborted();
      const credential = input.credentialId ? await this.credentials.resolve(ctx, input.credentialId, input.provider) : undefined;
      signal.throwIfAborted();
      const attemptId = await this.ledger.dispatch(ctx, requestId, credential?.version);
      dispatched = true;
      const providerSignal = AbortSignal.any([signal, AbortSignal.timeout(15_000)]);
      const result = await bounded(providerSignal, () => provider.run(input, { signal: providerSignal, attemptId, credential: credential?.value, quote }));
      if (typeof result.text !== "string" || Buffer.byteLength(result.text) > quote.maxOutputTokens * 16) throw new Error("provider_output_limit");
      // Even on disconnect, persist provider observations before deciding whether to return text.
      if (result.cost === null || result.usage === null) {
        await this.ledger.uncertain(ctx, requestId);
        return { request: await this.ledger.get(ctx, requestId), replay: false, text: null };
      }
      const request = await this.ledger.settle(ctx, requestId, result.cost, result.usage as unknown as Record<string, unknown>);
      return { request, replay: false, text: request.state === "settled" && !signal.aborted ? result.text : null };
    } catch (e) {
      // A DB failure must not release possibly billed spend. Recovery later inspects durable state.
      if (dispatched) await this.ledger.uncertain(ctx, requestId).catch(() => undefined);
      else await this.ledger.release(ctx, requestId).catch(() => undefined);
      throw e;
    }
  }
}
