/**
 * Цены моделей для оценки расхода (01.10.2026), $ за 1 млн токенов:
 * вход без кэша, вход из кэша, выход.
 *
 * OpenRouter сам присылает цену каждого обращения — её показываем точно.
 * DeepSeek, OpenAI, Kimi, Z.AI и Grok напрямую присылают только токены: цену
 * панель считает по этому прайсу и помечает «≈» — поставщики меняют цены, а
 * прайс обновляется с выпуском. Источники: DeepSeek — api-docs.deepseek.com
 * (цены в часы пик вдвое выше), остальные — каталог OpenRouter (те же модели
 * по цене поставщика). Сверка: тестовый запрос GPT-6.1 Sol сходится с расчётом.
 */

import type { StepUsage } from "./client";

interface Price { input: number; cached: number; output: number }

const PRICES: Record<string, Price> = {
  "openai/gpt-6.1-sol": { input: 2, cached: 0.1, output: 10 },
  "openai/gpt-6-sol": { input: 2, cached: 0.2, output: 10 },
  "openai/gpt-6-luna": { input: 0.1, cached: 0.01, output: 0.5 },
  "openai/gpt-6-astra": { input: 10, cached: 1, output: 50 },
  "openai/gpt-5.6-sol": { input: 2, cached: 0.2, output: 10 },
  "openai/gpt-5.6-terra": { input: 2, cached: 0.2, output: 12 },
  "openai/gpt-5.6-luna": { input: 0.2, cached: 0.02, output: 1.2 },
  "openai/gpt-5.5": { input: 5, cached: 0.5, output: 30 },
  "openai/gpt-5.4": { input: 2.5, cached: 0.25, output: 15 },
  "moonshot/kimi-k3": { input: 0.365, cached: 0.365, output: 10 },
  "moonshot/kimi-k2.7-code": { input: 0.6712, cached: 0.18, output: 3.35 },
  "moonshot/kimi-k2.6": { input: 0.4341, cached: 0.0731, output: 1.828 },
  "zai/glm-5.3": { input: 0.2219, cached: 0.1775, output: 3.39 },
  "zai/glm-5.3-flash": { input: 0.15, cached: 0.03, output: 0.5 },
  "zai/glm-5.2": { input: 0.41, cached: 0.26, output: 3.99 },
  "zai/glm-5.1": { input: 0.9646, cached: 0.1791, output: 3.0316 },
  "zai/glm-5-turbo": { input: 1.2, cached: 0.24, output: 4 },
  "zai/glm-5": { input: 0.6, cached: 0.12, output: 1.92 },
  "xai/grok-4.7": { input: 2, cached: 0.5, output: 6 },
  "xai/grok-4.6": { input: 2, cached: 0.5, output: 6 },
  "xai/grok-4.5": { input: 2, cached: 0.3, output: 6 },
  "xai/grok-4.3": { input: 1.25, cached: 0.2, output: 2.5 }
};

/** DeepSeek: цены вне пика; в часы пик (01–04 и 06–10 UTC, пн–пт) — вдвое выше. */
const DEEPSEEK: Record<string, Price> = {
  "deepseek-flash": { input: 0.15, cached: 0.003, output: 0.6 },
  "deepseek-v4-pro": { input: 0.66, cached: 0.022, output: 1.98 }
};

export function deepseekPeak(at: Date): boolean {
  const day = at.getUTCDay();
  const hour = at.getUTCHours();
  return day >= 1 && day <= 5 && ((hour >= 1 && hour < 4) || (hour >= 6 && hour < 10));
}

export function priceOf(provider: string, model: string, at = new Date()): Price | null {
  if (provider === "deepseek") {
    const base = DEEPSEEK[model];
    if (!base) return null;
    const k = deepseekPeak(at) ? 2 : 1;
    return { input: base.input * k, cached: base.cached * k, output: base.output * k };
  }
  return PRICES[`${provider}/${model}`] ?? null;
}

/** Цена обращения: точная от поставщика или оценка по прайсу; null — неизвестна. */
export function stepCost(provider: string, model: string, usage: StepUsage, at = new Date()): { cost: number; estimated: boolean } | null {
  if (usage.cost !== undefined) return { cost: usage.cost, estimated: false };
  if (provider === "ollama" || provider === "qwen") return { cost: 0, estimated: false };
  const price = priceOf(provider, model, at);
  if (!price) return null;
  const cached = Math.min(usage.cachedTokens, usage.promptTokens);
  const cost = ((usage.promptTokens - cached) * price.input + cached * price.cached + usage.completionTokens * price.output) / 1_000_000;
  return { cost, estimated: true };
}
