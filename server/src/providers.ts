/**
 * Реестр провайдеров. Все перечисленные — OpenAI-совместимые:
 * один и тот же POST {baseURL}/chat/completions и один формат tool_calls.
 *
 * capabilities заложены на будущее: когда появится insert_image,
 * панель будет показывать инструмент только там, где есть "images".
 */

export interface Provider {
  id: string;
  label: string;
  baseURL: string;
  envKey: string;
  models: string[];
  defaultModel: string;
  capabilities: Array<"chat" | "images">;
  /** Дополнительные заголовки, если провайдер их требует. */
  headers?: Record<string, string>;
  enabled: boolean;
}

export const PROVIDERS: Provider[] = [
  {
    id: "deepseek",
    label: "DeepSeek",
    baseURL: "https://api.deepseek.com",
    envKey: "DEEPSEEK_API_KEY",
    models: ["deepseek-flash", "deepseek-v4-pro"],
    defaultModel: "deepseek-flash",
    capabilities: ["chat"],
    enabled: true
  },
  {
    id: "moonshot",
    label: "Kimi",
    baseURL: "https://api.moonshot.ai/v1",
    envKey: "MOONSHOT_API_KEY",
    models: ["kimi-k3", "kimi-k2.7-code", "kimi-k2.7-code-highspeed", "kimi-k2.6"],
    defaultModel: "kimi-k3",
    capabilities: ["chat"],
    enabled: true
  },
  {
    id: "zai",
    label: "Z.AI",
    baseURL: "https://api.z.ai/api/paas/v4",
    envKey: "ZAI_API_KEY",
    models: ["glm-5.3", "glm-5.3-flash", "glm-5.2", "glm-5.1", "glm-5-turbo", "glm-5"],
    defaultModel: "glm-5.3",
    capabilities: ["chat", "images"],
    enabled: true
  },
  {
    id: "openai",
    label: "OpenAI",
    baseURL: "https://api.openai.com/v1",
    envKey: "OPENAI_API_KEY",
    models: ["gpt-6-astra", "gpt-5.6-sol", "gpt-5.6-terra", "gpt-5.6-luna", "gpt-5.5", "gpt-5.4"],
    defaultModel: "gpt-5.6-terra",
    capabilities: ["chat", "images"],
    enabled: true
  },
  {
    id: "xai",
    label: "Grok",
    baseURL: "https://api.x.ai/v1",
    envKey: "XAI_API_KEY",
    models: ["grok-4.6", "grok-4.20-0309-reasoning", "grok-4.20-0309-non-reasoning", "grok-4.5", "grok-4.3"],
    defaultModel: "grok-4.6",
    capabilities: ["chat", "images"],
    enabled: true
  },
  {
    id: "openrouter",
    label: "OpenRouter",
    baseURL: "https://openrouter.ai/api/v1",
    envKey: "OPENROUTER_API_KEY",
    models: [],
    defaultModel: "",
    capabilities: ["chat"],
    headers: {
      "HTTP-Referer": "https://localhost:3000",
      "X-Title": "Excel AI pane"
    },
    // Выключен намеренно: каталог в сотни моделей, список нужно тянуть
    // с /models и фильтровать по поддержке tool calling.
    enabled: false
  }
];

/** Списки моделей меняются чаще, чем код. Проверяйте актуальность в документации провайдера. */
export function availableProviders() {
  return PROVIDERS.filter((p) => p.enabled && !!process.env[p.envKey]?.trim()).map((p) => ({
    id: p.id,
    label: p.label,
    models: p.models,
    defaultModel: p.defaultModel,
    capabilities: p.capabilities
  }));
}

export function getProvider(id: string): Provider | undefined {
  return PROVIDERS.find((p) => p.id === id && p.enabled);
}
