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
    // Каталог OpenRouter — сотни моделей, и поддержка функций там заявлена
    // шире, чем работает на деле. Этот список отобран вручную 16 сентября
    // 2026 года: каждая модель вызвана живым запросом с tool_choice required
    // и действительно вернула вызов инструмента. Отсеяны: gemini-3.5-flash
    // и gemini-3.1-pro-preview отвечали текстом даже принудительно,
    // thinkingmachines/inkling:free отдаёт 403 вне своих клиентов,
    // nex-n2.5-mini:free не ответил за минуту. Варианты :batch не включены:
    // они асинхронные и для интерактивной панели не годятся.
    models: [
      // Claude
      "anthropic/claude-opus-5",
      "anthropic/claude-sonnet-5",
      "anthropic/claude-haiku-4.5",
      "anthropic/claude-fable-5.1",
      // Gemini
      "google/gemini-3.8-flash",
      "google/gemini-3.7-flash",
      "google/gemini-3.5-flash-lite",
      "google/gemini-2.5-flash",
      // Mistral
      "mistralai/mistral-large-2512",
      "mistralai/mistral-medium-3-5",
      "mistralai/mistral-small-2603",
      "mistralai/ministral-8b-2512",
      // Бесплатные
      "nvidia/nemotron-3.5-lightning:free",
      "nex-agi/nex-n2.5-pro:free",
      "dots-studio/dots-3-note-preview:free",
      "inclusionai/ling-3.0-flash-vl:free"
    ],
    defaultModel: "anthropic/claude-sonnet-5",
    capabilities: ["chat"],
    headers: {
      "HTTP-Referer": "https://localhost:3000",
      "X-Title": "Excel AI pane"
    },
    enabled: true
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
