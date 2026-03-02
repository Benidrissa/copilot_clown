// ── Provider & Model Types ──────────────────────────────────────────

export type ProviderName = "anthropic" | "openai";

export interface ModelInfo {
  id: string;
  displayName: string;
  provider: ProviderName;
  contextWindow: number;
  pricingTier: "low" | "medium" | "high";
}

export interface CompletionOptions {
  maxTokens?: number;
  temperature?: number;
  systemPrompt?: string;
}

export interface CompletionResponse {
  text: string;
  usage?: {
    inputTokens: number;
    outputTokens: number;
  };
}

export interface LLMProvider {
  readonly name: ProviderName;
  getModels(): ModelInfo[];
  complete(prompt: string, apiKey: string, model: string, options?: CompletionOptions): Promise<CompletionResponse>;
  validateKey(apiKey: string): Promise<boolean>;
}

// ── Cache Types ─────────────────────────────────────────────────────

export interface CacheEntry {
  response: string;
  timestamp: number;
  provider: ProviderName;
  model: string;
  promptPreview: string;
}

export interface CacheStats {
  totalEntries: number;
  totalSizeBytes: number;
  hitRate: number;
  hits: number;
  misses: number;
}

// ── Settings Types ──────────────────────────────────────────────────

export interface AppSettings {
  activeProvider: ProviderName;
  activeModel: string;
  cacheEnabled: boolean;
  cacheTtlMs: number;
  cacheMaxEntries: number;
  rateLimitMax: number;
  rateLimitWindowMs: number;
  apiTimeoutMs: number;
}

export const DEFAULT_SETTINGS: AppSettings = {
  activeProvider: "openai",
  activeModel: "gpt-4.1-mini",
  cacheEnabled: true,
  cacheTtlMs: 24 * 60 * 60 * 1000, // 24 hours
  cacheMaxEntries: 1000,
  rateLimitMax: 100,
  rateLimitWindowMs: 10 * 60 * 1000, // 10 minutes
  apiTimeoutMs: 30_000,
};

// ── Model Registry ──────────────────────────────────────────────────

export const CLAUDE_MODELS: ModelInfo[] = [
  { id: "claude-opus-4-20250514", displayName: "Claude Opus 4", provider: "anthropic", contextWindow: 200000, pricingTier: "high" },
  { id: "claude-sonnet-4-20250514", displayName: "Claude Sonnet 4", provider: "anthropic", contextWindow: 200000, pricingTier: "medium" },
  { id: "claude-haiku-3-5-20241022", displayName: "Claude Haiku 3.5", provider: "anthropic", contextWindow: 200000, pricingTier: "low" },
];

export const OPENAI_MODELS: ModelInfo[] = [
  { id: "gpt-4.1", displayName: "GPT-4.1", provider: "openai", contextWindow: 1047576, pricingTier: "high" },
  { id: "gpt-4.1-mini", displayName: "GPT-4.1 Mini", provider: "openai", contextWindow: 1047576, pricingTier: "medium" },
  { id: "gpt-4.1-nano", displayName: "GPT-4.1 Nano", provider: "openai", contextWindow: 1047576, pricingTier: "low" },
  { id: "gpt-4o", displayName: "GPT-4o", provider: "openai", contextWindow: 128000, pricingTier: "high" },
  { id: "gpt-4o-mini", displayName: "GPT-4o Mini", provider: "openai", contextWindow: 128000, pricingTier: "medium" },
  { id: "o3-mini", displayName: "o3-mini", provider: "openai", contextWindow: 200000, pricingTier: "medium" },
];

export const ALL_MODELS: ModelInfo[] = [...CLAUDE_MODELS, ...OPENAI_MODELS];

export function getModelsForProvider(provider: ProviderName): ModelInfo[] {
  return provider === "anthropic" ? CLAUDE_MODELS : OPENAI_MODELS;
}
