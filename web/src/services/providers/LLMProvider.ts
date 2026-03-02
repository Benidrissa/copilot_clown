// Re-export the interface from types — this file serves as the barrel for providers
export type { LLMProvider, CompletionOptions, CompletionResponse, ModelInfo, ProviderName } from "../../types";

export { ClaudeProvider } from "./ClaudeProvider";
export { OpenAIProvider } from "./OpenAIProvider";

import { ProviderName, LLMProvider } from "../../types";
import { ClaudeProvider } from "./ClaudeProvider";
import { OpenAIProvider } from "./OpenAIProvider";

const providers: Record<ProviderName, LLMProvider> = {
  anthropic: new ClaudeProvider(),
  openai: new OpenAIProvider(),
};

export function getProvider(name: ProviderName): LLMProvider {
  return providers[name];
}
