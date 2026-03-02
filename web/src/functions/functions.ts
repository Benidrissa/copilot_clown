import { CacheService } from "../services/CacheService";
import { SettingsService } from "../services/SettingsService";
import { PromptBuilder } from "../services/PromptBuilder";
import { RateLimiter } from "../services/RateLimiter";
import { getProvider } from "../services/providers/LLMProvider";

// Shared instances (available across the shared runtime)
const cacheService = new CacheService();
const settingsService = new SettingsService();
const promptBuilder = new PromptBuilder();
const rateLimiter = new RateLimiter();

// Expose services to the task pane via window globals
declare global {
  interface Window {
    aillmServices: {
      cacheService: CacheService;
      settingsService: SettingsService;
      rateLimiter: RateLimiter;
    };
  }
}
window.aillmServices = { cacheService, settingsService, rateLimiter };

/**
 * Calls an AI language model with the given prompt and context.
 * @customfunction USEAI
 * @param {any[]} args Alternating prompt parts and context ranges
 * @returns {Promise<string | string[][]>} AI-generated response
 * @helpurl https://github.com/copilot-clown
 */
async function useai(...args: unknown[]): Promise<string | string[][]> {
  // 1. Build prompt
  const prompt = promptBuilder.buildPrompt(args);
  if (!prompt || prompt.trim().length === 0) {
    return "Error: Prompt cannot be empty.";
  }
  if (prompt.length > 100_000) {
    return "Error: Prompt too large. Reduce prompt or context size.";
  }

  // 2. Get settings
  const settings = settingsService.getSettings();
  const { activeProvider, activeModel } = settings;

  // 3. Check API key
  const apiKey = settingsService.getApiKey(activeProvider);
  if (!apiKey) {
    return `Error: No API key configured for ${activeProvider}. Open the task pane to add your key.`;
  }

  // 4. Check cache
  if (settings.cacheEnabled) {
    const cached = await cacheService.get(activeProvider, activeModel, prompt);
    if (cached !== null) {
      return parseResponse(cached);
    }
  }

  // 5. Check rate limit
  rateLimiter.updateLimits(settings.rateLimitMax, settings.rateLimitWindowMs);
  if (!rateLimiter.tryAcquire()) {
    return "Error: Rate limit exceeded. Wait and try again.";
  }

  // 6. Call API
  const provider = getProvider(activeProvider);
  try {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), settings.apiTimeoutMs);

    const result = await provider.complete(prompt, apiKey, activeModel);
    clearTimeout(timeout);

    // 7. Cache the result
    if (settings.cacheEnabled) {
      await cacheService.set(activeProvider, activeModel, prompt, result.text);
    }

    return parseResponse(result.text);
  } catch (error: unknown) {
    if (error instanceof Error) {
      if (error.name === "AbortError") {
        return "Error: Request timed out. Try a simpler prompt or check your connection.";
      }
      return `Error: ${error.message}`;
    }
    return "Error: Unexpected error occurred.";
  }
}

/**
 * Parse AI response text into a single value or spill array.
 * If the response contains multiple lines, it returns a 2D array for spilling.
 */
function parseResponse(text: string): string | string[][] {
  const trimmed = text.trim();
  const lines = trimmed.split("\n").filter((line) => line.trim().length > 0);

  if (lines.length <= 1) {
    return trimmed;
  }

  // Return as a column array for spilling
  return lines.map((line) => [line.trim()]);
}

// Register the function with the CustomFunctions runtime
CustomFunctions.associate("USEAI", useai);
