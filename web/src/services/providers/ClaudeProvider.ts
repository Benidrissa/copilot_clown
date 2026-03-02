import { LLMProvider, ModelInfo, CompletionOptions, CompletionResponse, CLAUDE_MODELS } from "../../types";

const SYSTEM_PROMPT =
  "You are a helpful assistant embedded in Microsoft Excel. Respond concisely and directly. When asked to produce a list, return one item per line with no numbering or bullets unless requested. Do not add explanations unless asked.";

export class ClaudeProvider implements LLMProvider {
  readonly name = "anthropic" as const;

  getModels(): ModelInfo[] {
    return CLAUDE_MODELS;
  }

  async complete(prompt: string, apiKey: string, model: string, options?: CompletionOptions): Promise<CompletionResponse> {
    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
        "anthropic-dangerous-direct-browser-access": "true",
        "content-type": "application/json",
      },
      body: JSON.stringify({
        model,
        max_tokens: options?.maxTokens ?? 4096,
        system: options?.systemPrompt ?? SYSTEM_PROMPT,
        messages: [{ role: "user", content: prompt }],
      }),
    });

    if (!response.ok) {
      const status = response.status;
      if (status === 401 || status === 403) {
        throw new ApiError("Invalid API key for Anthropic. Check your key in the task pane.", status);
      }
      if (status === 429) {
        throw new ApiError("Anthropic API rate limit reached. Wait and retry.", status);
      }
      if (status >= 500) {
        throw new ApiError("Anthropic API is unavailable. Try again later.", status);
      }
      const body = await response.text();
      throw new ApiError(`Anthropic API error (${status}): ${body}`, status);
    }

    const data = await response.json();
    const text = data.content?.[0]?.text ?? "";
    return {
      text,
      usage: data.usage
        ? { inputTokens: data.usage.input_tokens, outputTokens: data.usage.output_tokens }
        : undefined,
    };
  }

  async validateKey(apiKey: string): Promise<boolean> {
    try {
      // Use a minimal prompt to test the key
      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "x-api-key": apiKey,
          "anthropic-version": "2023-06-01",
          "anthropic-dangerous-direct-browser-access": "true",
          "content-type": "application/json",
        },
        body: JSON.stringify({
          model: "claude-haiku-3-5-20241022",
          max_tokens: 1,
          messages: [{ role: "user", content: "hi" }],
        }),
      });
      return response.ok;
    } catch {
      return false;
    }
  }
}

export class ApiError extends Error {
  constructor(
    message: string,
    public readonly statusCode: number,
  ) {
    super(message);
    this.name = "ApiError";
  }
}
