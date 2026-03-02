import { LLMProvider, ModelInfo, CompletionOptions, CompletionResponse, OPENAI_MODELS } from "../../types";

const SYSTEM_PROMPT =
  "You are a helpful assistant embedded in Microsoft Excel. Respond concisely and directly. When asked to produce a list, return one item per line with no numbering or bullets unless requested. Do not add explanations unless asked.";

export class OpenAIProvider implements LLMProvider {
  readonly name = "openai" as const;

  getModels(): ModelInfo[] {
    return OPENAI_MODELS;
  }

  async complete(prompt: string, apiKey: string, model: string, options?: CompletionOptions): Promise<CompletionResponse> {
    const response = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${apiKey}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model,
        max_tokens: options?.maxTokens ?? 4096,
        messages: [
          { role: "system", content: options?.systemPrompt ?? SYSTEM_PROMPT },
          { role: "user", content: prompt },
        ],
      }),
    });

    if (!response.ok) {
      const status = response.status;
      if (status === 401 || status === 403) {
        throw new ApiError("Invalid API key for OpenAI. Check your key in the task pane.", status);
      }
      if (status === 429) {
        throw new ApiError("OpenAI API rate limit reached. Wait and retry.", status);
      }
      if (status >= 500) {
        throw new ApiError("OpenAI API is unavailable. Try again later.", status);
      }
      const body = await response.text();
      throw new ApiError(`OpenAI API error (${status}): ${body}`, status);
    }

    const data = await response.json();
    const text = data.choices?.[0]?.message?.content ?? "";
    return {
      text,
      usage: data.usage
        ? { inputTokens: data.usage.prompt_tokens, outputTokens: data.usage.completion_tokens }
        : undefined,
    };
  }

  async validateKey(apiKey: string): Promise<boolean> {
    try {
      const response = await fetch("https://api.openai.com/v1/models", {
        headers: { Authorization: `Bearer ${apiKey}` },
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
