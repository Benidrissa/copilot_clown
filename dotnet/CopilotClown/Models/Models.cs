using System;
using System.Linq;

namespace CopilotClown.Models;

public enum ProviderName
{
    Anthropic,
    OpenAI,
    Google
}

public class ModelInfo
{
    public string Id { get; }
    public string DisplayName { get; }
    public ProviderName Provider { get; }
    public int ContextWindow { get; }
    public string PricingTier { get; }

    public ModelInfo(string id, string displayName, ProviderName provider, int contextWindow, string pricingTier)
    {
        Id = id;
        DisplayName = displayName;
        Provider = provider;
        ContextWindow = contextWindow;
        PricingTier = pricingTier;
    }
}

public class CompletionResponse
{
    public string Text { get; }
    public int InputTokens { get; }
    public int OutputTokens { get; }

    public CompletionResponse(string text, int inputTokens = 0, int outputTokens = 0)
    {
        Text = text;
        InputTokens = inputTokens;
        OutputTokens = outputTokens;
    }
}

public class AppSettings
{
    public ProviderName ActiveProvider { get; set; } = ProviderName.OpenAI;
    public string ActiveModel { get; set; } = "gpt-5.2";
    public bool CacheEnabled { get; set; } = true;
    public int CacheTtlMinutes { get; set; } = 1440; // 24 hours
    public int CacheMaxEntries { get; set; } = 1000;
    public int RateLimitMax { get; set; } = 500;
    public int RateLimitWindowMinutes { get; set; } = 10;
    public int ApiTimeoutSeconds { get; set; } = 30;
    public string SystemPrompt { get; set; } = "";
    public double Temperature { get; set; } = 1.0;
    public int MaxTokens { get; set; } = 8192;
    public double TopP { get; set; } = 1.0;
}

public static class ModelRegistry
{
    public static readonly ModelInfo[] ClaudeModels = new ModelInfo[]
    {
        new ModelInfo("claude-opus-4-6", "Claude Opus 4.6", ProviderName.Anthropic, 200_000, "High"),
        new ModelInfo("claude-opus-4-5-20251101", "Claude Opus 4.5", ProviderName.Anthropic, 200_000, "High"),
        new ModelInfo("claude-opus-4-1-20250805", "Claude Opus 4.1", ProviderName.Anthropic, 200_000, "High"),
        new ModelInfo("claude-opus-4-20250514", "Claude Opus 4", ProviderName.Anthropic, 200_000, "High"),
        new ModelInfo("claude-sonnet-4-6", "Claude Sonnet 4.6", ProviderName.Anthropic, 200_000, "Medium"),
        new ModelInfo("claude-sonnet-4-5-20250929", "Claude Sonnet 4.5", ProviderName.Anthropic, 200_000, "Medium"),
        new ModelInfo("claude-sonnet-4-20250514", "Claude Sonnet 4", ProviderName.Anthropic, 200_000, "Medium"),
        new ModelInfo("claude-haiku-4-5-20251001", "Claude Haiku 4.5", ProviderName.Anthropic, 200_000, "Low"),
        new ModelInfo("claude-3-haiku-20240307", "Claude 3 Haiku", ProviderName.Anthropic, 200_000, "Low"),
    };

    public static readonly ModelInfo[] OpenAIModels = new ModelInfo[]
    {
        // GPT-5 series
        new ModelInfo("gpt-5.2", "GPT-5.2", ProviderName.OpenAI, 200_000, "High"),
        new ModelInfo("gpt-5.1", "GPT-5.1", ProviderName.OpenAI, 200_000, "High"),
        new ModelInfo("gpt-5", "GPT-5", ProviderName.OpenAI, 200_000, "High"),
        new ModelInfo("gpt-5-mini", "GPT-5 Mini", ProviderName.OpenAI, 200_000, "Medium"),
        new ModelInfo("gpt-5-nano", "GPT-5 Nano", ProviderName.OpenAI, 200_000, "Low"),
        new ModelInfo("gpt-5-pro", "GPT-5 Pro", ProviderName.OpenAI, 200_000, "High"),
        new ModelInfo("gpt-5.2-pro", "GPT-5.2 Pro", ProviderName.OpenAI, 200_000, "High"),
        // GPT-4 series
        new ModelInfo("gpt-4.1", "GPT-4.1", ProviderName.OpenAI, 1_047_576, "High"),
        new ModelInfo("gpt-4.1-mini", "GPT-4.1 Mini", ProviderName.OpenAI, 1_047_576, "Medium"),
        new ModelInfo("gpt-4.1-nano", "GPT-4.1 Nano", ProviderName.OpenAI, 1_047_576, "Low"),
        new ModelInfo("gpt-4o", "GPT-4o", ProviderName.OpenAI, 128_000, "High"),
        new ModelInfo("gpt-4o-mini", "GPT-4o Mini", ProviderName.OpenAI, 128_000, "Medium"),
        new ModelInfo("gpt-4-turbo", "GPT-4 Turbo", ProviderName.OpenAI, 128_000, "High"),
        // Reasoning models
        new ModelInfo("o4-mini", "o4-mini", ProviderName.OpenAI, 200_000, "Medium"),
        new ModelInfo("o3", "o3", ProviderName.OpenAI, 200_000, "High"),
        new ModelInfo("o3-mini", "o3-mini", ProviderName.OpenAI, 200_000, "Medium"),
        new ModelInfo("o1", "o1", ProviderName.OpenAI, 200_000, "High"),
        new ModelInfo("o1-pro", "o1 Pro", ProviderName.OpenAI, 200_000, "High"),
        // Legacy
        new ModelInfo("gpt-3.5-turbo", "GPT-3.5 Turbo", ProviderName.OpenAI, 16_385, "Low"),
    };

    public static readonly ModelInfo[] GeminiModels = new ModelInfo[]
    {
        new ModelInfo("gemini-2.5-pro-preview-06-05", "Gemini 2.5 Pro", ProviderName.Google, 1_048_576, "Medium"),
        new ModelInfo("gemini-2.5-flash-preview-05-20", "Gemini 2.5 Flash", ProviderName.Google, 1_048_576, "Low"),
        new ModelInfo("gemini-2.0-flash", "Gemini 2.0 Flash", ProviderName.Google, 1_048_576, "Low"),
        new ModelInfo("gemini-2.0-flash-lite", "Gemini 2.0 Flash Lite", ProviderName.Google, 1_048_576, "Low"),
    };

    public static ModelInfo[] GetModels(ProviderName provider)
    {
        switch (provider)
        {
            case ProviderName.Anthropic: return ClaudeModels;
            case ProviderName.Google: return GeminiModels;
            default: return OpenAIModels;
        }
    }

    public static ModelInfo[] AllModels => ClaudeModels.Concat(OpenAIModels).Concat(GeminiModels).ToArray();
}
