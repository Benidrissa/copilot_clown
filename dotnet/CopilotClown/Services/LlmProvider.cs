using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using CopilotClown.Models;

namespace CopilotClown.Services;

public interface ILlmProvider
{
    ProviderName Name { get; }
    ModelInfo[] GetModels();
    Task<CompletionResponse> CompleteAsync(string prompt, string apiKey, string model, int maxTokens = 8192, CancellationToken ct = default);
    Task<CompletionResponse> CompleteAsync(ResolvedPrompt prompt, string apiKey, string model, int maxTokens = 8192, CancellationToken ct = default);
    Task<CompletionResponse> CompleteAsync(ResolvedPrompt prompt, string apiKey, string model, AppSettings settings, CancellationToken ct = default);
    Task<string> UploadFileAsync(byte[] fileBytes, string fileName, string mimeType, string apiKey, CancellationToken ct = default);
    Task<bool> ValidateKeyAsync(string apiKey, CancellationToken ct = default);
}

public static class ProviderFactory
{
    private static readonly Dictionary<ProviderName, ILlmProvider> Providers = new Dictionary<ProviderName, ILlmProvider>
    {
        [ProviderName.Anthropic] = new ClaudeProvider(),
        [ProviderName.OpenAI] = new OpenAIProvider(),
        [ProviderName.Google] = new GeminiProvider(),
    };

    public static ILlmProvider Get(ProviderName name) => Providers[name];
}
