using System;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;
using CopilotClown.Models;
using CopilotClown.Services;

namespace CopilotClown.Functions;

public static class UseAiFunction
{
    private static readonly CacheService Cache = new CacheService();
    private static readonly SettingsService Settings = new SettingsService();
    private static readonly RateLimiter RateLimiter = new RateLimiter();

    // Force TLS 1.2 — required by both Anthropic and OpenAI APIs.
    // .NET Framework 4.8 doesn't always enable this by default.
    static UseAiFunction()
    {
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13;
    }

    // Expose services for the settings UI
    internal static CacheService CacheInstance => Cache;
    internal static SettingsService SettingsInstance => Settings;
    internal static RateLimiter RateLimiterInstance => RateLimiter;

    [ExcelFunction(
        Name = "USEAI",
        Description = "Calls an AI language model (Claude or OpenAI) with the given prompt and context. Configure API keys via the ribbon Settings button.",
        HelpTopic = "https://github.com/Benidrissa/copilot_clown")]
    public static object UseAi(
        [ExcelArgument(Name = "prompt_part1", Description = "Text describing the task or question")] object arg1,
        [ExcelArgument(Name = "context1", Description = "[Optional] Cell reference or range providing context")] object arg2,
        [ExcelArgument(Name = "prompt_part2", Description = "[Optional] Additional prompt text")] object arg3,
        [ExcelArgument(Name = "context2", Description = "[Optional] Additional context")] object arg4,
        [ExcelArgument(Name = "prompt_part3", Description = "[Optional] Additional prompt text")] object arg5,
        [ExcelArgument(Name = "context3", Description = "[Optional] Additional context")] object arg6,
        [ExcelArgument(Name = "prompt_part4", Description = "[Optional] Additional prompt text")] object arg7,
        [ExcelArgument(Name = "context4", Description = "[Optional] Additional context")] object arg8)
    {
        var args = new[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8 };

        // Build prompt
        var prompt = PromptBuilder.Build(args);
        if (string.IsNullOrWhiteSpace(prompt))
            return "Error: Prompt cannot be empty.";
        if (prompt.Length > 100_000)
            return "Error: Prompt too large. Reduce prompt or context size.";

        // Load settings
        var settings = Settings.LoadSettings();
        var provider = settings.ActiveProvider;
        var model = settings.ActiveModel;

        // Check API key
        var apiKey = Settings.GetApiKey(provider);
        if (string.IsNullOrEmpty(apiKey))
            return $"Error: No API key configured for {provider}. Click 'AI Settings' in the ribbon.";

        // Check cache
        if (settings.CacheEnabled)
        {
            var cached = Cache.Get(provider, model, prompt);
            if (cached != null)
                return ParseResponse(cached);
        }

        // Use ExcelAsyncUtil for async API call
        var cacheKey = $"{provider}|{model}|{prompt}";
        return ExcelAsyncUtil.Run("USEAI", cacheKey, () =>
        {
            // Rate limit
            RateLimiter.UpdateLimits(settings.RateLimitMax, settings.RateLimitWindowMinutes);
            if (!RateLimiter.TryAcquire())
                return (object)"Error: Rate limit exceeded. Wait and try again.";

            try
            {
                var llm = ProviderFactory.Get(provider);
                var result = llm.CompleteAsync(prompt, apiKey, model, ct: CancellationToken.None)
                    .GetAwaiter().GetResult();

                // Cache the result
                if (settings.CacheEnabled)
                    Cache.Set(provider, model, prompt, result.Text, settings.CacheTtlMinutes);

                return ParseResponse(result.Text);
            }
            catch (ApiException ex)
            {
                return (object)$"Error: {ex.Message}";
            }
            catch (TaskCanceledException)
            {
                return (object)"Error: Request timed out. Try a simpler prompt or check your connection.";
            }
            catch (HttpRequestException)
            {
                return (object)"Error: Network error. Check your internet connection.";
            }
            catch (Exception ex)
            {
                return (object)$"Error: {ex.Message}";
            }
        });
    }

    /// <summary>
    /// Return full response as a single cell value. No splitting, no truncation.
    /// </summary>
    private static object ParseResponse(string text)
    {
        return text.Trim();
    }
}
