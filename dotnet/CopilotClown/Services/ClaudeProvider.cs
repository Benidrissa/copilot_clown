using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using CopilotClown.Models;

namespace CopilotClown.Services;

public class ClaudeProvider : ILlmProvider
{
    private static readonly HttpClient Http = new HttpClient();

    public ProviderName Name => ProviderName.Anthropic;

    public ModelInfo[] GetModels() => ModelRegistry.ClaudeModels;

    public async Task<CompletionResponse> CompleteAsync(
        string prompt, string apiKey, string model, int maxTokens = 8192, CancellationToken ct = default)
    {
        var requestBody = new Dictionary<string, object>
        {
            { "model", model },
            { "max_tokens", maxTokens },
            { "messages", new[] { new Dictionary<string, object> { { "role", "user" }, { "content", prompt } } } }
        };

        var request = new HttpRequestMessage(HttpMethod.Post, "https://api.anthropic.com/v1/messages")
        {
            Content = new StringContent(JsonHelper.Serialize(requestBody), Encoding.UTF8, "application/json")
        };
        request.Headers.Add("x-api-key", apiKey);
        request.Headers.Add("anthropic-version", "2023-06-01");

        var response = await Http.SendAsync(request, ct);

        if (!response.IsSuccessStatusCode)
        {
            var status = (int)response.StatusCode;
            var body = await response.Content.ReadAsStringAsync();
            throw status switch
            {
                401 or 403 => new ApiException("Invalid API key for Anthropic. Check your key in Settings.", status),
                429 => new ApiException("Anthropic API rate limit reached. Wait and retry.", status),
                >= 500 => new ApiException("Anthropic API is unavailable. Try again later.", status),
                _ => new ApiException($"Anthropic API error ({status}): {body}", status),
            };
        }

        var json = await response.Content.ReadAsStringAsync();
        var root = JsonHelper.Parse(json);
        var text = JsonHelper.GetString(root, "content.0.text");
        var inputTokens = JsonHelper.GetInt(root, "usage.input_tokens");
        var outputTokens = JsonHelper.GetInt(root, "usage.output_tokens");

        return new CompletionResponse(text, inputTokens, outputTokens);
    }

    public async Task<bool> ValidateKeyAsync(string apiKey, CancellationToken ct = default)
    {
        try
        {
            await CompleteAsync("hi", apiKey, "claude-haiku-3-5-20241022", maxTokens: 1, ct: ct);
            return true;
        }
        catch
        {
            return false;
        }
    }
}

public class ApiException : Exception
{
    public int StatusCode { get; }

    public ApiException(string message, int statusCode) : base(message)
    {
        StatusCode = statusCode;
    }
}
