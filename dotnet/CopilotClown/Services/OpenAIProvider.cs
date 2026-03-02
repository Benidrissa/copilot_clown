using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using CopilotClown.Models;

namespace CopilotClown.Services;

public class OpenAIProvider : ILlmProvider
{
    private static readonly HttpClient Http = new HttpClient { Timeout = TimeSpan.FromMinutes(3) };

    public ProviderName Name => ProviderName.OpenAI;

    public ModelInfo[] GetModels() => ModelRegistry.OpenAIModels;

    public async Task<CompletionResponse> CompleteAsync(
        string prompt, string apiKey, string model, int maxTokens = 8192, CancellationToken ct = default)
    {
        var requestBody = new Dictionary<string, object>
        {
            { "model", model },
            { "max_tokens", maxTokens },
            { "messages", new[]
                {
                    new Dictionary<string, object> { { "role", "user" }, { "content", prompt } }
                }
            }
        };

        var request = new HttpRequestMessage(HttpMethod.Post, "https://api.openai.com/v1/chat/completions")
        {
            Content = new StringContent(JsonHelper.Serialize(requestBody), Encoding.UTF8, "application/json")
        };
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

        var response = await Http.SendAsync(request, ct);

        if (!response.IsSuccessStatusCode)
        {
            var status = (int)response.StatusCode;
            var body = await response.Content.ReadAsStringAsync();
            throw status switch
            {
                401 or 403 => new ApiException("Invalid API key for OpenAI. Check your key in Settings.", status),
                429 => new ApiException("OpenAI API rate limit reached. Wait and retry.", status),
                >= 500 => new ApiException("OpenAI API is unavailable. Try again later.", status),
                _ => new ApiException($"OpenAI API error ({status}): {body}", status),
            };
        }

        var json = await response.Content.ReadAsStringAsync();
        var root = JsonHelper.Parse(json);
        var text = JsonHelper.GetString(root, "choices.0.message.content");
        var inputTokens = JsonHelper.GetInt(root, "usage.prompt_tokens");
        var outputTokens = JsonHelper.GetInt(root, "usage.completion_tokens");

        return new CompletionResponse(text, inputTokens, outputTokens);
    }

    public async Task<bool> ValidateKeyAsync(string apiKey, CancellationToken ct = default)
    {
        try
        {
            var request = new HttpRequestMessage(HttpMethod.Get, "https://api.openai.com/v1/models");
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
            var response = await Http.SendAsync(request, ct);
            return response.IsSuccessStatusCode;
        }
        catch
        {
            return false;
        }
    }
}
