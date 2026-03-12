using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using CopilotClown.Models;

namespace CopilotClown.Services;

public class ClaudeProvider : ILlmProvider
{
    private static readonly HttpClient Http = new HttpClient { Timeout = TimeSpan.FromMinutes(5) };

    public ProviderName Name => ProviderName.Anthropic;

    public ModelInfo[] GetModels() => ModelRegistry.ClaudeModels;

    public Task<CompletionResponse> CompleteAsync(
        string prompt, string apiKey, string model, int maxTokens = 8192, CancellationToken ct = default)
    {
        return CompleteAsync(new ResolvedPrompt(prompt), apiKey, model, maxTokens, ct);
    }

    public async Task<CompletionResponse> CompleteAsync(
        ResolvedPrompt prompt, string apiKey, string model, int maxTokens = 8192, CancellationToken ct = default)
    {
        return await CompleteAsync(prompt, apiKey, model, new AppSettings { MaxTokens = maxTokens }, ct);
    }

    public async Task<CompletionResponse> CompleteAsync(
        ResolvedPrompt prompt, string apiKey, string model, AppSettings settings, CancellationToken ct = default)
    {
        var maxTokens = settings.MaxTokens;
        object content;

        if (!prompt.HasAttachments)
        {
            content = prompt.CleanText;
        }
        else
        {
            content = BuildMultimodalContent(prompt);
        }

        var requestBody = new Dictionary<string, object>
        {
            { "model", model },
            { "max_tokens", maxTokens },
            { "messages", new[] { new Dictionary<string, object> { { "role", "user" }, { "content", content } } } }
        };

        // Anthropic: temperature and top_p are mutually exclusive — send only one.
        // Prefer temperature unless user left it at default and changed top_p.
        if (Math.Abs(settings.TopP - 1.0) > 0.001 && Math.Abs(settings.Temperature - 1.0) < 0.001)
            requestBody["top_p"] = settings.TopP;
        else
            requestBody["temperature"] = settings.Temperature;

        if (!string.IsNullOrWhiteSpace(settings.SystemPrompt))
            requestBody["system"] = settings.SystemPrompt;

        var request = new HttpRequestMessage(HttpMethod.Post, "https://api.anthropic.com/v1/messages")
        {
            Content = new StringContent(JsonHelper.Serialize(requestBody), Encoding.UTF8, "application/json")
        };
        request.Headers.Add("x-api-key", apiKey);
        request.Headers.Add("anthropic-version", "2023-06-01");
        // Required for file_id references in content blocks
        request.Headers.Add("anthropic-beta", "files-api-2025-04-14");

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

    public async Task<string> UploadFileAsync(byte[] fileBytes, string fileName, string mimeType, string apiKey, CancellationToken ct = default)
    {
        var boundary = "----CopilotClown" + Guid.NewGuid().ToString("N");
        using (var formContent = new MultipartFormDataContent(boundary))
        {
            var fileContent = new ByteArrayContent(fileBytes);
            fileContent.Headers.ContentType = new MediaTypeHeaderValue(mimeType);
            formContent.Add(fileContent, "file", fileName);

            var request = new HttpRequestMessage(HttpMethod.Post, "https://api.anthropic.com/v1/files")
            {
                Content = formContent
            };
            request.Headers.Add("x-api-key", apiKey);
            request.Headers.Add("anthropic-version", "2023-06-01");
            request.Headers.Add("anthropic-beta", "files-api-2025-04-14");

            var response = await Http.SendAsync(request, ct);
            response.EnsureSuccessStatusCode();

            var json = await response.Content.ReadAsStringAsync();
            var root = JsonHelper.Parse(json);
            return JsonHelper.GetString(root, "id");
        }
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

    private static object[] BuildMultimodalContent(ResolvedPrompt prompt)
    {
        var blocks = new List<object>();
        var cacheControl = new Dictionary<string, object> { { "type", "ephemeral" } };

        // Document attachments: prefer file_id, fall back to inline text
        foreach (var att in prompt.Attachments.Where(a => a.Type == AttachmentType.Text))
        {
            if (!string.IsNullOrEmpty(att.RemoteFileId))
            {
                // Use API-uploaded file reference (avoids re-sending content)
                blocks.Add(new Dictionary<string, object>
                {
                    { "type", "document" },
                    { "source", new Dictionary<string, object>
                        {
                            { "type", "file" },
                            { "file_id", att.RemoteFileId }
                        }
                    },
                    { "cache_control", cacheControl }
                });
            }
            else
            {
                // Inline text fallback
                blocks.Add(new Dictionary<string, object>
                {
                    { "type", "text" },
                    { "text", $"--- Content from {att.FileName} ---\n{att.Content}\n--- End ---\n" },
                    { "cache_control", cacheControl }
                });
            }
        }

        // Image attachments: prefer file_id, fall back to base64
        foreach (var att in prompt.Attachments.Where(a => a.Type == AttachmentType.Image))
        {
            if (!string.IsNullOrEmpty(att.RemoteFileId))
            {
                blocks.Add(new Dictionary<string, object>
                {
                    { "type", "image" },
                    { "source", new Dictionary<string, object>
                        {
                            { "type", "file" },
                            { "file_id", att.RemoteFileId }
                        }
                    },
                    { "cache_control", cacheControl }
                });
            }
            else
            {
                blocks.Add(new Dictionary<string, object>
                {
                    { "type", "image" },
                    { "source", new Dictionary<string, object>
                        {
                            { "type", "base64" },
                            { "media_type", att.MimeType },
                            { "data", att.Content }
                        }
                    }
                });
            }
        }

        // User prompt text (always last)
        blocks.Add(new Dictionary<string, object>
        {
            { "type", "text" },
            { "text", prompt.CleanText }
        });

        return blocks.ToArray();
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
