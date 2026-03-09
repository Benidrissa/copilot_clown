using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using CopilotClown.Models;

namespace CopilotClown.Services;

public class GeminiProvider : ILlmProvider
{
    private static readonly HttpClient Http = new HttpClient { Timeout = TimeSpan.FromMinutes(3) };
    private const string BaseUrl = "https://generativelanguage.googleapis.com/v1beta/models/";

    public ProviderName Name => ProviderName.Google;

    public ModelInfo[] GetModels() => ModelRegistry.GeminiModels;

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
        var parts = BuildParts(prompt);

        var requestBody = new Dictionary<string, object>
        {
            { "contents", new[] { new Dictionary<string, object> { { "role", "user" }, { "parts", parts } } } },
            { "generationConfig", new Dictionary<string, object>
                {
                    { "maxOutputTokens", settings.MaxTokens },
                    { "temperature", settings.Temperature },
                    { "topP", settings.TopP }
                }
            }
        };

        if (!string.IsNullOrWhiteSpace(settings.SystemPrompt))
        {
            requestBody["systemInstruction"] = new Dictionary<string, object>
            {
                { "parts", new[] { new Dictionary<string, object> { { "text", settings.SystemPrompt } } } }
            };
        }

        var url = $"{BaseUrl}{model}:generateContent";
        var request = new HttpRequestMessage(HttpMethod.Post, url)
        {
            Content = new StringContent(JsonHelper.Serialize(requestBody), Encoding.UTF8, "application/json")
        };
        request.Headers.Add("x-goog-api-key", apiKey);

        var response = await Http.SendAsync(request, ct);

        if (!response.IsSuccessStatusCode)
        {
            var status = (int)response.StatusCode;
            var body = await response.Content.ReadAsStringAsync();
            throw status switch
            {
                401 or 403 => new ApiException("Invalid API key for Google Gemini. Check your key in Settings.", status),
                429 => new ApiException("Google Gemini API rate limit reached. Wait and retry.", status),
                >= 500 => new ApiException("Google Gemini API is unavailable. Try again later.", status),
                _ => new ApiException($"Gemini API error ({status}): {body}", status),
            };
        }

        var json = await response.Content.ReadAsStringAsync();
        var root = JsonHelper.Parse(json);
        var text = JsonHelper.GetString(root, "candidates.0.content.parts.0.text");
        var inputTokens = JsonHelper.GetInt(root, "usageMetadata.promptTokenCount");
        var outputTokens = JsonHelper.GetInt(root, "usageMetadata.candidatesTokenCount");

        return new CompletionResponse(text ?? "", inputTokens, outputTokens);
    }

    public async Task<string> UploadFileAsync(byte[] fileBytes, string fileName, string mimeType, string apiKey, CancellationToken ct = default)
    {
        // Gemini Files API: media.upload
        var url = $"https://generativelanguage.googleapis.com/upload/v1beta/files";
        var boundary = "----CopilotClown" + Guid.NewGuid().ToString("N");
        using (var formContent = new MultipartFormDataContent(boundary))
        {
            var metadata = JsonHelper.Serialize(new Dictionary<string, object>
            {
                { "file", new Dictionary<string, object> { { "display_name", fileName } } }
            });
            formContent.Add(new StringContent(metadata, Encoding.UTF8, "application/json"), "metadata");

            var fileContent = new ByteArrayContent(fileBytes);
            fileContent.Headers.ContentType = new MediaTypeHeaderValue(mimeType);
            formContent.Add(fileContent, "file", fileName);

            var request = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = formContent
            };
            request.Headers.Add("x-goog-api-key", apiKey);

            var response = await Http.SendAsync(request, ct);
            response.EnsureSuccessStatusCode();

            var json = await response.Content.ReadAsStringAsync();
            var root = JsonHelper.Parse(json);
            return JsonHelper.GetString(root, "file.name");
        }
    }

    public async Task<bool> ValidateKeyAsync(string apiKey, CancellationToken ct = default)
    {
        try
        {
            var url = $"{BaseUrl}gemini-2.0-flash-lite:generateContent";
            var requestBody = new Dictionary<string, object>
            {
                { "contents", new[] { new Dictionary<string, object>
                    {
                        { "role", "user" },
                        { "parts", new[] { new Dictionary<string, object> { { "text", "hi" } } } }
                    }
                }},
                { "generationConfig", new Dictionary<string, object> { { "maxOutputTokens", 1 } } }
            };

            var request = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new StringContent(JsonHelper.Serialize(requestBody), Encoding.UTF8, "application/json")
            };
            request.Headers.Add("x-goog-api-key", apiKey);

            var response = await Http.SendAsync(request, ct);
            return response.IsSuccessStatusCode;
        }
        catch
        {
            return false;
        }
    }

    private static object[] BuildParts(ResolvedPrompt prompt)
    {
        var parts = new List<object>();

        if (!prompt.HasAttachments)
        {
            parts.Add(new Dictionary<string, object> { { "text", prompt.CleanText } });
            return parts.ToArray();
        }

        // Text attachments: prepend as context
        foreach (var att in prompt.Attachments.Where(a => a.Type == AttachmentType.Text))
        {
            if (!string.IsNullOrEmpty(att.RemoteFileId))
            {
                parts.Add(new Dictionary<string, object>
                {
                    { "fileData", new Dictionary<string, object>
                        {
                            { "fileUri", att.RemoteFileId },
                            { "mimeType", att.MimeType }
                        }
                    }
                });
            }
            else
            {
                parts.Add(new Dictionary<string, object>
                {
                    { "text", $"--- Content from {att.FileName} ---\n{att.Content}\n--- End ---\n" }
                });
            }
        }

        // Image attachments: inline as base64
        foreach (var att in prompt.Attachments.Where(a => a.Type == AttachmentType.Image))
        {
            if (!string.IsNullOrEmpty(att.RemoteFileId))
            {
                parts.Add(new Dictionary<string, object>
                {
                    { "fileData", new Dictionary<string, object>
                        {
                            { "fileUri", att.RemoteFileId },
                            { "mimeType", att.MimeType }
                        }
                    }
                });
            }
            else
            {
                parts.Add(new Dictionary<string, object>
                {
                    { "inlineData", new Dictionary<string, object>
                        {
                            { "mimeType", att.MimeType },
                            { "data", att.Content }
                        }
                    }
                });
            }
        }

        // User prompt text last
        parts.Add(new Dictionary<string, object> { { "text", prompt.CleanText } });

        return parts.ToArray();
    }
}
