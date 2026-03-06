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

public class OpenAIProvider : ILlmProvider
{
    private static readonly HttpClient Http = new HttpClient { Timeout = TimeSpan.FromMinutes(3) };

    public ProviderName Name => ProviderName.OpenAI;

    public ModelInfo[] GetModels() => ModelRegistry.OpenAIModels;

    public Task<CompletionResponse> CompleteAsync(
        string prompt, string apiKey, string model, int maxTokens = 8192, CancellationToken ct = default)
    {
        return CompleteAsync(new ResolvedPrompt(prompt), apiKey, model, maxTokens, ct);
    }

    public async Task<CompletionResponse> CompleteAsync(
        ResolvedPrompt prompt, string apiKey, string model, int maxTokens = 8192, CancellationToken ct = default)
    {
        var maxTokensKey = GetMaxTokensKey(model);

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
            { "messages", new[]
                {
                    new Dictionary<string, object> { { "role", "user" }, { "content", content } }
                }
            }
        };

        requestBody[maxTokensKey] = maxTokens;

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

    public async Task<string> UploadFileAsync(byte[] fileBytes, string fileName, string mimeType, string apiKey, CancellationToken ct = default)
    {
        var boundary = "----CopilotClown" + Guid.NewGuid().ToString("N");
        using (var formContent = new MultipartFormDataContent(boundary))
        {
            var fileContent = new ByteArrayContent(fileBytes);
            fileContent.Headers.ContentType = new MediaTypeHeaderValue(mimeType);
            formContent.Add(fileContent, "file", fileName);
            formContent.Add(new StringContent("user_data"), "purpose");

            var request = new HttpRequestMessage(HttpMethod.Post, "https://api.openai.com/v1/files")
            {
                Content = formContent
            };
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

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

    private static string GetMaxTokensKey(string model)
    {
        // GPT-5.x models expect max_completion_tokens; older models still accept max_tokens.
        return model.StartsWith("gpt-5", StringComparison.OrdinalIgnoreCase)
            ? "max_completion_tokens"
            : "max_tokens";
    }

    private static object[] BuildMultimodalContent(ResolvedPrompt prompt)
    {
        var blocks = new List<object>();

        // Document/text attachments: prefer file_id, fall back to inline text
        foreach (var att in prompt.Attachments.Where(a => a.Type == AttachmentType.Text))
        {
            if (!string.IsNullOrEmpty(att.RemoteFileId))
            {
                blocks.Add(new Dictionary<string, object>
                {
                    { "type", "file" },
                    { "file", new Dictionary<string, object>
                        {
                            { "file_id", att.RemoteFileId }
                        }
                    }
                });
            }
            else
            {
                blocks.Add(new Dictionary<string, object>
                {
                    { "type", "text" },
                    { "text", $"--- Content from {att.FileName} ---\n{att.Content}\n--- End ---\n" }
                });
            }
        }

        // Image attachments: prefer file_id, fall back to base64 data URI
        foreach (var att in prompt.Attachments.Where(a => a.Type == AttachmentType.Image))
        {
            if (!string.IsNullOrEmpty(att.RemoteFileId))
            {
                blocks.Add(new Dictionary<string, object>
                {
                    { "type", "file" },
                    { "file", new Dictionary<string, object>
                        {
                            { "file_id", att.RemoteFileId }
                        }
                    }
                });
            }
            else
            {
                blocks.Add(new Dictionary<string, object>
                {
                    { "type", "image_url" },
                    { "image_url", new Dictionary<string, object>
                        {
                            { "url", $"data:{att.MimeType};base64,{att.Content}" }
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
