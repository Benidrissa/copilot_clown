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
    private static readonly HttpClient Http = new HttpClient { Timeout = TimeSpan.FromMinutes(5) };

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
        return await CompleteAsync(prompt, apiKey, model, new AppSettings { MaxTokens = maxTokens }, ct);
    }

    public async Task<CompletionResponse> CompleteAsync(
        ResolvedPrompt prompt, string apiKey, string model, AppSettings settings, CancellationToken ct = default)
    {
        var maxTokens = settings.MaxTokens;
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

        var messages = new List<Dictionary<string, object>>();
        if (!string.IsNullOrWhiteSpace(settings.SystemPrompt))
            messages.Add(new Dictionary<string, object> { { "role", "system" }, { "content", settings.SystemPrompt } });
        messages.Add(new Dictionary<string, object> { { "role", "user" }, { "content", content } });

        var requestBody = new Dictionary<string, object>
        {
            { "model", model },
            { "messages", messages.ToArray() }
        };

        requestBody[maxTokensKey] = maxTokens;

        // Reasoning models (o1, o3, o4) don't support temperature, top_p, or system messages
        if (!IsReasoningModel(model))
        {
            requestBody["temperature"] = settings.Temperature;
            requestBody["top_p"] = settings.TopP;
        }
        else
        {
            // Remove system message for reasoning models — they don't support it
            if (messages.Count > 1 && messages[0].ContainsKey("role") && (string)messages[0]["role"] == "system")
                messages.RemoveAt(0);
            requestBody["messages"] = messages.ToArray();
        }

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
        // Newer models (GPT-5.x, GPT-4.1, o-series) use max_completion_tokens to unlock
        // the full context window. Legacy max_tokens can artificially cap input size.
        if (model.StartsWith("gpt-5", StringComparison.OrdinalIgnoreCase)
            || model.StartsWith("gpt-4.1", StringComparison.OrdinalIgnoreCase)
            || IsReasoningModel(model))
            return "max_completion_tokens";
        return "max_tokens";
    }

    private static bool IsReasoningModel(string model)
    {
        return model.StartsWith("o1", StringComparison.OrdinalIgnoreCase)
            || model.StartsWith("o3", StringComparison.OrdinalIgnoreCase)
            || model.StartsWith("o4", StringComparison.OrdinalIgnoreCase);
    }

    private static object[] BuildMultimodalContent(ResolvedPrompt prompt)
    {
        var blocks = new List<object>();

        // Document/text attachments: file_id → inline text
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

        // Image attachments: file_id → base64 image_url
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
