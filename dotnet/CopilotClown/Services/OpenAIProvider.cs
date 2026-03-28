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

        // Build input content
        var content = prompt.HasAttachments
            ? (object)BuildMultimodalContent(prompt)
            : new object[]
              {
                  new Dictionary<string, object> { { "type", "input_text" }, { "text", prompt.CleanText } }
              };

        // Build input array with optional system message
        var input = new List<object>();

        if (!string.IsNullOrWhiteSpace(settings.SystemPrompt) && !IsReasoningModel(model))
        {
            input.Add(new Dictionary<string, object>
            {
                { "role", "developer" },
                { "content", new object[]
                    {
                        new Dictionary<string, object> { { "type", "input_text" }, { "text", settings.SystemPrompt } }
                    }
                }
            });
        }

        input.Add(new Dictionary<string, object>
        {
            { "role", "user" },
            { "content", content }
        });

        var requestBody = new Dictionary<string, object>
        {
            { "model", model },
            { "input", input.ToArray() },
            { "max_output_tokens", maxTokens }
        };

        // Reasoning models (o1, o3, o4) don't support temperature or top_p
        if (!IsReasoningModel(model))
        {
            requestBody["temperature"] = settings.Temperature;
            requestBody["top_p"] = settings.TopP;
        }

        var request = new HttpRequestMessage(HttpMethod.Post, "https://api.openai.com/v1/responses")
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

        // Responses API: extract text from output[].content[].text
        var text = ExtractResponseText(root);
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

    private static bool IsReasoningModel(string model)
    {
        return model.StartsWith("o1", StringComparison.OrdinalIgnoreCase)
            || model.StartsWith("o3", StringComparison.OrdinalIgnoreCase)
            || model.StartsWith("o4", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Extract text from Responses API output.
    /// Response shape: { output: [ { type: "message", content: [ { type: "output_text", text: "..." } ] } ] }
    /// </summary>
    private static string ExtractResponseText(Dictionary<string, object> root)
    {
        // Try output[i].content[j].text for output_text blocks
        for (int i = 0; i < 10; i++)
        {
            var contentText = JsonHelper.GetString(root, $"output.{i}.content.0.text");
            if (!string.IsNullOrEmpty(contentText))
                return contentText;
        }
        // Fallback: try output_text at top level or single output text
        var fallback = JsonHelper.GetString(root, "output.0.text");
        return fallback ?? "";
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
                    { "type", "input_file" },
                    { "file_id", att.RemoteFileId }
                });
            }
            else
            {
                blocks.Add(new Dictionary<string, object>
                {
                    { "type", "input_text" },
                    { "text", $"--- Content from {att.FileName} ---\n{att.Content}\n--- End ---\n" }
                });
            }
        }

        // Image attachments: file_id → base64 input_image
        foreach (var att in prompt.Attachments.Where(a => a.Type == AttachmentType.Image))
        {
            if (!string.IsNullOrEmpty(att.RemoteFileId))
            {
                blocks.Add(new Dictionary<string, object>
                {
                    { "type", "input_file" },
                    { "file_id", att.RemoteFileId }
                });
            }
            else
            {
                blocks.Add(new Dictionary<string, object>
                {
                    { "type", "input_image" },
                    { "image_url", $"data:{att.MimeType};base64,{att.Content}" }
                });
            }
        }

        // User prompt text (always last)
        blocks.Add(new Dictionary<string, object>
        {
            { "type", "input_text" },
            { "text", prompt.CleanText }
        });

        return blocks.ToArray();
    }
}
