using System;
using System.Net;
using System.Net.Http;
using System.Text;
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
    private static readonly ContentCache FileCache = new ContentCache();
    private static readonly FileUploadCache UploadCache = new FileUploadCache();
    private static readonly FileResolver Resolver = new FileResolver(FileCache);

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
    internal static ContentCache FileCacheInstance => FileCache;
    internal static FileUploadCache UploadCacheInstance => UploadCache;

    // ───────────────────────────────────────────────────────────────
    //  USEAI  —  spills multi-line responses into separate rows
    // ───────────────────────────────────────────────────────────────
    [ExcelFunction(
        Name = "USEAI",
        Description = "Calls an AI model and spills the response into separate rows. Use USEAISINGLE to keep everything in one cell.",
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
        return CallLlm(new[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8 }, singleCell: false);
    }

    // ───────────────────────────────────────────────────────────────
    //  USEAISINGLE  —  returns full response in one cell (with line breaks)
    // ───────────────────────────────────────────────────────────────
    [ExcelFunction(
        Name = "USEAI.SINGLE",
        Description = "Calls an AI model and returns the full response in a single cell (enable Wrap Text to see all lines).",
        HelpTopic = "https://github.com/Benidrissa/copilot_clown")]
    public static object UseAiSingle(
        [ExcelArgument(Name = "prompt_part1", Description = "Text describing the task or question")] object arg1,
        [ExcelArgument(Name = "context1", Description = "[Optional] Cell reference or range providing context")] object arg2,
        [ExcelArgument(Name = "prompt_part2", Description = "[Optional] Additional prompt text")] object arg3,
        [ExcelArgument(Name = "context2", Description = "[Optional] Additional context")] object arg4,
        [ExcelArgument(Name = "prompt_part3", Description = "[Optional] Additional prompt text")] object arg5,
        [ExcelArgument(Name = "context3", Description = "[Optional] Additional context")] object arg6,
        [ExcelArgument(Name = "prompt_part4", Description = "[Optional] Additional prompt text")] object arg7,
        [ExcelArgument(Name = "context4", Description = "[Optional] Additional context")] object arg8)
    {
        return CallLlm(new[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8 }, singleCell: true);
    }

    // ───────────────────────────────────────────────────────────────
    //  Shared implementation
    // ───────────────────────────────────────────────────────────────
    private static object CallLlm(object[] args, bool singleCell)
    {
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

        // Resolve file/folder/URL references enclosed in { }
        ResolvedPrompt resolved;
        try
        {
            resolved = Resolver.Resolve(prompt);
        }
        catch (Exception ex)
        {
            return $"Error resolving file references: {ex.Message}";
        }

        // Build cache key that includes attachment identity
        var cachePromptKey = BuildCachePromptKey(resolved);

        // Check cache (cached values are already markdown-stripped)
        if (settings.CacheEnabled)
        {
            var cached = Cache.Get(provider, model, cachePromptKey);
            if (cached != null)
                return FormatResponse(cached, singleCell);
        }

        // Use ExcelAsyncUtil.Observe for async API call with "Loading..." indicator
        var asyncKey = $"{provider}|{model}|{cachePromptKey}|{(singleCell ? "S" : "M")}";
        return ExcelAsyncUtil.Observe(
            singleCell ? "USEAI.SINGLE" : "USEAI",
            asyncKey,
            () => new LoadingObservable(() =>
            {
                // Rate limit
                RateLimiter.UpdateLimits(settings.RateLimitMax, settings.RateLimitWindowMinutes);
                if (!RateLimiter.TryAcquire())
                    return (object)"Error: Rate limit exceeded. Wait and try again.";

                try
                {
                    var llm = ProviderFactory.Get(provider);

                    // Upload attachments to API if not already uploaded (Layer 2 cache)
                    if (resolved.HasAttachments)
                        UploadAttachments(llm, resolved, provider, apiKey);

                    var result = llm.CompleteAsync(resolved, apiKey, model, ct: CancellationToken.None)
                        .GetAwaiter().GetResult();

                    // Strip markdown ONCE, then cache the clean version
                    var cleanText = StripMarkdown(result.Text.Trim());

                    if (settings.CacheEnabled)
                        Cache.Set(provider, model, cachePromptKey, cleanText, settings.CacheTtlMinutes);

                    return FormatResponse(cleanText, singleCell);
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
            })
        );
    }

    private static string BuildCachePromptKey(ResolvedPrompt resolved)
    {
        if (!resolved.HasAttachments)
            return resolved.CleanText;

        var sb = new StringBuilder(resolved.CleanText);
        foreach (var att in resolved.Attachments)
        {
            sb.Append('|');
            sb.Append(att.SourcePath);
            sb.Append(':');
            sb.Append(att.Content.Length);
        }
        return sb.ToString();
    }

    /// <summary>
    /// Upload attachments to the provider's Files API if not already cached.
    /// Sets RemoteFileId on each attachment so providers use file_id references.
    /// Falls back silently to inline content on upload failure.
    /// </summary>
    private static void UploadAttachments(ILlmProvider llm, ResolvedPrompt resolved, ProviderName provider, string apiKey)
    {
        foreach (var att in resolved.Attachments)
        {
            try
            {
                // Check if already uploaded
                var cachedFileId = IsUrl(att.SourcePath)
                    ? UploadCache.GetUrl(provider, att.SourcePath)
                    : UploadCache.Get(provider, att.SourcePath);

                if (cachedFileId != null)
                {
                    att.RemoteFileId = cachedFileId;
                    continue;
                }

                // Get raw bytes for upload
                byte[] fileBytes;
                if (att.RawBytes != null)
                {
                    fileBytes = att.RawBytes;
                }
                else if (att.Type == AttachmentType.Image)
                {
                    fileBytes = Convert.FromBase64String(att.Content);
                }
                else if (System.IO.File.Exists(att.SourcePath))
                {
                    fileBytes = System.IO.File.ReadAllBytes(att.SourcePath);
                }
                else
                {
                    // Can't upload without raw bytes — use inline fallback
                    continue;
                }

                var fileId = llm.UploadFileAsync(fileBytes, att.FileName, att.MimeType, apiKey)
                    .GetAwaiter().GetResult();

                if (!string.IsNullOrEmpty(fileId))
                {
                    att.RemoteFileId = fileId;

                    if (IsUrl(att.SourcePath))
                        UploadCache.SetUrl(provider, att.SourcePath, fileId);
                    else
                        UploadCache.Set(provider, att.SourcePath, fileId);
                }
            }
            catch
            {
                // Upload failed — provider will use inline content as fallback
            }
        }
    }

    private static bool IsUrl(string path)
    {
        return path != null && (path.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
            || path.StartsWith("https://", StringComparison.OrdinalIgnoreCase));
    }

    // Excel cell character limit
    private const int CellMaxChars = 32767;

    /// <summary>
    /// Truncate to fit Excel's 32,767 char cell limit. No markdown stripping here
    /// (already stripped before caching).
    /// </summary>
    private static string FitCell(string s)
    {
        if (s == null) return "";
        return s.Length <= CellMaxChars ? s : s.Substring(0, CellMaxChars - 3) + "...";
    }

    /// <summary>
    /// Remove markdown formatting: bold, italic, headers, bullets, code blocks, tables.
    /// </summary>
    private static string StripMarkdown(string s)
    {
        if (string.IsNullOrEmpty(s)) return s;

        // Remove code blocks (```...```)
        while (s.Contains("```"))
        {
            int start = s.IndexOf("```");
            int end = s.IndexOf("```", start + 3);
            if (end > start)
            {
                var inner = s.Substring(start + 3, end - start - 3);
                int nl = inner.IndexOf('\n');
                if (nl >= 0 && nl < 20)
                    inner = inner.Substring(nl + 1);
                s = s.Substring(0, start) + inner + s.Substring(end + 3);
            }
            else
            {
                s = s.Substring(0, start) + s.Substring(start + 3);
            }
        }

        var lines = s.Split('\n');
        var result = new System.Collections.Generic.List<string>();

        for (int i = 0; i < lines.Length; i++)
        {
            var line = lines[i];

            // Skip table separator lines (|---|---|--- or | :--- | :---: |)
            var stripped = line.Trim();
            if (stripped.StartsWith("|") && stripped.Contains("---"))
            {
                bool isSeparator = true;
                foreach (var cell in stripped.Split('|'))
                {
                    var c = cell.Trim().Replace("-", "").Replace(":", "").Replace(" ", "");
                    if (c.Length > 0) { isSeparator = false; break; }
                }
                if (isSeparator) continue; // skip separator rows
            }

            // Convert table rows: | col1 | col2 | col3 | → col1 | col2 | col3
            if (stripped.StartsWith("|") && stripped.EndsWith("|") && stripped.Length > 2)
            {
                // Remove leading and trailing pipes, keep inner pipes as delimiters
                line = stripped.Substring(1, stripped.Length - 2).Trim();
                // Clean each cell
                var cells = line.Split('|');
                for (int c = 0; c < cells.Length; c++)
                    cells[c] = CleanInline(cells[c].Trim());
                line = string.Join(" | ", cells);
                result.Add(line);
                continue;
            }

            // Remove headers (# ## ### etc.)
            while (line.StartsWith("#"))
                line = line.Substring(1);
            line = line.TrimStart(' ');

            // Remove bullet markers (- item, * item, + item)
            if (line.StartsWith("- ") || line.StartsWith("* ") || line.StartsWith("+ "))
                line = line.Substring(2);

            // Remove numbered list markers (1. item, 2. item)
            if (line.Length > 2 && char.IsDigit(line[0]))
            {
                int dot = line.IndexOf(". ");
                if (dot > 0 && dot <= 3)
                    line = line.Substring(dot + 2);
            }

            line = CleanInline(line);
            result.Add(line);
        }

        return string.Join("\n", result);
    }

    /// <summary>
    /// Remove inline markdown: bold, italic, code, strikethrough.
    /// </summary>
    private static string CleanInline(string line)
    {
        line = RemoveWrapping(line, "***");
        line = RemoveWrapping(line, "___");
        line = RemoveWrapping(line, "**");
        line = RemoveWrapping(line, "__");
        line = RemoveWrapping(line, "~~");
        line = RemoveWrapping(line, "`");
        line = RemoveWrapping(line, "*");
        line = RemoveWrapping(line, "_");
        return line;
    }

    private static string RemoveWrapping(string s, string marker)
    {
        while (true)
        {
            int start = s.IndexOf(marker);
            if (start < 0) break;
            int end = s.IndexOf(marker, start + marker.Length);
            if (end < 0) break;
            s = s.Substring(0, start) + s.Substring(start + marker.Length, end - start - marker.Length) + s.Substring(end + marker.Length);
        }
        return s;
    }

    /// <summary>
    /// Format the AI response for Excel.
    /// singleCell=true  → full text in one cell with line breaks (USEAISINGLE).
    /// singleCell=false → spill as dynamic array (USEAI).
    /// </summary>
    private static object FormatResponse(string text, bool singleCell)
    {
        var trimmed = text.Trim();

        // ── Single-cell mode: return full text with embedded LF (CHAR(10)) ──
        if (singleCell)
        {
            var normalized = trimmed.Replace("\r\n", "\n").Replace("\r", "\n");
            return FitCell(normalized);
        }

        // ── Spill mode: split into rows ──
        var splitLines = trimmed.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
        if (splitLines.Length <= 1)
            return FitCell(trimmed);

        // Detect tabular data: if most lines contain " | ", split into 2D array
        int pipeLineCount = 0;
        int maxCols = 1;
        for (int i = 0; i < splitLines.Length; i++)
        {
            if (splitLines[i].Contains(" | "))
            {
                pipeLineCount++;
                var colCount = splitLines[i].Split(new[] { " | " }, StringSplitOptions.None).Length;
                if (colCount > maxCols) maxCols = colCount;
            }
        }

        // If majority of lines are pipe-delimited, return as 2D table
        if (pipeLineCount > splitLines.Length / 2 && maxCols > 1)
        {
            var table = new object[splitLines.Length, maxCols];
            for (int r = 0; r < splitLines.Length; r++)
            {
                var cells = splitLines[r].Trim().Split(new[] { " | " }, StringSplitOptions.None);
                for (int c = 0; c < maxCols; c++)
                {
                    if (c < cells.Length)
                    {
                        var val = cells[c].Trim().Trim('"');
                        // Try to parse as number for Excel
                        if (double.TryParse(val, System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out double num))
                            table[r, c] = num;
                        else
                            table[r, c] = FitCell(val);
                    }
                    else
                    {
                        table[r, c] = ExcelEmpty.Value;
                    }
                }
            }
            return table;
        }

        // Non-tabular: spill as single column
        var spill = new object[splitLines.Length, 1];
        for (int i = 0; i < splitLines.Length; i++)
            spill[i, 0] = FitCell(splitLines[i].Trim());
        return spill;
    }

    // ───────────────────────────────────────────────────────────────
    //  IExcelObservable that shows "Loading..." then the real result
    // ───────────────────────────────────────────────────────────────

    private class LoadingObservable : IExcelObservable
    {
        private readonly Func<object> _compute;

        public LoadingObservable(Func<object> compute)
        {
            _compute = compute;
        }

        public IDisposable Subscribe(IExcelObserver observer)
        {
            observer.OnNext("Loading...");

            Task.Run(() =>
            {
                try
                {
                    var result = _compute();
                    observer.OnNext(result);
                    observer.OnCompleted();
                }
                catch (Exception ex)
                {
                    observer.OnError(ex);
                }
            });

            return new ActionDisposable(() => { });
        }
    }

    private class ActionDisposable : IDisposable
    {
        private Action _action;

        public ActionDisposable(Action action) { _action = action; }

        public void Dispose()
        {
            var action = Interlocked.Exchange(ref _action, null);
            action?.Invoke();
        }
    }
}
