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

        // Check cache (cached values are already markdown-stripped)
        if (settings.CacheEnabled)
        {
            var cached = Cache.Get(provider, model, prompt);
            if (cached != null)
                return FormatResponse(cached, singleCell);
        }

        // Use ExcelAsyncUtil for async API call
        var cacheKey = $"{provider}|{model}|{prompt}|{(singleCell ? "S" : "M")}";
        return ExcelAsyncUtil.Run(singleCell ? "USEAI.SINGLE" : "USEAI", cacheKey, () =>
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

                // Strip markdown ONCE, then cache the clean version
                var cleanText = StripMarkdown(result.Text.Trim());

                if (settings.CacheEnabled)
                    Cache.Set(provider, model, prompt, cleanText, settings.CacheTtlMinutes);

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
        });
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
}
