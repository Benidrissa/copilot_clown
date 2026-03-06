using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using CopilotClown.Models;

namespace CopilotClown.Services;

public class FileResolver
{
    private readonly ContentCache _cache;
    private readonly ContentExtractor _extractor;

    private const int MaxFilesPerFolder = 50;

    // Match anything between { } — file paths, folder paths, or URLs
    private static readonly Regex BraceRefRegex = new Regex(
        @"\{([^}]+)\}",
        RegexOptions.Compiled);

    public FileResolver(ContentCache cache)
    {
        _cache = cache;
        _extractor = new ContentExtractor();
    }

    public ResolvedPrompt Resolve(string prompt)
    {
        var matches = BraceRefRegex.Matches(prompt);
        if (matches.Count == 0)
            return new ResolvedPrompt(prompt);

        var attachments = new List<Attachment>();
        var cleanPrompt = new StringBuilder(prompt);
        int offset = 0; // Track offset as we replace strings of different lengths

        foreach (Match match in matches)
        {
            var fullMatch = match.Value;                    // e.g. {C:\docs\file.pdf}
            var reference = match.Groups[1].Value.Trim();  // e.g. C:\docs\file.pdf

            // Strip surrounding quotes: {"C:\path with spaces"} → C:\path with spaces
            if (reference.Length >= 2 && reference[0] == '"' && reference[reference.Length - 1] == '"')
                reference = reference.Substring(1, reference.Length - 2).Trim();

            string replacement;
            try
            {
                if (IsUrl(reference))
                {
                    var resolved = ResolveUrl(reference);
                    if (resolved != null)
                    {
                        attachments.Add(resolved);
                        replacement = $"[attached: {resolved.FileName}]";
                    }
                    else
                    {
                        replacement = $"[Error: could not download {reference}]";
                    }
                }
                else if (Directory.Exists(reference))
                {
                    string diagnostic;
                    var folderFiles = ResolveFolder(reference, out diagnostic);
                    attachments.AddRange(folderFiles);
                    var folderName = Path.GetFileName(reference.TrimEnd('\\', '/'));
                    replacement = folderFiles.Count > 0
                        ? $"[attached: {folderFiles.Count} files from {folderName}]"
                        : $"[folder: {folderName} — no supported files ({diagnostic})]";
                }
                else if (File.Exists(reference))
                {
                    var resolved = ResolveFile(reference);
                    if (resolved != null)
                    {
                        attachments.Add(resolved);
                        replacement = $"[attached: {resolved.FileName}]";
                    }
                    else
                    {
                        replacement = $"[Error: unsupported file type {Path.GetExtension(reference)}]";
                    }
                }
                else if (LooksLikePath(reference))
                {
                    replacement = $"[Warning: path not found: {reference}]";
                }
                else
                {
                    // Not a path or URL — leave as-is (might be intentional {braces})
                    continue;
                }
            }
            catch (Exception ex)
            {
                replacement = $"[Error: {ex.Message}]";
            }

            // Replace in the StringBuilder, accounting for offset from previous replacements
            var adjustedIndex = match.Index + offset;
            cleanPrompt.Remove(adjustedIndex, fullMatch.Length);
            cleanPrompt.Insert(adjustedIndex, replacement);
            offset += replacement.Length - fullMatch.Length;
        }

        return new ResolvedPrompt(cleanPrompt.ToString(), attachments);
    }

    private Attachment ResolveFile(string filePath)
    {
        var ext = Path.GetExtension(filePath);
        if (!ContentExtractor.IsSupported(ext))
            return null;

        // Check cache first
        var cached = _cache.GetLocal(filePath);
        if (cached != null) return cached;

        var result = _extractor.Extract(filePath);
        if (result != null)
            _cache.SetLocal(filePath, result);

        return result;
    }

    private Attachment ResolveUrl(string url)
    {
        // Check cache first
        var cached = _cache.GetUrl(url);
        if (cached != null) return cached;

        var (data, contentType, fileName) = _extractor.DownloadUrl(url);
        var result = _extractor.ExtractFromBytes(data, fileName, contentType);
        if (result != null)
            _cache.SetUrl(url, result);

        return result;
    }

    private List<Attachment> ResolveFolder(string folderPath, out string diagnostic)
    {
        var results = new List<Attachment>();
        var allFiles = new List<string>();
        var errors = new List<string>();
        int totalScanned = 0;
        int dirsScanned = 0;

        SafeCollectFiles(folderPath, allFiles, MaxFilesPerFolder, ref totalScanned, ref dirsScanned, errors);

        foreach (var file in allFiles)
        {
            try
            {
                var resolved = ResolveFile(file);
                if (resolved != null)
                    results.Add(resolved);
            }
            catch (Exception ex)
            {
                errors.Add($"extract({Path.GetFileName(file)}): {ex.Message}");
            }
        }

        var diag = $"scanned {totalScanned} files in {dirsScanned} dirs, {allFiles.Count} supported";
        if (errors.Count > 0)
            diag += $"; errors: {string.Join("; ", errors)}";
        diagnostic = diag;

        return results;
    }

    /// <summary>
    /// Recursively collect supported files into a list, skipping directories
    /// that throw (access denied, OneDrive cloud-only, long paths, etc.).
    /// Uses Directory.GetFiles/GetDirectories (eager, not lazy) so exceptions
    /// are caught at call time. Errors are recorded for diagnostics.
    /// </summary>
    private static void SafeCollectFiles(string dir, List<string> results, int maxFiles,
        ref int totalScanned, ref int dirsScanned, List<string> errors)
    {
        if (results.Count >= maxFiles)
            return;

        dirsScanned++;

        // Collect files in current directory
        string[] files = null;
        try
        {
            files = Directory.GetFiles(dir);
        }
        catch (Exception ex)
        {
            errors.Add($"GetFiles({Path.GetFileName(dir)}): {ex.GetType().Name}: {ex.Message}");
        }

        if (files != null)
        {
            foreach (var file in files)
            {
                totalScanned++;
                if (results.Count >= maxFiles) return;
                if (ContentExtractor.IsSupported(Path.GetExtension(file)))
                    results.Add(file);
            }
        }

        // Recurse into subdirectories
        string[] subdirs = null;
        try
        {
            subdirs = Directory.GetDirectories(dir);
        }
        catch (Exception ex)
        {
            errors.Add($"GetDirs({Path.GetFileName(dir)}): {ex.GetType().Name}: {ex.Message}");
        }

        if (subdirs != null)
        {
            foreach (var subdir in subdirs)
            {
                if (results.Count >= maxFiles) return;
                SafeCollectFiles(subdir, results, maxFiles, ref totalScanned, ref dirsScanned, errors);
            }
        }
    }

    private static bool IsUrl(string reference)
    {
        return reference.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
            || reference.StartsWith("https://", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Heuristic: does this look like a file/folder path (drive letter, UNC, or rooted)?
    /// Used to give a "not found" warning instead of silently ignoring.
    /// </summary>
    private static bool LooksLikePath(string reference)
    {
        // Drive letter path: C:\... or C:/...
        if (reference.Length >= 3 && char.IsLetter(reference[0]) && reference[1] == ':'
            && (reference[2] == '\\' || reference[2] == '/'))
            return true;

        // UNC path: \\server\...
        if (reference.StartsWith("\\\\"))
            return true;

        // Rooted path: \something or /something
        if (reference.Length > 1 && (reference[0] == '\\' || reference[0] == '/'))
            return true;

        return false;
    }
}
