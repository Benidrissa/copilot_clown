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
            var fullMatch = match.Value;          // e.g. {C:\docs\file.pdf}
            var reference = match.Groups[1].Value; // e.g. C:\docs\file.pdf

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
                    var folderFiles = ResolveFolder(reference);
                    attachments.AddRange(folderFiles);
                    replacement = folderFiles.Count > 0
                        ? $"[attached: {folderFiles.Count} files from {Path.GetFileName(reference.TrimEnd('\\', '/'))}]"
                        : $"[folder: {reference} — no supported files found]";
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
                else
                {
                    // Not a valid file/folder/URL — leave as-is (might be intentional {braces})
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

    private List<Attachment> ResolveFolder(string folderPath)
    {
        var results = new List<Attachment>();
        var files = Directory.EnumerateFiles(folderPath, "*.*", SearchOption.AllDirectories)
            .Where(f => ContentExtractor.IsSupported(Path.GetExtension(f)))
            .Take(MaxFilesPerFolder)
            .ToList();

        foreach (var file in files)
        {
            try
            {
                var resolved = ResolveFile(file);
                if (resolved != null)
                    results.Add(resolved);
            }
            catch
            {
                // Skip files that can't be processed
            }
        }

        return results;
    }

    private static bool IsUrl(string reference)
    {
        return reference.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
            || reference.StartsWith("https://", StringComparison.OrdinalIgnoreCase);
    }
}
