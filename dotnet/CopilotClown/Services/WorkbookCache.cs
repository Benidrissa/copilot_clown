using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
using ExcelDna.Integration;

namespace CopilotClown.Services;

/// <summary>
/// Stores prompt→result mappings in the active workbook's Custom XML Parts.
/// Results survive save/close/reopen and travel with the .xlsx file.
/// </summary>
public static class WorkbookCache
{
    private const string Namespace = "urn:copilotclown:cache:v1";
    private static readonly SHA256 Sha = SHA256.Create();

    /// <summary>
    /// Try to retrieve a cached result for the given prompt key from the active workbook.
    /// Returns null if not found or workbook inaccessible.
    /// </summary>
    public static string Get(string promptKey)
    {
        try
        {
            var hash = HashKey(promptKey);
            var xml = GetCacheXml();
            if (xml == null) return null;

            // Parse entries from XML
            var entries = ParseEntries(xml);
            string value;
            return entries.TryGetValue(hash, out value) ? value : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Store a prompt→result pair in the active workbook's Custom XML Part.
    /// </summary>
    public static void Set(string promptKey, string result)
    {
        try
        {
            var hash = HashKey(promptKey);
            var xml = GetCacheXml();
            var entries = xml != null ? ParseEntries(xml) : new Dictionary<string, string>();

            entries[hash] = result;

            var newXml = BuildXml(entries);
            SetCacheXml(newXml);
        }
        catch
        {
            // Workbook not available or read-only — silently skip
        }
    }

    /// <summary>
    /// Remove a specific entry from the workbook cache.
    /// </summary>
    public static void Remove(string promptKey)
    {
        try
        {
            var hash = HashKey(promptKey);
            var xml = GetCacheXml();
            if (xml == null) return;

            var entries = ParseEntries(xml);
            if (entries.Remove(hash))
                SetCacheXml(BuildXml(entries));
        }
        catch { }
    }

    /// <summary>
    /// Clear all cached entries from the active workbook.
    /// </summary>
    public static void Clear()
    {
        try
        {
            dynamic app = ExcelDnaUtil.Application;
            dynamic wb = app.ActiveWorkbook;
            if (wb == null) return;

            dynamic parts = wb.CustomXMLParts;
            for (int i = parts.Count; i >= 1; i--)
            {
                dynamic part = parts[i];
                try
                {
                    if (part.NamespaceURI == Namespace)
                        part.Delete();
                }
                catch { }
            }
        }
        catch { }
    }

    /// <summary>
    /// Get count of entries in the workbook cache.
    /// </summary>
    public static int Count
    {
        get
        {
            try
            {
                var xml = GetCacheXml();
                if (xml == null) return 0;
                return ParseEntries(xml).Count;
            }
            catch { return 0; }
        }
    }

    // ── Internal helpers ──────────────────────────────────────────

    private static string HashKey(string key)
    {
        var data = Encoding.UTF8.GetBytes(key);
        byte[] hash;
        lock (Sha) { hash = Sha.ComputeHash(data); }
        return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant().Substring(0, 16);
    }

    private static string GetCacheXml()
    {
        dynamic app = ExcelDnaUtil.Application;
        dynamic wb = app.ActiveWorkbook;
        if (wb == null) return null;

        dynamic parts = wb.CustomXMLParts;
        for (int i = 1; i <= parts.Count; i++)
        {
            dynamic part = parts[i];
            try
            {
                if (part.NamespaceURI == Namespace)
                    return part.XML;
            }
            catch { }
        }
        return null;
    }

    private static void SetCacheXml(string xml)
    {
        dynamic app = ExcelDnaUtil.Application;
        dynamic wb = app.ActiveWorkbook;
        if (wb == null) return;

        // Remove existing cache part
        dynamic parts = wb.CustomXMLParts;
        for (int i = parts.Count; i >= 1; i--)
        {
            dynamic part = parts[i];
            try
            {
                if (part.NamespaceURI == Namespace)
                    part.Delete();
            }
            catch { }
        }

        // Add updated part
        parts.Add(xml);
    }

    /// <summary>
    /// Parse XML of form:
    /// &lt;cache xmlns="urn:copilotclown:cache:v1"&gt;
    ///   &lt;e k="hash"&gt;base64-result&lt;/e&gt;
    /// &lt;/cache&gt;
    /// </summary>
    private static Dictionary<string, string> ParseEntries(string xml)
    {
        var entries = new Dictionary<string, string>();
        // Simple parser — avoid LINQ/XDocument dependency
        int pos = 0;
        while (true)
        {
            int tagStart = xml.IndexOf("<e k=\"", pos, StringComparison.Ordinal);
            if (tagStart < 0) break;

            int keyStart = tagStart + 6;
            int keyEnd = xml.IndexOf("\"", keyStart, StringComparison.Ordinal);
            if (keyEnd < 0) break;

            var key = xml.Substring(keyStart, keyEnd - keyStart);

            int valStart = xml.IndexOf(">", keyEnd, StringComparison.Ordinal);
            if (valStart < 0) break;
            valStart++;

            int valEnd = xml.IndexOf("</e>", valStart, StringComparison.Ordinal);
            if (valEnd < 0) break;

            var b64 = xml.Substring(valStart, valEnd - valStart);
            try
            {
                entries[key] = Encoding.UTF8.GetString(Convert.FromBase64String(b64));
            }
            catch { }

            pos = valEnd + 4;
        }
        return entries;
    }

    private static string BuildXml(Dictionary<string, string> entries)
    {
        var sb = new StringBuilder();
        sb.Append("<cache xmlns=\"").Append(Namespace).Append("\">");
        foreach (var kv in entries)
        {
            sb.Append("<e k=\"").Append(kv.Key).Append("\">");
            sb.Append(Convert.ToBase64String(Encoding.UTF8.GetBytes(kv.Value)));
            sb.Append("</e>");
        }
        sb.Append("</cache>");
        return sb.ToString();
    }
}
