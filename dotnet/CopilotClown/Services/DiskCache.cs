using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace CopilotClown.Services;

/// <summary>
/// Disk-backed prompt→result cache stored at %APPDATA%\CopilotClown\cache.dat.
/// Survives Excel restarts. Entries expire based on TTL.
/// Thread-safe via lock.
/// </summary>
public class DiskCache
{
    private static readonly string CacheDir =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "CopilotClown");
    private static readonly string CachePath = Path.Combine(CacheDir, "cache.dat");
    private static readonly SHA256 Sha = SHA256.Create();
    private static readonly JavaScriptSerializer Json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

    private readonly object _lock = new object();
    private Dictionary<string, CacheEntry> _entries;
    private bool _loaded;

    public DiskCache()
    {
        _entries = new Dictionary<string, CacheEntry>();
    }

    /// <summary>
    /// Try to retrieve a cached result. Returns null if not found or expired.
    /// </summary>
    public string Get(string promptKey, int ttlMinutes)
    {
        lock (_lock)
        {
            EnsureLoaded();
            var hash = HashKey(promptKey);
            CacheEntry entry;
            if (!_entries.TryGetValue(hash, out entry)) return null;

            if (ttlMinutes > 0 && entry.Created.AddMinutes(ttlMinutes) < DateTime.UtcNow)
            {
                _entries.Remove(hash);
                SaveToDisk();
                return null;
            }
            return entry.Value;
        }
    }

    /// <summary>
    /// Store a prompt→result pair on disk.
    /// </summary>
    public void Set(string promptKey, string result)
    {
        lock (_lock)
        {
            EnsureLoaded();
            var hash = HashKey(promptKey);
            _entries[hash] = new CacheEntry { Value = result, Created = DateTime.UtcNow };
            SaveToDisk();
        }
    }

    /// <summary>
    /// Remove a specific entry.
    /// </summary>
    public void Remove(string promptKey)
    {
        lock (_lock)
        {
            EnsureLoaded();
            var hash = HashKey(promptKey);
            if (_entries.Remove(hash))
                SaveToDisk();
        }
    }

    /// <summary>
    /// Clear all disk cache entries.
    /// </summary>
    public void Clear()
    {
        lock (_lock)
        {
            _entries.Clear();
            _loaded = true;
            try { if (File.Exists(CachePath)) File.Delete(CachePath); } catch { }
        }
    }

    /// <summary>
    /// Number of entries currently in the disk cache.
    /// </summary>
    public int Count
    {
        get
        {
            lock (_lock)
            {
                EnsureLoaded();
                return _entries.Count;
            }
        }
    }

    /// <summary>
    /// Prune expired entries given a TTL. Returns number removed.
    /// </summary>
    public int Prune(int ttlMinutes)
    {
        if (ttlMinutes <= 0) return 0;
        lock (_lock)
        {
            EnsureLoaded();
            var cutoff = DateTime.UtcNow.AddMinutes(-ttlMinutes);
            var toRemove = new List<string>();
            foreach (var kv in _entries)
                if (kv.Value.Created < cutoff)
                    toRemove.Add(kv.Key);

            foreach (var key in toRemove)
                _entries.Remove(key);

            if (toRemove.Count > 0)
                SaveToDisk();

            return toRemove.Count;
        }
    }

    // ── Internal ──────────────────────────────────────────────────

    private void EnsureLoaded()
    {
        if (_loaded) return;
        _loaded = true;
        LoadFromDisk();
    }

    private void LoadFromDisk()
    {
        try
        {
            if (!File.Exists(CachePath)) return;
            var text = File.ReadAllText(CachePath, Encoding.UTF8);
            var raw = Json.Deserialize<Dictionary<string, object>>(text);
            if (raw == null) return;

            foreach (var kv in raw)
            {
                var dict = kv.Value as Dictionary<string, object>;
                if (dict == null) continue;
                var entry = new CacheEntry();
                object val;
                if (dict.TryGetValue("v", out val)) entry.Value = val as string;
                if (dict.TryGetValue("c", out val) && val is string s)
                {
                    DateTime dt;
                    if (DateTime.TryParse(s, null, System.Globalization.DateTimeStyles.RoundtripKind, out dt))
                        entry.Created = dt;
                }
                if (entry.Value != null)
                    _entries[kv.Key] = entry;
            }
        }
        catch { }
    }

    private void SaveToDisk()
    {
        try
        {
            if (!Directory.Exists(CacheDir))
                Directory.CreateDirectory(CacheDir);

            var raw = new Dictionary<string, object>();
            foreach (var kv in _entries)
            {
                raw[kv.Key] = new Dictionary<string, object>
                {
                    { "v", kv.Value.Value },
                    { "c", kv.Value.Created.ToString("o") }
                };
            }
            File.WriteAllText(CachePath, Json.Serialize(raw), Encoding.UTF8);
        }
        catch { }
    }

    private static string HashKey(string key)
    {
        var data = Encoding.UTF8.GetBytes(key);
        byte[] hash;
        lock (Sha) { hash = Sha.ComputeHash(data); }
        return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant().Substring(0, 16);
    }

    private class CacheEntry
    {
        public string Value;
        public DateTime Created;
    }
}
