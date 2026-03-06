using System;
using System.IO;
using System.Runtime.Caching;
using System.Threading;
using CopilotClown.Models;

namespace CopilotClown.Services;

public class ContentCache
{
    private static readonly MemoryCache Cache = new MemoryCache("FileContent");
    private long _hits;
    private long _misses;

    // Local files: no auto-expiration — persists until user clears cache.
    // Key includes LastWriteTimeUtc so modified files get re-extracted automatically.
    public Attachment GetLocal(string filePath)
    {
        var key = BuildLocalKey(filePath);
        var result = key != null ? Cache.Get(key) as Attachment : null;
        if (result != null)
            Interlocked.Increment(ref _hits);
        else
            Interlocked.Increment(ref _misses);
        return result;
    }

    public void SetLocal(string filePath, Attachment attachment)
    {
        var key = BuildLocalKey(filePath);
        if (key == null) return;
        // No expiration — stays until user clears or file changes (different key)
        Cache.Set(key, attachment, new CacheItemPolicy
        {
            AbsoluteExpiration = ObjectCache.InfiniteAbsoluteExpiration
        });
    }

    // URLs: cached with configurable TTL (default 24h to match API cache)
    public Attachment GetUrl(string url)
    {
        var result = Cache.Get("url_" + url) as Attachment;
        if (result != null)
            Interlocked.Increment(ref _hits);
        else
            Interlocked.Increment(ref _misses);
        return result;
    }

    public void SetUrl(string url, Attachment attachment, int ttlMinutes = 1440)
    {
        Cache.Set("url_" + url, attachment, new CacheItemPolicy
        {
            AbsoluteExpiration = DateTimeOffset.UtcNow.AddMinutes(ttlMinutes)
        });
    }

    public void Clear()
    {
        Cache.Trim(100);
        Interlocked.Exchange(ref _hits, 0);
        Interlocked.Exchange(ref _misses, 0);
    }

    public (long Entries, long Hits, long Misses) GetStats()
    {
        return (Cache.GetCount(), Interlocked.Read(ref _hits), Interlocked.Read(ref _misses));
    }

    private static string BuildLocalKey(string filePath)
    {
        try
        {
            var normalized = Path.GetFullPath(filePath).ToLowerInvariant();
            var lastWrite = File.GetLastWriteTimeUtc(filePath).Ticks;
            return $"file_{normalized}_{lastWrite}";
        }
        catch
        {
            return null;
        }
    }
}
