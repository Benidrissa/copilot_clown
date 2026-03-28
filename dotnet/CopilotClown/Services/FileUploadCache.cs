using System;
using System.IO;
using System.Runtime.Caching;
using System.Threading;
using CopilotClown.Models;

namespace CopilotClown.Services;

/// <summary>
/// Caches API-level file IDs (provider file uploads) to avoid re-uploading
/// the same file on every API call. Maps (provider, filePath) → remote file_id.
/// </summary>
public class FileUploadCache
{
    private static readonly MemoryCache Cache = new MemoryCache("FileUpload");
    private long _uploads;
    private long _reuses;

    /// <summary>
    /// Get a cached remote file_id for a local file + provider combination.
    /// Returns null on cache miss.
    /// </summary>
    public string Get(ProviderName provider, string sourcePath)
    {
        var key = BuildKey(provider, sourcePath);
        if (key == null) return null;

        var result = Cache.Get(key) as string;
        if (result != null)
            Interlocked.Increment(ref _reuses);
        return result;
    }

    /// <summary>
    /// Get a cached remote file_id for a URL + provider combination.
    /// </summary>
    public string GetUrl(ProviderName provider, string url)
    {
        var key = $"upload_{provider}_{url}";
        var result = Cache.Get(key) as string;
        if (result != null)
            Interlocked.Increment(ref _reuses);
        return result;
    }

    /// <summary>
    /// Cache a remote file_id for a local file. No expiration — persists until cleared
    /// or file changes (key includes LastWriteTimeUtc).
    /// </summary>
    public void Set(ProviderName provider, string sourcePath, string fileId)
    {
        var key = BuildKey(provider, sourcePath);
        if (key == null) return;

        Cache.Set(key, fileId, new CacheItemPolicy
        {
            AbsoluteExpiration = ObjectCache.InfiniteAbsoluteExpiration
        });
        Interlocked.Increment(ref _uploads);
    }

    /// <summary>
    /// Cache a remote file_id for a URL. 24h TTL.
    /// </summary>
    public void SetUrl(ProviderName provider, string url, string fileId)
    {
        Cache.Set($"upload_{provider}_{url}", fileId, new CacheItemPolicy
        {
            AbsoluteExpiration = DateTimeOffset.UtcNow.AddHours(24)
        });
        Interlocked.Increment(ref _uploads);
    }

    /// <summary>
    /// Remove a cached file_id for a local file (e.g., when the remote file has expired).
    /// </summary>
    public void Remove(ProviderName provider, string sourcePath)
    {
        var key = BuildKey(provider, sourcePath);
        if (key != null) Cache.Remove(key);
    }

    /// <summary>
    /// Remove a cached file_id for a URL.
    /// </summary>
    public void RemoveUrl(ProviderName provider, string url)
    {
        Cache.Remove($"upload_{provider}_{url}");
    }

    public void Clear()
    {
        Cache.Trim(100);
        Interlocked.Exchange(ref _uploads, 0);
        Interlocked.Exchange(ref _reuses, 0);
    }

    public (long Uploads, long Reuses) GetStats()
    {
        return (Interlocked.Read(ref _uploads), Interlocked.Read(ref _reuses));
    }

    private static string BuildKey(ProviderName provider, string sourcePath)
    {
        try
        {
            // For local files, include last write time so modified files get re-uploaded
            if (File.Exists(sourcePath))
            {
                var normalized = Path.GetFullPath(sourcePath).ToLowerInvariant();
                var lastWrite = File.GetLastWriteTimeUtc(sourcePath).Ticks;
                return $"upload_{provider}_{normalized}_{lastWrite}";
            }
            // For non-file sources (e.g., temp files from URL downloads), key by path only
            return $"upload_{provider}_{sourcePath}";
        }
        catch
        {
            return null;
        }
    }
}
