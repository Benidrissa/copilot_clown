using System;
using System.Runtime.Caching;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace CopilotClown.Services;

public class CacheService
{
    private static readonly MemoryCache Cache = MemoryCache.Default;
    private static readonly SHA256 Sha = SHA256.Create();
    private long _hits;
    private long _misses;

    public string Get(string prompt)
    {
        var key = BuildKey(prompt);
        var result = Cache.Get(key) as string;
        if (result != null)
            Interlocked.Increment(ref _hits);
        else
            Interlocked.Increment(ref _misses);
        return result;
    }

    public void Set(string prompt, string response, int ttlMinutes)
    {
        var key = BuildKey(prompt);
        var policy = new CacheItemPolicy
        {
            AbsoluteExpiration = DateTimeOffset.UtcNow.AddMinutes(ttlMinutes)
        };
        Cache.Set(key, response, policy);
    }

    public void Clear()
    {
        Cache.Trim(100);
        Interlocked.Exchange(ref _hits, 0);
        Interlocked.Exchange(ref _misses, 0);
    }

    public (long Entries, long Hits, long Misses, double HitRate) GetStats()
    {
        var entries = Cache.GetCount();
        var hits = Interlocked.Read(ref _hits);
        var misses = Interlocked.Read(ref _misses);
        var total = hits + misses;
        var hitRate = total > 0 ? (double)hits / total : 0.0;
        return (entries, hits, misses, hitRate);
    }

    private static string BuildKey(string prompt)
    {
        // For long prompts, fingerprint instead of hashing the entire string
        var fingerprint = prompt.Length <= 2048
            ? prompt
            : string.Concat(prompt.Length.ToString(), prompt.Substring(0, 512), prompt.Substring(prompt.Length - 512));

        var data = Encoding.UTF8.GetBytes(fingerprint);
        byte[] hash;
        lock (Sha) { hash = Sha.ComputeHash(data); }
        return "aillm_" + BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
    }
}
