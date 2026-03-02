using System;
using System.Runtime.Caching;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using CopilotClown.Models;

namespace CopilotClown.Services;

public class CacheService
{
    private static readonly MemoryCache Cache = MemoryCache.Default;
    private long _hits;
    private long _misses;

    public string Get(ProviderName provider, string model, string prompt)
    {
        var key = BuildKey(provider, model, prompt);
        var result = Cache.Get(key) as string;
        if (result != null)
            Interlocked.Increment(ref _hits);
        else
            Interlocked.Increment(ref _misses);
        return result;
    }

    public void Set(ProviderName provider, string model, string prompt, string response, int ttlMinutes)
    {
        var key = BuildKey(provider, model, prompt);
        var policy = new CacheItemPolicy
        {
            AbsoluteExpiration = DateTimeOffset.UtcNow.AddMinutes(ttlMinutes)
        };
        Cache.Set(key, response, policy);
    }

    public void Clear()
    {
        // MemoryCache doesn't have a Clear method — dispose and recreate is not possible
        // with the default instance, so we trim 100% of entries
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

    private static string BuildKey(ProviderName provider, string model, string prompt)
    {
        var data = $"{provider}|{model}|{prompt}";
        byte[] hash;
        using (var sha = SHA256.Create())
        {
            hash = sha.ComputeHash(Encoding.UTF8.GetBytes(data));
        }
        return "aillm_" + BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
    }
}
