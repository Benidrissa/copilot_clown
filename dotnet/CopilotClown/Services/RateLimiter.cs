using System;
using System.Collections.Generic;

namespace CopilotClown.Services;

public class RateLimiter
{
    private readonly List<long> _timestamps = new List<long>();
    private int _maxCalls;
    private int _windowMs;

    public RateLimiter(int maxCalls = 100, int windowMinutes = 10)
    {
        _maxCalls = maxCalls;
        _windowMs = windowMinutes * 60 * 1000;
    }

    public bool TryAcquire()
    {
        var now = (long)Environment.TickCount;
        PruneExpired(now);

        if (_timestamps.Count >= _maxCalls)
            return false;

        _timestamps.Add(now);
        return true;
    }

    public int RemainingCalls
    {
        get
        {
            PruneExpired((long)Environment.TickCount);
            return Math.Max(0, _maxCalls - _timestamps.Count);
        }
    }

    public void UpdateLimits(int maxCalls, int windowMinutes)
    {
        _maxCalls = maxCalls;
        _windowMs = windowMinutes * 60 * 1000;
    }

    private void PruneExpired(long now)
    {
        var cutoff = now - _windowMs;
        _timestamps.RemoveAll(t => t <= cutoff);
    }
}
