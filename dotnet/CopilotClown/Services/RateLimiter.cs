using System;
using System.Collections.Generic;

namespace CopilotClown.Services;

public class RateLimiter
{
    private readonly List<DateTime> _timestamps = new List<DateTime>();
    private int _maxCalls;
    private int _windowMinutes;

    public RateLimiter(int maxCalls = 500, int windowMinutes = 10)
    {
        _maxCalls = maxCalls;
        _windowMinutes = windowMinutes;
    }

    public bool TryAcquire()
    {
        var now = DateTime.UtcNow;
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
            PruneExpired(DateTime.UtcNow);
            return Math.Max(0, _maxCalls - _timestamps.Count);
        }
    }

    public void UpdateLimits(int maxCalls, int windowMinutes)
    {
        _maxCalls = maxCalls;
        _windowMinutes = windowMinutes;
    }

    public void Reset()
    {
        _timestamps.Clear();
    }

    private void PruneExpired(DateTime now)
    {
        var cutoff = now.AddMinutes(-_windowMinutes);
        _timestamps.RemoveAll(t => t <= cutoff);
    }
}
