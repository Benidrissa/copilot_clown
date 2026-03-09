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

    public int MaxCalls => _maxCalls;

    public int WindowMinutes => _windowMinutes;

    public double UsagePercentage
    {
        get
        {
            PruneExpired(DateTime.UtcNow);
            return _maxCalls > 0 ? (_timestamps.Count / (double)_maxCalls) * 100.0 : 0;
        }
    }

    public bool IsNearLimit
    {
        get
        {
            PruneExpired(DateTime.UtcNow);
            return _maxCalls > 0 && RemainingCalls <= _maxCalls * 0.1 && RemainingCalls > 0;
        }
    }

    public bool IsLimited
    {
        get
        {
            PruneExpired(DateTime.UtcNow);
            return _timestamps.Count >= _maxCalls;
        }
    }

    public DateTime? OldestCallTime
    {
        get
        {
            PruneExpired(DateTime.UtcNow);
            return _timestamps.Count > 0 ? (DateTime?)_timestamps[0] : null;
        }
    }

    public TimeSpan? TimeUntilNextSlot
    {
        get
        {
            var now = DateTime.UtcNow;
            PruneExpired(now);
            if (_timestamps.Count < _maxCalls)
                return null; // slots available

            // Oldest call + window duration = when it expires
            var oldest = _timestamps[0];
            var expiresAt = oldest.AddMinutes(_windowMinutes);
            var remaining = expiresAt - now;
            return remaining > TimeSpan.Zero ? remaining : TimeSpan.Zero;
        }
    }

    public string FormatWaitTime()
    {
        var wait = TimeUntilNextSlot;
        if (wait == null) return null;
        var ts = wait.Value;
        if (ts.TotalSeconds < 1) return null;
        return ts.TotalMinutes >= 1
            ? $"~{(int)ts.TotalMinutes}m {ts.Seconds}s"
            : $"~{ts.Seconds}s";
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
