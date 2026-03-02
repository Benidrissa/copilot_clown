import React, { useState, useCallback, useEffect } from "react";
import { CacheService } from "../../services/CacheService";
import { SettingsService } from "../../services/SettingsService";
import { CacheStats } from "../../types";

interface Props {
  cacheService: CacheService;
  settingsService: SettingsService;
}

const TTL_PRESETS = [
  { label: "1 hour", ms: 60 * 60 * 1000 },
  { label: "6 hours", ms: 6 * 60 * 60 * 1000 },
  { label: "24 hours", ms: 24 * 60 * 60 * 1000 },
  { label: "7 days", ms: 7 * 24 * 60 * 60 * 1000 },
  { label: "30 days", ms: 30 * 24 * 60 * 60 * 1000 },
];

export function CacheManager({ cacheService, settingsService }: Props) {
  const settings = settingsService.getSettings();
  const [stats, setStats] = useState<CacheStats>(cacheService.getStats());
  const [cacheEnabled, setCacheEnabled] = useState(settings.cacheEnabled);
  const [ttlMs, setTtlMs] = useState(settings.cacheTtlMs);

  const refreshStats = useCallback(() => {
    setStats(cacheService.getStats());
  }, [cacheService]);

  useEffect(() => {
    const interval = setInterval(refreshStats, 5000);
    return () => clearInterval(interval);
  }, [refreshStats]);

  const handleClearCache = useCallback(() => {
    cacheService.clearAll();
    refreshStats();
  }, [cacheService, refreshStats]);

  const handleToggleCache = useCallback(
    (enabled: boolean) => {
      setCacheEnabled(enabled);
      settingsService.saveSettings({ cacheEnabled: enabled });
    },
    [settingsService],
  );

  const handleTtlChange = useCallback(
    (ms: number) => {
      setTtlMs(ms);
      settingsService.saveSettings({ cacheTtlMs: ms });
    },
    [settingsService],
  );

  const formatBytes = (bytes: number): string => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  };

  return (
    <div className="section">
      <div className="section-title">Cache</div>

      <div className="toggle-row">
        <span>Enable caching</span>
        <label className="toggle-switch">
          <input
            type="checkbox"
            checked={cacheEnabled}
            onChange={(e) => handleToggleCache(e.target.checked)}
          />
          <span className="toggle-slider" />
        </label>
      </div>

      <div className="stats-grid">
        <div className="stat-card">
          <div className="stat-value">{stats.totalEntries}</div>
          <div className="stat-label">Cached entries</div>
        </div>
        <div className="stat-card">
          <div className="stat-value">{formatBytes(stats.totalSizeBytes)}</div>
          <div className="stat-label">Cache size</div>
        </div>
        <div className="stat-card">
          <div className="stat-value">{(stats.hitRate * 100).toFixed(0)}%</div>
          <div className="stat-label">Hit rate</div>
        </div>
        <div className="stat-card">
          <div className="stat-value">
            {stats.hits}/{stats.hits + stats.misses}
          </div>
          <div className="stat-label">Hits / Total</div>
        </div>
      </div>

      <div className="field">
        <label>Cache TTL</label>
        <select value={ttlMs} onChange={(e) => handleTtlChange(Number(e.target.value))}>
          {TTL_PRESETS.map((preset) => (
            <option key={preset.ms} value={preset.ms}>
              {preset.label}
            </option>
          ))}
        </select>
      </div>

      <div className="btn-row">
        <button className="btn btn-danger btn-sm" onClick={handleClearCache}>
          Clear All Cache
        </button>
        <button className="btn btn-secondary btn-sm" onClick={refreshStats}>
          Refresh Stats
        </button>
      </div>
    </div>
  );
}
