import { AppSettings, DEFAULT_SETTINGS, ProviderName } from "../types";

const KEYS = {
  provider: "aillm_provider",
  model: "aillm_model",
  keyPrefix: "aillm_key_",
  cacheEnabled: "aillm_cache_enabled",
  cacheTtl: "aillm_cache_ttl",
  cacheMax: "aillm_cache_max",
  rateMax: "aillm_rate_max",
  rateWindow: "aillm_rate_window",
  apiTimeout: "aillm_api_timeout",
};

export class SettingsService {
  getSettings(): AppSettings {
    return {
      activeProvider: (localStorage.getItem(KEYS.provider) as ProviderName) || DEFAULT_SETTINGS.activeProvider,
      activeModel: localStorage.getItem(KEYS.model) || DEFAULT_SETTINGS.activeModel,
      cacheEnabled: localStorage.getItem(KEYS.cacheEnabled) !== "false",
      cacheTtlMs: this.getInt(KEYS.cacheTtl, DEFAULT_SETTINGS.cacheTtlMs),
      cacheMaxEntries: this.getInt(KEYS.cacheMax, DEFAULT_SETTINGS.cacheMaxEntries),
      rateLimitMax: this.getInt(KEYS.rateMax, DEFAULT_SETTINGS.rateLimitMax),
      rateLimitWindowMs: this.getInt(KEYS.rateWindow, DEFAULT_SETTINGS.rateLimitWindowMs),
      apiTimeoutMs: this.getInt(KEYS.apiTimeout, DEFAULT_SETTINGS.apiTimeoutMs),
    };
  }

  saveSettings(settings: Partial<AppSettings>): void {
    if (settings.activeProvider !== undefined) localStorage.setItem(KEYS.provider, settings.activeProvider);
    if (settings.activeModel !== undefined) localStorage.setItem(KEYS.model, settings.activeModel);
    if (settings.cacheEnabled !== undefined) localStorage.setItem(KEYS.cacheEnabled, String(settings.cacheEnabled));
    if (settings.cacheTtlMs !== undefined) localStorage.setItem(KEYS.cacheTtl, String(settings.cacheTtlMs));
    if (settings.cacheMaxEntries !== undefined) localStorage.setItem(KEYS.cacheMax, String(settings.cacheMaxEntries));
    if (settings.rateLimitMax !== undefined) localStorage.setItem(KEYS.rateMax, String(settings.rateLimitMax));
    if (settings.rateLimitWindowMs !== undefined) localStorage.setItem(KEYS.rateWindow, String(settings.rateLimitWindowMs));
    if (settings.apiTimeoutMs !== undefined) localStorage.setItem(KEYS.apiTimeout, String(settings.apiTimeoutMs));
  }

  getApiKey(provider: ProviderName): string | null {
    const encoded = localStorage.getItem(KEYS.keyPrefix + provider);
    if (!encoded) return null;
    try {
      return atob(encoded);
    } catch {
      return encoded; // fallback: stored in plain text
    }
  }

  setApiKey(provider: ProviderName, key: string): void {
    localStorage.setItem(KEYS.keyPrefix + provider, btoa(key));
  }

  removeApiKey(provider: ProviderName): void {
    localStorage.removeItem(KEYS.keyPrefix + provider);
  }

  hasApiKey(provider: ProviderName): boolean {
    return !!localStorage.getItem(KEYS.keyPrefix + provider);
  }

  private getInt(key: string, defaultValue: number): number {
    const raw = localStorage.getItem(key);
    if (!raw) return defaultValue;
    const parsed = parseInt(raw, 10);
    return isNaN(parsed) ? defaultValue : parsed;
  }
}
