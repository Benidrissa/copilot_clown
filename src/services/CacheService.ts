import { CacheEntry, CacheStats, ProviderName } from "../types";

const CACHE_PREFIX = "aillm_cache_";
const STATS_KEY = "aillm_cache_stats";

interface InternalStats {
  hits: number;
  misses: number;
}

export class CacheService {
  private stats: InternalStats;

  constructor() {
    this.stats = this.loadStats();
  }

  async get(provider: ProviderName, model: string, prompt: string): Promise<string | null> {
    const key = await this.buildKey(provider, model, prompt);
    const raw = localStorage.getItem(key);

    if (!raw) {
      this.stats.misses++;
      this.saveStats();
      return null;
    }

    try {
      const entry: CacheEntry = JSON.parse(raw);
      const ttl = this.getTtl();
      const age = Date.now() - entry.timestamp;

      if (age > ttl) {
        localStorage.removeItem(key);
        this.stats.misses++;
        this.saveStats();
        return null;
      }

      this.stats.hits++;
      this.saveStats();
      return entry.response;
    } catch {
      localStorage.removeItem(key);
      this.stats.misses++;
      this.saveStats();
      return null;
    }
  }

  async set(provider: ProviderName, model: string, prompt: string, response: string): Promise<void> {
    if (!this.isEnabled()) return;

    this.enforceMaxEntries();

    const key = await this.buildKey(provider, model, prompt);
    const entry: CacheEntry = {
      response,
      timestamp: Date.now(),
      provider,
      model,
      promptPreview: prompt.slice(0, 100),
    };

    try {
      localStorage.setItem(key, JSON.stringify(entry));
    } catch {
      // localStorage full — evict oldest entries and retry
      this.evictOldest(10);
      try {
        localStorage.setItem(key, JSON.stringify(entry));
      } catch {
        // Still full — give up silently
      }
    }
  }

  clearAll(): void {
    const keys = this.getCacheKeys();
    for (const key of keys) {
      localStorage.removeItem(key);
    }
    this.stats = { hits: 0, misses: 0 };
    this.saveStats();
  }

  getStats(): CacheStats {
    const keys = this.getCacheKeys();
    let totalSizeBytes = 0;
    for (const key of keys) {
      const val = localStorage.getItem(key);
      if (val) totalSizeBytes += key.length + val.length;
    }

    const total = this.stats.hits + this.stats.misses;
    return {
      totalEntries: keys.length,
      totalSizeBytes: totalSizeBytes * 2, // UTF-16 chars = 2 bytes each
      hitRate: total > 0 ? this.stats.hits / total : 0,
      hits: this.stats.hits,
      misses: this.stats.misses,
    };
  }

  isEnabled(): boolean {
    return localStorage.getItem("aillm_cache_enabled") !== "false";
  }

  private getTtl(): number {
    const raw = localStorage.getItem("aillm_cache_ttl");
    return raw ? parseInt(raw, 10) : 24 * 60 * 60 * 1000;
  }

  private getMaxEntries(): number {
    const raw = localStorage.getItem("aillm_cache_max");
    return raw ? parseInt(raw, 10) : 1000;
  }

  private async buildKey(provider: string, model: string, prompt: string): Promise<string> {
    const data = JSON.stringify({ provider, model, prompt });
    const hash = await this.sha256(data);
    return CACHE_PREFIX + hash;
  }

  private async sha256(message: string): Promise<string> {
    const encoder = new TextEncoder();
    const data = encoder.encode(message);
    const hashBuffer = await crypto.subtle.digest("SHA-256", data);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    return hashArray.map((b) => b.toString(16).padStart(2, "0")).join("");
  }

  private getCacheKeys(): string[] {
    const keys: string[] = [];
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (key?.startsWith(CACHE_PREFIX)) {
        keys.push(key);
      }
    }
    return keys;
  }

  private enforceMaxEntries(): void {
    const keys = this.getCacheKeys();
    const max = this.getMaxEntries();
    if (keys.length >= max) {
      this.evictOldest(keys.length - max + 1);
    }
  }

  private evictOldest(count: number): void {
    const entries: { key: string; timestamp: number }[] = [];

    for (const key of this.getCacheKeys()) {
      const raw = localStorage.getItem(key);
      if (raw) {
        try {
          const entry: CacheEntry = JSON.parse(raw);
          entries.push({ key, timestamp: entry.timestamp });
        } catch {
          localStorage.removeItem(key);
        }
      }
    }

    entries.sort((a, b) => a.timestamp - b.timestamp);
    for (let i = 0; i < Math.min(count, entries.length); i++) {
      localStorage.removeItem(entries[i].key);
    }
  }

  private loadStats(): InternalStats {
    const raw = localStorage.getItem(STATS_KEY);
    if (raw) {
      try {
        return JSON.parse(raw);
      } catch {
        return { hits: 0, misses: 0 };
      }
    }
    return { hits: 0, misses: 0 };
  }

  private saveStats(): void {
    localStorage.setItem(STATS_KEY, JSON.stringify(this.stats));
  }
}
