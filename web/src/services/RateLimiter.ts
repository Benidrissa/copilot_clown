/**
 * Sliding-window rate limiter that runs entirely in memory.
 * Tracks timestamps of calls and rejects if the window is full.
 */
export class RateLimiter {
  private timestamps: number[] = [];
  private maxCalls: number;
  private windowMs: number;

  constructor(maxCalls: number = 100, windowMs: number = 10 * 60 * 1000) {
    this.maxCalls = maxCalls;
    this.windowMs = windowMs;
  }

  /**
   * Check if a call is allowed. If yes, records it and returns true.
   * If the rate limit is exceeded, returns false.
   */
  tryAcquire(): boolean {
    const now = Date.now();
    this.pruneExpired(now);

    if (this.timestamps.length >= this.maxCalls) {
      return false;
    }

    this.timestamps.push(now);
    return true;
  }

  getRemainingCalls(): number {
    this.pruneExpired(Date.now());
    return Math.max(0, this.maxCalls - this.timestamps.length);
  }

  getResetTimeMs(): number {
    if (this.timestamps.length === 0) return 0;
    const oldest = this.timestamps[0];
    return Math.max(0, oldest + this.windowMs - Date.now());
  }

  updateLimits(maxCalls: number, windowMs: number): void {
    this.maxCalls = maxCalls;
    this.windowMs = windowMs;
  }

  private pruneExpired(now: number): void {
    const cutoff = now - this.windowMs;
    while (this.timestamps.length > 0 && this.timestamps[0] <= cutoff) {
      this.timestamps.shift();
    }
  }
}
