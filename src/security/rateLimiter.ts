import { RateLimitState } from "../mcp/types.js";

export class RateLimiter {
  private readonly limits: Map<string, RateLimitState> = new Map();
  private readonly maxTokens: number;
  private readonly refillRate: number;
  private readonly windowMs: number;

  constructor(maxTokens: number, windowMs: number = 60000) {
    this.maxTokens = maxTokens;
    this.windowMs = windowMs;
    this.refillRate = maxTokens / (windowMs / 1000);
  }

  tryConsume(key: string, tokens: number = 1): boolean {
    const now = Date.now();
    let state = this.limits.get(key);

    if (!state) {
      state = { tokens: this.maxTokens, lastRefill: now };
      this.limits.set(key, state);
    }

    this.refill(state, now);

    if (state.tokens >= tokens) {
      state.tokens -= tokens;
      return true;
    }

    return false;
  }

  getRemaining(key: string): number {
    const now = Date.now();
    const state = this.limits.get(key);
    if (!state) return this.maxTokens;
    this.refill(state, now);
    return Math.floor(state.tokens);
  }

  getResetTime(key: string): number {
    const state = this.limits.get(key);
    if (!state) return Date.now() + this.windowMs;
    const tokensNeeded = this.maxTokens - state.tokens;
    return state.lastRefill + Math.ceil(tokensNeeded / this.refillRate) * 1000;
  }

  private refill(state: RateLimitState, now: number): void {
    const elapsed = now - state.lastRefill;
    if (elapsed >= this.windowMs) {
      state.tokens = this.maxTokens;
      state.lastRefill = now;
    } else {
      const tokensToAdd = (elapsed / 1000) * this.refillRate;
      state.tokens = Math.min(this.maxTokens, state.tokens + tokensToAdd);
      state.lastRefill = now;
    }
  }

  reset(key: string): void {
    this.limits.delete(key);
  }

  cleanup(): void {
    const now = Date.now();
    for (const [key, state] of this.limits.entries()) {
      if (now - state.lastRefill > this.windowMs * 2) {
        this.limits.delete(key);
      }
    }
  }
}

export class TokenBucketRateLimiter {
  private readonly buckets: Map<string, { tokens: number; lastRefill: number }> = new Map();

  constructor(
    private readonly capacity: number,
    private readonly refillPerSecond: number
  ) {}

  consume(key: string, tokens: number = 1): boolean {
    const now = Date.now();
    let bucket = this.buckets.get(key);

    if (!bucket) {
      bucket = { tokens: this.capacity, lastRefill: now };
      this.buckets.set(key, bucket);
    }

    const elapsed = (now - bucket.lastRefill) / 1000;
    bucket.tokens = Math.min(this.capacity, bucket.tokens + elapsed * this.refillPerSecond);
    bucket.lastRefill = now;

    if (bucket.tokens >= tokens) {
      bucket.tokens -= tokens;
      return true;
    }

    return false;
  }

  getAvailable(key: string): number {
    const now = Date.now();
    const bucket = this.buckets.get(key);
    if (!bucket) return this.capacity;
    const elapsed = (now - bucket.lastRefill) / 1000;
    return Math.min(this.capacity, bucket.tokens + elapsed * this.refillPerSecond);
  }
}