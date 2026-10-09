import { CircuitBreakerState } from "../mcp/types.js";

export class CircuitBreaker {
  private state: CircuitBreakerState = {
    failures: 0,
    lastFailure: 0,
    state: "closed",
    nextAttempt: 0,
  };

  constructor(
    private readonly failureThreshold: number = 5,
    private readonly resetTimeoutMs: number = 30000,
    private readonly halfOpenAttempts: number = 3
  ) {}

  async execute<T>(operation: () => Promise<T>): Promise<T> {
    if (this.state.state === "open") {
      if (Date.now() >= this.state.nextAttempt) {
        this.state.state = "half-open";
        this.state.failures = 0;
      } else {
        throw new Error("Circuit breaker is open");
      }
    }

    try {
      const result = await operation();
      this.onSuccess();
      return result;
    } catch (error) {
      this.onFailure();
      throw error;
    }
  }

  private onSuccess(): void {
    if (this.state.state === "half-open") {
      this.state.failures++;
      if (this.state.failures >= this.halfOpenAttempts) {
        this.state.state = "closed";
        this.state.failures = 0;
      }
    } else {
      this.state.failures = 0;
    }
  }

  private onFailure(): void {
    this.state.failures++;
    this.state.lastFailure = Date.now();

    if (this.state.state === "half-open") {
      this.state.state = "open";
      this.state.nextAttempt = Date.now() + this.resetTimeoutMs;
    } else if (this.state.failures >= this.failureThreshold) {
      this.state.state = "open";
      this.state.nextAttempt = Date.now() + this.resetTimeoutMs;
    }
  }

  getState(): Readonly<CircuitBreakerState> {
    return { ...this.state };
  }

  reset(): void {
    this.state = {
      failures: 0,
      lastFailure: 0,
      state: "closed",
      nextAttempt: 0,
    };
  }

  isAvailable(): boolean {
    if (this.state.state === "closed") return true;
    if (this.state.state === "open" && Date.now() >= this.state.nextAttempt) return true;
    return this.state.state === "half-open";
  }
}