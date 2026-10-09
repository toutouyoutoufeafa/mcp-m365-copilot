import { Client } from "@microsoft/microsoft-graph-client";
import { EntraAuthManager, TokenResult } from "../auth/entraAuth.js";
import { RateLimiter, TokenBucketRateLimiter } from "../security/rateLimiter.js";
import { CircuitBreaker } from "../security/circuitBreaker.js";
import { logger, createCorrelationId, hashUserId } from "../utils/logger.js";
import { scanForSecrets, sanitizeText } from "../dlp/secretScanner.js";
import { AgentConfig, AgentType, MCPTool, AuditLogEntry } from "../mcp/types.js";

export interface ToolResult {
  success: boolean;
  data?: unknown;
  error?: string;
  auditEntry: Omit<AuditLogEntry, "timestamp">;
}

export abstract class BaseAgent {
  protected readonly config: AgentConfig;
  protected readonly authManager: EntraAuthManager;
  protected readonly graphClient: Client;
  protected readonly rateLimiter: RateLimiter;
  protected readonly tokenBucket: TokenBucketRateLimiter;
  protected readonly circuitBreaker: CircuitBreaker;
  protected readonly agentType: AgentType;

  constructor(config: AgentConfig, authManager: EntraAuthManager, agentType: AgentType) {
    this.config = config;
    this.authManager = authManager;
    this.agentType = agentType;
    this.rateLimiter = new RateLimiter(config.rateLimitPerMinute, 60000);
    this.tokenBucket = new TokenBucketRateLimiter(config.rateLimitPerMinute, config.rateLimitPerMinute / 60);
    this.circuitBreaker = new CircuitBreaker(5, 30000);

    this.graphClient = Client.init({
      authProvider: async (done) => {
        try {
          const token = await this.authManager.getToken(agentType);
          done(null, token);
        } catch (error) {
          done(error as Error, null);
        }
      },
    });
  }

  abstract getTools(): MCPTool[];

  abstract executeTool(name: string, args: Record<string, unknown>): Promise<ToolResult>;

  async initialize(): Promise<void> {
    await this.authManager.initialize(this.agentType);
  }

  protected async executeWithGuardrails<T>(
    toolName: string,
    args: Record<string, unknown>,
    operation: () => Promise<T>
  ): Promise<ToolResult> {
    const correlationId = createCorrelationId();
    const startTime = Date.now();
    const userIdHash = hashUserId(correlationId);

    const auditBase: Omit<AuditLogEntry, "timestamp"> = {
      correlationId,
      userIdHash,
      tool: `${this.agentType}.${toolName}`,
      action: toolName,
      result: "error",
      durationMs: 0,
      agentType: this.agentType,
    };

    if (!this.rateLimiter.tryConsume(this.agentType)) {
      return {
        success: false,
        error: "Rate limit exceeded",
        auditEntry: { ...auditBase, result: "blocked", durationMs: Date.now() - startTime },
      };
    }

    if (!this.circuitBreaker.isAvailable()) {
      return {
        success: false,
        error: "Circuit breaker open",
        auditEntry: { ...auditBase, result: "blocked", durationMs: Date.now() - startTime },
      };
    }

    const validationError = this.validateInput(args);
    if (validationError) {
      return {
        success: false,
        error: validationError,
        auditEntry: { ...auditBase, result: "blocked", durationMs: Date.now() - startTime },
      };
    }

    const dlpCheck = this.checkDLP(args);
    if (dlpCheck) {
      return {
        success: false,
        error: "DLP policy violation: sensitive data detected",
        auditEntry: { ...auditBase, result: "blocked", durationMs: Date.now() - startTime },
      };
    }

    try {
      const result = await this.circuitBreaker.execute(operation);
      const durationMs = Date.now() - startTime;

      logger.audit({
        ...auditBase,
        result: "success",
        durationMs,
      });

      return {
        success: true,
        data: result,
        auditEntry: { ...auditBase, result: "success", durationMs },
      };
    } catch (error) {
      const durationMs = Date.now() - startTime;
      const errorMessage = error instanceof Error ? error.message : String(error);

      logger.audit({
        ...auditBase,
        result: "error",
        durationMs,
        metadata: { error: errorMessage },
      });

      return {
        success: false,
        error: errorMessage,
        auditEntry: { ...auditBase, result: "error", durationMs },
      };
    }
  }

  protected validateInput(args: Record<string, unknown>): string | null {
    for (const [key, value] of Object.entries(args)) {
      if (typeof value === "string") {
        if (value.length > 100000) {
          return `Input too large: ${key} exceeds 100KB`;
        }
        const secrets = scanForSecrets(value);
        if (secrets.some((s) => s.severity === "critical")) {
          return `Critical secret detected in ${key}`;
        }
      }
    }
    return null;
  }

  protected checkDLP(args: Record<string, unknown>): boolean {
    for (const value of Object.values(args)) {
      if (typeof value === "string" && scanForSecrets(value).some((s) => s.severity === "critical")) {
        return true;
      }
    }
    return false;
  }

  protected sanitizeOutput(data: unknown): unknown {
    if (typeof data === "string") {
      return sanitizeText(data);
    }
    if (Array.isArray(data)) {
      return data.map((item) => this.sanitizeOutput(item));
    }
    if (typeof data === "object" && data !== null) {
      const sanitized: Record<string, unknown> = {};
      for (const [key, value] of Object.entries(data)) {
        sanitized[key] = this.sanitizeOutput(value);
      }
      return sanitized;
    }
    return data;
  }

  getRateLimitStatus(): { remaining: number; resetTime: number } {
    return {
      remaining: this.rateLimiter.getRemaining(this.agentType),
      resetTime: this.rateLimiter.getResetTime(this.agentType),
    };
  }

  getCircuitBreakerState(): { state: string; failures: number } {
    const state = this.circuitBreaker.getState();
    return { state: state.state, failures: state.failures };
  }
}