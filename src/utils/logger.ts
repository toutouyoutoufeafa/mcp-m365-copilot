import { AuditLogEntry, AgentType } from "../mcp/types.js";
import { sanitizeText } from "../dlp/secretScanner.js";

const LOG_LEVELS = ["debug", "info", "warn", "error"] as const;
type LogLevel = (typeof LOG_LEVELS)[number];

interface LogEntry {
  level: LogLevel;
  message: string;
  timestamp: string;
  correlationId?: string;
  metadata?: Record<string, unknown>;
}

class SecureLogger {
  private readonly minLevel: LogLevel;
  private readonly auditLogs: AuditLogEntry[] = [];
  private readonly maxAuditLogs = 10000;

  constructor(minLevel: LogLevel = "info") {
    this.minLevel = minLevel;
  }

  private shouldLog(level: LogLevel): boolean {
    return LOG_LEVELS.indexOf(level) >= LOG_LEVELS.indexOf(this.minLevel);
  }

  private formatMessage(message: string, metadata?: Record<string, unknown>): string {
    let formatted = message;
    if (metadata) {
      const sanitizedMeta = this.sanitizeMetadata(metadata);
      formatted += ` ${JSON.stringify(sanitizedMeta)}`;
    }
    return sanitizeText(formatted);
  }

  private sanitizeMetadata(metadata: Record<string, unknown>): Record<string, unknown> {
    const sanitized: Record<string, unknown> = {};
    const sensitiveKeys = [
      "token",
      "accessToken",
      "refreshToken",
      "clientSecret",
      "password",
      "secret",
      "key",
      "authorization",
      "cookie",
      "session",
      "email",
      "userPrincipalName",
      "upn",
      "mail",
      "content",
      "body",
      "prompt",
      "query",
      "searchQuery",
    ];

    for (const [key, value] of Object.entries(metadata)) {
      const lowerKey = key.toLowerCase();
      if (sensitiveKeys.some((sk) => lowerKey.includes(sk))) {
        sanitized[key] = "[REDACTED]";
      } else if (typeof value === "string") {
        sanitized[key] = sanitizeText(value);
      } else if (typeof value === "object" && value !== null) {
        sanitized[key] = this.sanitizeMetadata(value as Record<string, unknown>);
      } else {
        sanitized[key] = value;
      }
    }
    return sanitized;
  }

  private write(level: LogLevel, message: string, metadata?: Record<string, unknown>): void {
    if (!this.shouldLog(level)) return;

    const entry: LogEntry = {
      level,
      message: this.formatMessage(message, metadata),
      timestamp: new Date().toISOString(),
      correlationId: (metadata as Record<string, unknown>)?.["correlationId"] as string | undefined,
      metadata: metadata ? this.sanitizeMetadata(metadata) : undefined,
    };

    const output = JSON.stringify(entry);
    if (level === "error" || level === "warn") {
      console.error(output);
    } else {
      console.log(output);
    }
  }

  debug(message: string, metadata?: Record<string, unknown>): void {
    this.write("debug", message, metadata);
  }

  info(message: string, metadata?: Record<string, unknown>): void {
    this.write("info", message, metadata);
  }

  warn(message: string, metadata?: Record<string, unknown>): void {
    this.write("warn", message, metadata);
  }

  error(message: string, metadata?: Record<string, unknown>): void {
    this.write("error", message, metadata);
  }

  audit(entry: Omit<AuditLogEntry, "timestamp">): void {
    const auditEntry: AuditLogEntry = {
      ...entry,
      timestamp: new Date().toISOString(),
    };

    this.auditLogs.push(auditEntry);
    if (this.auditLogs.length > this.maxAuditLogs) {
      this.auditLogs.shift();
    }

    console.log(JSON.stringify(auditEntry));
  }

  getAuditLogs(since?: Date, limit?: number): AuditLogEntry[] {
    let logs = this.auditLogs;
    if (since) {
      logs = logs.filter((l) => new Date(l.timestamp) >= since);
    }
    if (limit) {
      logs = logs.slice(-limit);
    }
    return [...logs];
  }

  clearAuditLogs(): void {
    this.auditLogs.length = 0;
  }
}

export const logger = new SecureLogger((process.env["LOG_LEVEL"] as LogLevel) || "info");

export function createCorrelationId(): string {
  return `corr_${Date.now()}_${Math.random().toString(36).slice(2, 11)}`;
}

export function hashUserId(userId: string): string {
  let hash = 0;
  for (let i = 0; i < userId.length; i++) {
    const char = userId.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash |= 0;
  }
  return `user_${Math.abs(hash).toString(36)}`;
}