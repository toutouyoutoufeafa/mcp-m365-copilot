import { PreviewPlan, AgentType } from "../mcp/types.js";
import { logger, createCorrelationId, hashUserId } from "../utils/logger.js";
import { scanForSecrets, sanitizeText } from "../dlp/secretScanner.js";
import { createHash } from "crypto";

export interface ApprovalRequest {
  requestId: string;
  plan: PreviewPlan;
  requesterHash: string;
  requestedAt: string;
  expiresAt: string;
}

export class ApprovalWorkflow {
  private readonly pendingApprovals: Map<string, ApprovalRequest> = new Map();
  private readonly approvedPlans: Map<string, PreviewPlan> = new Map();
  private readonly executedPlans: Set<string> = new Set();
  private readonly maxPendingAgeMs = 3600000; // 1 hour
  private readonly planTtlMs = 86400000; // 24 hours

  createPreviewPlan(
    agentType: AgentType,
    action: string,
    target: string,
    payload: Record<string, unknown>,
    ttlMs?: number
  ): PreviewPlan {
    const correlationId = createCorrelationId();
    const planId = `plan_${Date.now()}_${Math.random().toString(36).slice(2, 11)}`;
    const payloadStr = JSON.stringify(payload);
    const hash = createHash("sha256").update(payloadStr).digest("hex");
    const now = new Date();
    const expiresAt = new Date(now.getTime() + (ttlMs || this.planTtlMs));

    const sanitizedPayload = this.sanitizePayload(payload);

    const plan: PreviewPlan = {
      planId,
      hash,
      agentType,
      action,
      target,
      payload: sanitizedPayload,
      expiresAt: expiresAt.toISOString(),
      createdAt: now.toISOString(),
      status: "pending",
    };

    const request: ApprovalRequest = {
      requestId: `req_${planId}`,
      plan,
      requesterHash: hashUserId(correlationId),
      requestedAt: now.toISOString(),
      expiresAt: expiresAt.toISOString(),
    };

    this.pendingApprovals.set(planId, request);

    logger.audit({
      correlationId,
      userIdHash: request.requesterHash,
      tool: "approval_workflow",
      action: "create_preview",
      result: "preview",
      durationMs: 0,
      agentType,
      metadata: { planId, action, target, hash, expiresAt: plan.expiresAt },
    });

    return plan;
  }

  private sanitizePayload(payload: Record<string, unknown>): Record<string, unknown> {
    const sanitized: Record<string, unknown> = {};
    for (const [key, value] of Object.entries(payload)) {
      if (typeof value === "string") {
        const secrets = scanForSecrets(value);
        if (secrets.length > 0) {
          sanitized[key] = sanitizeText(value);
        } else {
          sanitized[key] = value;
        }
      } else if (typeof value === "object" && value !== null) {
        sanitized[key] = this.sanitizePayload(value as Record<string, unknown>);
      } else {
        sanitized[key] = value;
      }
    }
    return sanitized;
  }

  async requestApproval(planId: string, approverId: string): Promise<PreviewPlan> {
    const correlationId = createCorrelationId();
    const request = this.pendingApprovals.get(planId);

    if (!request) {
      throw new Error(`Preview plan not found: ${planId}`);
    }

    if (new Date() > new Date(request.expiresAt)) {
      request.plan.status = "expired";
      this.pendingApprovals.delete(planId);
      throw new Error(`Preview plan expired: ${planId}`);
    }

    if (request.plan.status !== "pending") {
      throw new Error(`Preview plan already ${request.plan.status}: ${planId}`);
    }

    const approverHash = hashUserId(approverId);
    request.plan.status = "approved";
    request.plan.approvedBy = approverHash;
    request.plan.approvedAt = new Date().toISOString();

    this.approvedPlans.set(planId, request.plan);
    this.pendingApprovals.delete(planId);

    logger.audit({
      correlationId,
      userIdHash: approverHash,
      tool: "approval_workflow",
      action: "approve",
      result: "success",
      durationMs: 0,
      agentType: request.plan.agentType,
      metadata: { planId, hash: request.plan.hash },
    });

    return request.plan;
  }

  async rejectApproval(planId: string, approverId: string, reason?: string): Promise<void> {
    const correlationId = createCorrelationId();
    const request = this.pendingApprovals.get(planId);

    if (!request) {
      throw new Error(`Preview plan not found: ${planId}`);
    }

    const approverHash = hashUserId(approverId);
    request.plan.status = "rejected";
    this.pendingApprovals.delete(planId);

    logger.audit({
      correlationId,
      userIdHash: approverHash,
      tool: "approval_workflow",
      action: "reject",
      result: "success",
      durationMs: 0,
      agentType: request.plan.agentType,
      metadata: { planId, reason: reason || "not specified" },
    });
  }

  async executePlan(planId: string): Promise<PreviewPlan> {
    const correlationId = createCorrelationId();
    const plan = this.approvedPlans.get(planId);

    if (!plan) {
      throw new Error(`Approved plan not found: ${planId}`);
    }

    if (plan.status !== "approved") {
      throw new Error(`Plan not approved: ${planId}`);
    }

    if (new Date() > new Date(plan.expiresAt)) {
      plan.status = "expired";
      this.approvedPlans.delete(planId);
      throw new Error(`Plan expired: ${planId}`);
    }

    if (this.executedPlans.has(planId)) {
      throw new Error(`Plan already executed: ${planId}`);
    }

    const currentHash = createHash("sha256").update(JSON.stringify(plan.payload)).digest("hex");
    if (currentHash !== plan.hash) {
      throw new Error(`Plan integrity check failed: hash mismatch for ${planId}`);
    }

    plan.status = "executed";
    this.executedPlans.add(planId);
    this.approvedPlans.delete(planId);

    logger.audit({
      correlationId,
      userIdHash: plan.approvedBy || "system",
      tool: "approval_workflow",
      action: "execute",
      result: "success",
      durationMs: 0,
      agentType: plan.agentType,
      metadata: { planId, hash: plan.hash },
    });

    return plan;
  }

  getPendingApprovals(): ApprovalRequest[] {
    this.cleanupExpired();
    return Array.from(this.pendingApprovals.values());
  }

  getApprovedPlans(): PreviewPlan[] {
    this.cleanupExpired();
    return Array.from(this.approvedPlans.values());
  }

  getPlan(planId: string): PreviewPlan | undefined {
    const pending = this.pendingApprovals.get(planId)?.plan;
    if (pending) return pending;
    return this.approvedPlans.get(planId);
  }

  isPlanExecutable(planId: string): boolean {
    const plan = this.approvedPlans.get(planId);
    if (!plan) return false;
    if (plan.status !== "approved") return false;
    if (new Date() > new Date(plan.expiresAt)) return false;
    if (this.executedPlans.has(planId)) return false;
    return true;
  }

  private cleanupExpired(): void {
    const now = new Date();
    for (const [planId, request] of this.pendingApprovals.entries()) {
      if (now > new Date(request.expiresAt)) {
        request.plan.status = "expired";
        this.pendingApprovals.delete(planId);
      }
    }
    for (const [planId, plan] of this.approvedPlans.entries()) {
      if (now > new Date(plan.expiresAt)) {
        plan.status = "expired";
        this.approvedPlans.delete(planId);
      }
    }
  }

  getStats(): { pending: number; approved: number; executed: number } {
    this.cleanupExpired();
    return {
      pending: this.pendingApprovals.size,
      approved: this.approvedPlans.size,
      executed: this.executedPlans.size,
    };
  }
}

export const approvalWorkflow = new ApprovalWorkflow();