import { MCPTool, ToolResult, PreviewPlan, AgentType } from "../mcp/types.js";
import { BaseAgent } from "./baseAgent.js";
import { EntraAuthManager } from "../auth/entraAuth.js";
import { logger, createCorrelationId, hashUserId } from "../utils/logger.js";
import { sanitizeText } from "../dlp/secretScanner.js";
import { approvalWorkflow, ApprovalRequest } from "../approval/approvalWorkflow.js";
import { Client } from "@microsoft/microsoft-graph-client";

export class CollaborationAgent extends BaseAgent {

  constructor(authManager: EntraAuthManager) {
    const config = {
      type: "collaboration" as const,
      name: "Collaboration Agent",
      scopes: ["User.Read", "Files.Read", "Team.ReadBasic.All", "Channel.ReadBasic.All"],
      requiresApproval: true,
      maxTokensPerRequest: 4000,
      rateLimitPerMinute: 20,
    };
    super(config, authManager, "collaboration");
  }

  override getTools(): MCPTool[] {
    return [
      {
        name: "collab_preview_teams_message",
        description: "Create a preview plan for posting a message to a Teams channel (requires human approval)",
        inputSchema: {
          type: "object",
          properties: {
            teamId: { type: "string", description: "Team ID" },
            channelId: { type: "string", description: "Channel ID" },
            message: { type: "string", description: "Message content" },
            ttlMinutes: { type: "number", description: "Preview TTL in minutes", default: 60 },
          },
          required: ["teamId", "channelId", "message"],
        },
      },
      {
        name: "collab_preview_teams_channel",
        description: "Create a preview plan for creating a Teams channel (requires human approval)",
        inputSchema: {
          type: "object",
          properties: {
            teamId: { type: "string", description: "Team ID" },
            displayName: { type: "string", description: "Channel display name" },
            description: { type: "string", description: "Channel description" },
            membershipType: { type: "string", enum: ["standard", "private"], default: "standard" },
            ttlMinutes: { type: "number", description: "Preview TTL in minutes", default: 60 },
          },
          required: ["teamId", "displayName"],
        },
      },
      {
        name: "collab_preview_sharepoint_file",
        description: "Create a preview plan for uploading a file to SharePoint (requires human approval)",
        inputSchema: {
          type: "object",
          properties: {
            siteId: { type: "string", description: "SharePoint site ID" },
            driveId: { type: "string", description: "Drive ID" },
            folderPath: { type: "string", description: "Folder path in drive" },
            fileName: { type: "string", description: "File name" },
            fileContent: { type: "string", description: "File content (base64 encoded)" },
            ttlMinutes: { type: "number", description: "Preview TTL in minutes", default: 60 },
          },
          required: ["siteId", "driveId", "folderPath", "fileName", "fileContent"],
        },
      },
      {
        name: "collab_list_teams",
        description: "List Teams the user has access to (read-only, no approval needed)",
        inputSchema: {
          type: "object",
          properties: {
            maxResults: { type: "number", default: 50 },
          },
        },
      },
      {
        name: "collab_list_channels",
        description: "List channels in a Team (read-only, no approval needed)",
        inputSchema: {
          type: "object",
          properties: {
            teamId: { type: "string", description: "Team ID" },
            maxResults: { type: "number", default: 50 },
          },
          required: ["teamId"],
        },
      },
      {
        name: "collab_get_preview_plan",
        description: "Get details of a preview plan by ID",
        inputSchema: {
          type: "object",
          properties: {
            planId: { type: "string", description: "Preview plan ID" },
          },
          required: ["planId"],
        },
      },
      {
        name: "collab_approve_plan",
        description: "Approve a preview plan for execution (human approval required)",
        inputSchema: {
          type: "object",
          properties: {
            planId: { type: "string", description: "Preview plan ID" },
            approverId: { type: "string", description: "Approver identifier" },
          },
          required: ["planId", "approverId"],
        },
      },
      {
        name: "collab_reject_plan",
        description: "Reject a preview plan",
        inputSchema: {
          type: "object",
          properties: {
            planId: { type: "string", description: "Preview plan ID" },
            approverId: { type: "string", description: "Approver identifier" },
            reason: { type: "string", description: "Rejection reason" },
          },
          required: ["planId", "approverId"],
        },
      },
      {
        name: "collab_execute_plan",
        description: "Execute an approved preview plan (validates hash integrity)",
        inputSchema: {
          type: "object",
          properties: {
            planId: { type: "string", description: "Preview plan ID" },
          },
          required: ["planId"],
        },
      },
      {
        name: "collab_list_pending_approvals",
        description: "List all pending preview plans awaiting approval",
        inputSchema: { type: "object", properties: {} },
      },
    ];
  }

  async executeTool(name: string, args: Record<string, unknown>): Promise<ToolResult> {
    const getArg = (key: string) => args[key];

    switch (name) {
      case "collab_preview_teams_message":
        return this.executeWithGuardrails(name, args, () =>
          this.createTeamsMessagePreview(getArg("teamId") as string, getArg("channelId") as string, getArg("message") as string, (getArg("ttlMinutes") as number) || 60)
        );
      case "collab_preview_teams_channel":
        return this.executeWithGuardrails(name, args, () =>
          this.createTeamsChannelPreview(getArg("teamId") as string, getArg("displayName") as string, (getArg("description") as string) || "", (getArg("membershipType") as string) || "standard", (getArg("ttlMinutes") as number) || 60)
        );
      case "collab_preview_sharepoint_file":
        return this.executeWithGuardrails(name, args, () =>
          this.createSharePointFilePreview(getArg("siteId") as string, getArg("driveId") as string, getArg("folderPath") as string, getArg("fileName") as string, getArg("fileContent") as string, (getArg("ttlMinutes") as number) || 60)
        );
      case "collab_list_teams":
        return this.executeWithGuardrails(name, args, () => this.listTeams((getArg("maxResults") as number) || 50));
      case "collab_list_channels":
        return this.executeWithGuardrails(name, args, () => this.listChannels(getArg("teamId") as string, (getArg("maxResults") as number) || 50));
      case "collab_get_preview_plan":
        return this.executeWithGuardrails(name, args, () => this.getPreviewPlan(getArg("planId") as string));
      case "collab_approve_plan":
        return this.executeWithGuardrails(name, args, () => this.approvePlan(getArg("planId") as string, getArg("approverId") as string));
      case "collab_reject_plan":
        return this.executeWithGuardrails(name, args, () => this.rejectPlan(getArg("planId") as string, getArg("approverId") as string, getArg("reason") as string));
      case "collab_execute_plan":
        return this.executeWithGuardrails(name, args, () => this.executePlan(getArg("planId") as string));
      case "collab_list_pending_approvals":
        return this.executeWithGuardrails(name, args, () => this.listPendingApprovals());
      default:
        return {
          success: false,
          error: `Unknown tool: ${name}`,
          auditEntry: {
            correlationId: createCorrelationId(),
            userIdHash: hashUserId(createCorrelationId()),
            tool: `collaboration.${name}`,
            action: name,
            result: "error",
            durationMs: 0,
            agentType: "collaboration",
          },
        };
    }
  }

  private async createTeamsMessagePreview(
    teamId: string,
    channelId: string,
    message: string,
    ttlMinutes: number = 60
  ): Promise<PreviewPlan> {
    const sanitizedMessage = sanitizeText(message);
    const plan = approvalWorkflow.createPreviewPlan(
      "collaboration" as AgentType,
      "post_teams_message",
      `teams/${teamId}/channels/${channelId}`,
      { teamId, channelId, message: sanitizedMessage },
      ttlMinutes * 60 * 1000
    );
    return plan;
  }

  private async createTeamsChannelPreview(
    teamId: string,
    displayName: string,
    description: string = "",
    membershipType: string = "standard",
    ttlMinutes: number = 60
  ): Promise<PreviewPlan> {
    const plan = approvalWorkflow.createPreviewPlan(
      "collaboration" as AgentType,
      "create_teams_channel",
      `teams/${teamId}/channels`,
      { teamId, displayName, description, membershipType },
      ttlMinutes * 60 * 1000
    );
    return plan;
  }

  private async createSharePointFilePreview(
    siteId: string,
    driveId: string,
    folderPath: string,
    fileName: string,
    fileContent: string,
    ttlMinutes: number = 60
  ): Promise<PreviewPlan> {
    const plan = approvalWorkflow.createPreviewPlan(
      "collaboration" as AgentType,
      "upload_sharepoint_file",
      `sites/${siteId}/drives/${driveId}/${folderPath}/${fileName}`,
      { siteId, driveId, folderPath, fileName, fileSize: fileContent.length },
      ttlMinutes * 60 * 1000
    );
    return plan;
  }

  private async listTeams(maxResults: number = 50): Promise<{ value: Array<{ id: string; displayName: string; description?: string }> }> {
    const response = await this.graphClient.api("/me/joinedTeams").top(maxResults).get();
    return { value: response.value.map((t: any) => ({ id: t.id, displayName: t.displayName, description: t.description })) };
  }

  private async listChannels(teamId: string, maxResults: number = 50): Promise<{ value: Array<{ id: string; displayName: string; description?: string }> }> {
    const response = await this.graphClient.api(`/teams/${teamId}/channels`).top(maxResults).get();
    return { value: response.value.map((c: any) => ({ id: c.id, displayName: c.displayName, description: c.description })) };
  }

  private async getPreviewPlan(planId: string): Promise<PreviewPlan | null> {
    return approvalWorkflow.getPlan(planId) || null;
  }

  private async approvePlan(planId: string, approverId: string): Promise<PreviewPlan> {
    return approvalWorkflow.requestApproval(planId, approverId);
  }

  private async rejectPlan(planId: string, approverId: string, reason?: string): Promise<void> {
    await approvalWorkflow.rejectApproval(planId, approverId, reason);
  }

  private async executePlan(planId: string): Promise<{ executed: boolean; plan: PreviewPlan }> {
    const plan = await approvalWorkflow.executePlan(planId);

    switch (plan.action) {
      case "post_teams_message":
        await this.executeTeamsMessage(plan.payload);
        break;
      case "create_teams_channel":
        await this.executeTeamsChannel(plan.payload);
        break;
      case "upload_sharepoint_file":
        await this.executeSharePointUpload(plan.payload);
        break;
      default:
        throw new Error(`Unknown action: ${plan.action}`);
    }

    return { executed: true, plan };
  }

  private async executeTeamsMessage(payload: Record<string, unknown>): Promise<void> {
    const { teamId, channelId, message } = payload as { teamId: string; channelId: string; message: string };
    await this.graphClient.api(`/teams/${teamId}/channels/${channelId}/messages`).post({ body: { content: message } });
  }

  private async executeTeamsChannel(payload: Record<string, unknown>): Promise<void> {
    const { teamId, displayName, description, membershipType } = payload as {
      teamId: string;
      displayName: string;
      description?: string;
      membershipType: string;
    };
    await this.graphClient.api(`/teams/${teamId}/channels`).post({
      displayName,
      description,
      membershipType,
    });
  }

  private async executeSharePointUpload(payload: Record<string, unknown>): Promise<void> {
    const { siteId, driveId, folderPath, fileName } = payload as {
      siteId: string;
      driveId: string;
      folderPath: string;
      fileName: string;
    };
    await this.graphClient
      .api(`/sites/${siteId}/drives/${driveId}/root:${folderPath}/${fileName}:/content`)
      .put(new Uint8Array(0));
  }

  private async listPendingApprovals(): Promise<ApprovalRequest[]> {
    return approvalWorkflow.getPendingApprovals();
  }
}