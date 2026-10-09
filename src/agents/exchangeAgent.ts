import { MCPTool, ToolResult } from "../mcp/types.js";
import { BaseAgent } from "./baseAgent.js";
import { EntraAuthManager } from "../auth/entraAuth.js";
import { logger, createCorrelationId, hashUserId } from "../utils/logger.js";
import { sanitizeText } from "../dlp/secretScanner.js";

export class ExchangeAgent extends BaseAgent {
  constructor(authManager: EntraAuthManager) {
    const config = {
      type: "exchange" as const,
      name: "Exchange Agent",
      scopes: [
        "https://graph.microsoft.com/Mail.Read",
        "https://graph.microsoft.com/Mail.ReadWrite",
        "https://graph.microsoft.com/Mail.Send",
        "https://graph.microsoft.com/MailboxSettings.Read",
        "https://graph.microsoft.com/MailboxSettings.ReadWrite",
        "https://graph.microsoft.com/Calendars.Read",
        "https://graph.microsoft.com/Calendars.ReadWrite",
        "https://graph.microsoft.com/Contacts.Read",
        "https://graph.microsoft.com/Contacts.ReadWrite",
        "https://graph.microsoft.com/Group.Read.All",
        "https://graph.microsoft.com/Group.ReadWrite.All",
      ],
      requiresApproval: true,
      maxTokensPerRequest: 4000,
      rateLimitPerMinute: 30,
    };
    super(config, authManager, "exchange");
  }

  override getTools(): MCPTool[] {
    return [
      {
        name: "exchange_list_mailboxes",
        description: "List mailboxes (requires Exchange Admin or Global Reader)",
        inputSchema: {
          type: "object",
          properties: {
            filter: { type: "string" },
            top: { type: "number", default: 100 },
          },
        },
      },
      {
        name: "exchange_get_mailbox",
        description: "Get mailbox details",
        inputSchema: {
          type: "object",
          properties: {
            userId: { type: "string", description: "User ID or UPN" },
          },
          required: ["userId"],
        },
      },
      {
        name: "exchange_get_mailbox_settings",
        description: "Get mailbox settings (automatic replies, signature, etc.)",
        inputSchema: {
          type: "object",
          properties: {
            userId: { type: "string" },
          },
          required: ["userId"],
        },
      },
      {
        name: "exchange_update_mailbox_settings",
        description: "Update mailbox settings (requires approval)",
        inputSchema: {
          type: "object",
          properties: {
            userId: { type: "string" },
            automaticRepliesSetting: { type: "object" },
            signature: { type: "string" },
            language: { type: "object" },
            workingHours: { type: "object" },
            previewOnly: { type: "boolean", default: true },
          },
          required: ["userId"],
        },
      },
      {
        name: "exchange_list_messages",
        description: "List messages in a mailbox folder",
        inputSchema: {
          type: "object",
          properties: {
            userId: { type: "string" },
            folderId: { type: "string", default: "inbox" },
            filter: { type: "string" },
            top: { type: "number", default: 50 },
            select: { type: "string", default: "id,subject,from,receivedDateTime,hasAttachments,importance" },
          },
          required: ["userId"],
        },
      },
      {
        name: "exchange_get_message",
        description: "Get message details including body",
        inputSchema: {
          type: "object",
          properties: {
            userId: { type: "string" },
            messageId: { type: "string" },
          },
          required: ["userId", "messageId"],
        },
      },
      {
        name: "exchange_send_mail",
        description: "Send email (requires approval)",
        inputSchema: {
          type: "object",
          properties: {
            userId: { type: "string" },
            toRecipients: { type: "array", items: { type: "object", properties: { emailAddress: { type: "object", properties: { address: { type: "string" } } } } } },
            ccRecipients: { type: "array", items: { type: "object", properties: { emailAddress: { type: "object", properties: { address: { type: "string" } } } } } },
            subject: { type: "string" },
            body: { type: "string" },
            bodyType: { type: "string", enum: ["Text", "HTML"], default: "HTML" },
            attachments: { type: "array", items: { type: "object" } },
            previewOnly: { type: "boolean", default: true },
          },
          required: ["userId", "toRecipients", "subject", "body"],
        },
      },
      {
        name: "exchange_create_folder",
        description: "Create mail folder (requires approval)",
        inputSchema: {
          type: "object",
          properties: {
            userId: { type: "string" },
            parentFolderId: { type: "string", default: "msgfolderroot" },
            displayName: { type: "string" },
            previewOnly: { type: "boolean", default: true },
          },
          required: ["userId", "displayName"],
        },
      },
      {
        name: "exchange_move_message",
        description: "Move message to folder (requires approval)",
        inputSchema: {
          type: "object",
          properties: {
            userId: { type: "string" },
            messageId: { type: "string" },
            destinationFolderId: { type: "string" },
            previewOnly: { type: "boolean", default: true },
          },
          required: ["userId", "messageId", "destinationFolderId"],
        },
      },
      {
        name: "exchange_list_groups",
        description: "List Microsoft 365 groups (Exchange groups)",
        inputSchema: {
          type: "object",
          properties: {
            filter: { type: "string" },
            top: { type: "number", default: 100 },
          },
        },
      },
      {
        name: "exchange_get_group",
        description: "Get group details including members",
        inputSchema: {
          type: "object",
          properties: {
            groupId: { type: "string" },
          },
          required: ["groupId"],
        },
      },
      {
        name: "exchange_create_group",
        description: "Create Microsoft 365 group (requires approval)",
        inputSchema: {
          type: "object",
          properties: {
            displayName: { type: "string" },
            description: { type: "string" },
            mailEnabled: { type: "boolean", default: true },
            securityEnabled: { type: "boolean", default: false },
            groupTypes: { type: "array", items: { type: "string" }, default: ["Unified"] },
            owners: { type: "array", items: { type: "string" } },
            members: { type: "array", items: { type: "string" } },
            previewOnly: { type: "boolean", default: true },
          },
          required: ["displayName"],
        },
      },
      {
        name: "exchange_add_group_members",
        description: "Add members to group (requires approval)",
        inputSchema: {
          type: "object",
          properties: {
            groupId: { type: "string" },
            memberIds: { type: "array", items: { type: "string" } },
            previewOnly: { type: "boolean", default: true },
          },
          required: ["groupId", "memberIds"],
        },
      },
      {
        name: "exchange_get_mailbox_usage",
        description: "Get mailbox usage statistics",
        inputSchema: {
          type: "object",
          properties: {
            userId: { type: "string" },
          },
          required: ["userId"],
        },
      },
      {
        name: "exchange_search_messages",
        description: "Search messages across mailbox",
        inputSchema: {
          type: "object",
          properties: {
            userId: { type: "string" },
            searchQuery: { type: "string", description: "KQL search query" },
            top: { type: "number", default: 50 },
          },
          required: ["userId", "searchQuery"],
        },
      },
      {
        name: "exchange_list_delegates",
        description: "List mailbox delegates (SendAs, SendOnBehalf, FullAccess)",
        inputSchema: {
          type: "object",
          properties: {
            userId: { type: "string" },
          },
          required: ["userId"],
        },
      },
    ];
  }

  async executeTool(name: string, args: Record<string, unknown>): Promise<ToolResult> {
    const getArg = (key: string) => args[key];

    switch (name) {
      case "exchange_list_mailboxes":
        return this.executeWithGuardrails(name, args, () => this.listMailboxes(getArg("filter") as string, getArg("top") as number));
      case "exchange_get_mailbox":
        return this.executeWithGuardrails(name, args, () => this.getMailbox(getArg("userId") as string));
      case "exchange_get_mailbox_settings":
        return this.executeWithGuardrails(name, args, () => this.getMailboxSettings(getArg("userId") as string));
      case "exchange_update_mailbox_settings":
        return this.executeWithGuardrails(name, args, () => this.updateMailboxSettings(getArg("userId") as string, getArg("automaticRepliesSetting") as any, getArg("signature") as string, getArg("language") as any, getArg("workingHours") as any, getArg("previewOnly") as boolean));
      case "exchange_list_messages":
        return this.executeWithGuardrails(name, args, () => this.listMessages(getArg("userId") as string, getArg("folderId") as string, getArg("filter") as string, getArg("top") as number, getArg("select") as string));
      case "exchange_get_message":
        return this.executeWithGuardrails(name, args, () => this.getMessage(getArg("userId") as string, getArg("messageId") as string));
      case "exchange_send_mail":
        return this.executeWithGuardrails(name, args, () => this.sendMail(getArg("userId") as string, getArg("toRecipients") as any[], getArg("ccRecipients") as any[], getArg("subject") as string, getArg("body") as string, getArg("bodyType") as string, getArg("attachments") as any[], getArg("previewOnly") as boolean));
      case "exchange_create_folder":
        return this.executeWithGuardrails(name, args, () => this.createFolder(getArg("userId") as string, getArg("parentFolderId") as string, getArg("displayName") as string, getArg("previewOnly") as boolean));
      case "exchange_move_message":
        return this.executeWithGuardrails(name, args, () => this.moveMessage(getArg("userId") as string, getArg("messageId") as string, getArg("destinationFolderId") as string, getArg("previewOnly") as boolean));
      case "exchange_list_groups":
        return this.executeWithGuardrails(name, args, () => this.listGroups(getArg("filter") as string, getArg("top") as number));
      case "exchange_get_group":
        return this.executeWithGuardrails(name, args, () => this.getGroup(getArg("groupId") as string));
      case "exchange_create_group":
        return this.executeWithGuardrails(name, args, () => this.createGroup(getArg("displayName") as string, getArg("description") as string, getArg("mailEnabled") as boolean, getArg("securityEnabled") as boolean, getArg("groupTypes") as string[], getArg("owners") as string[], getArg("members") as string[], getArg("previewOnly") as boolean));
      case "exchange_add_group_members":
        return this.executeWithGuardrails(name, args, () => this.addGroupMembers(getArg("groupId") as string, getArg("memberIds") as string[], getArg("previewOnly") as boolean));
      case "exchange_get_mailbox_usage":
        return this.executeWithGuardrails(name, args, () => this.getMailboxUsage(getArg("userId") as string));
      case "exchange_search_messages":
        return this.executeWithGuardrails(name, args, () => this.searchMessages(getArg("userId") as string, getArg("searchQuery") as string, getArg("top") as number));
      case "exchange_list_delegates":
        return this.executeWithGuardrails(name, args, () => this.listDelegates(getArg("userId") as string));
      default:
        return {
          success: false,
          error: `Unknown tool: ${name}`,
          auditEntry: {
            correlationId: createCorrelationId(),
            userIdHash: hashUserId(createCorrelationId()),
            tool: `exchange.${name}`,
            action: name,
            result: "error",
            durationMs: 0,
            agentType: "exchange",
          },
        };
    }
  }

  private async listMailboxes(filter?: string, top: number = 100): Promise<any> {
    const params = new URLSearchParams({ $top: top.toString() });
    if (filter) params.append("$filter", filter);
    return this.graphGet(`/users?${params}&$select=id,displayName,userPrincipalName,mail,mailNickname`);
  }

  private async getMailbox(userId: string): Promise<any> {
    return this.graphGet(`/users/${userId}?$select=id,displayName,userPrincipalName,mail,mailNickname,jobTitle,department,officeLocation`);
  }

  private async getMailboxSettings(userId: string): Promise<any> {
    return this.graphGet(`/users/${userId}/mailboxSettings`);
  }

  private async updateMailboxSettings(userId: string, automaticRepliesSetting?: any, signature?: string, language?: any, workingHours?: any, previewOnly: boolean = true): Promise<any> {
    const updates: any = {};
    if (automaticRepliesSetting) updates.automaticRepliesSetting = automaticRepliesSetting;
    if (signature) updates.signature = signature;
    if (language) updates.language = language;
    if (workingHours) updates.workingHours = workingHours;
    if (previewOnly) return { preview: true, userId, updates };
    return this.graphPatch(`/users/${userId}/mailboxSettings`, updates);
  }

  private async listMessages(userId: string, folderId: string = "inbox", filter?: string, top: number = 50, select: string = "id,subject,from,receivedDateTime,hasAttachments,importance"): Promise<any> {
    const params = new URLSearchParams({ $top: top.toString(), $select: select, $orderby: "receivedDateTime desc" });
    if (filter) params.append("$filter", filter);
    return this.graphGet(`/users/${userId}/mailFolders/${folderId}/messages?${params}`);
  }

  private async getMessage(userId: string, messageId: string): Promise<any> {
    return this.graphGet(`/users/${userId}/messages/${messageId}`);
  }

  private async sendMail(userId: string, toRecipients: any[], ccRecipients: any[], subject: string, body: string, bodyType: string = "HTML", attachments: any[] = [], previewOnly: boolean = true): Promise<any> {
    const message = {
      subject,
      body: { contentType: bodyType, content: body },
      toRecipients,
      ccRecipients: ccRecipients || [],
      attachments: attachments || [],
    };
    if (previewOnly) return { preview: true, userId, message };
    return this.graphPost(`/users/${userId}/sendMail`, { message, saveToSentItems: true });
  }

  private async createFolder(userId: string, parentFolderId: string = "msgfolderroot", displayName: string, previewOnly: boolean = true): Promise<any> {
    if (previewOnly) return { preview: true, userId, parentFolderId, displayName };
    return this.graphPost(`/users/${userId}/mailFolders`, { displayName, parentFolderId });
  }

  private async moveMessage(userId: string, messageId: string, destinationFolderId: string, previewOnly: boolean = true): Promise<any> {
    if (previewOnly) return { preview: true, userId, messageId, destinationFolderId };
    return this.graphPost(`/users/${userId}/messages/${messageId}/move`, { destinationId: destinationFolderId });
  }

  private async listGroups(filter?: string, top: number = 100): Promise<any> {
    const params = new URLSearchParams({ $top: top.toString() });
    if (filter) params.append("$filter", filter);
    return this.graphGet(`/groups?${params}`);
  }

  private async getGroup(groupId: string): Promise<any> {
    const group = await this.graphGet(`/groups/${groupId}`);
    const members = await this.graphGet(`/groups/${groupId}/members?$top=100`);
    const owners = await this.graphGet(`/groups/${groupId}/owners?$top=100`);
    return { ...group, members: members.value, owners: owners.value };
  }

  private async createGroup(displayName: string, description: string, mailEnabled: boolean = true, securityEnabled: boolean = false, groupTypes: string[] = ["Unified"], owners: string[] = [], members: string[] = [], previewOnly: boolean = true): Promise<any> {
    const group = { displayName, description, mailEnabled, securityEnabled, groupTypes };
    if (previewOnly) return { preview: true, group };
    const created = await this.graphPost("/groups", group);
    if (owners.length > 0) {
      for (const owner of owners) {
        await this.graphPost(`/groups/${created.id}/owners/$ref`, { "@odata.id": `https://graph.microsoft.com/v1.0/users/${owner}` });
      }
    }
    if (members.length > 0) {
      for (const member of members) {
        await this.graphPost(`/groups/${created.id}/members/$ref`, { "@odata.id": `https://graph.microsoft.com/v1.0/users/${member}` });
      }
    }
    return created;
  }

  private async addGroupMembers(groupId: string, memberIds: string[], previewOnly: boolean = true): Promise<any> {
    if (previewOnly) return { preview: true, groupId, memberIds };
    const results = [];
    for (const memberId of memberIds) {
      try {
        await this.graphPost(`/groups/${groupId}/members/$ref`, { "@odata.id": `https://graph.microsoft.com/v1.0/users/${memberId}` });
        results.push({ memberId, success: true });
      } catch (e) {
        results.push({ memberId, success: false, error: String(e) });
      }
    }
    return results;
  }

  private async getMailboxUsage(userId: string): Promise<any> {
    return this.graphGet(`/users/${userId}/mailboxSettings/usage`);
  }

  private async searchMessages(userId: string, searchQuery: string, top: number = 50): Promise<any> {
    return this.graphGet(`/users/${userId}/messages?$search="${searchQuery}"&$top=${top}`);
  }

  private async listDelegates(userId: string): Promise<any> {
    const permissions = await this.graphGet(`/users/${userId}/mailFolders/inbox/memberOf?$top=100`);
    return { delegates: permissions.value };
  }

  private async graphGet(path: string): Promise<any> {
    const token = await this.authManager.getToken("exchange");
    const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!response.ok) throw new Error(`Graph GET failed: ${response.status} ${response.statusText}`);
    return response.json();
  }

  private async graphPost(path: string, body: any): Promise<any> {
    const token = await this.authManager.getToken("exchange");
    const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    if (!response.ok) throw new Error(`Graph POST failed: ${response.status} ${response.statusText}`);
    return response.json();
  }

  private async graphPatch(path: string, body: any): Promise<any> {
    const token = await this.authManager.getToken("exchange");
    const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
      method: "PATCH",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    if (!response.ok) throw new Error(`Graph PATCH failed: ${response.status} ${response.statusText}`);
    return response.json();
  }
}