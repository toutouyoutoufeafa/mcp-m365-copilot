import { MCPTool, ToolResult } from "../mcp/types.js";
import { BaseAgent } from "./baseAgent.js";
import { EntraAuthManager } from "../auth/entraAuth.js";
import { logger, createCorrelationId, hashUserId } from "../utils/logger.js";
import { sanitizeText } from "../dlp/secretScanner.js";

export class DefenderAgent extends BaseAgent {
  constructor(authManager: EntraAuthManager) {
    const config = {
      type: "defender" as const,
      name: "Defender Agent",
      scopes: [
        "https://graph.microsoft.com/SecurityEvents.Read.All",
        "https://graph.microsoft.com/Alert.Read.All",
        "https://graph.microsoft.com/Alert.ReadWrite.All",
        "https://graph.microsoft.com/ThreatIndicators.Read.All",
        "https://graph.microsoft.com/ThreatIndicators.ReadWrite.OwnedBy",
        "https://graph.microsoft.com/Vulnerability.Read.All",
        "https://graph.microsoft.com/Machine.Read.All",
        "https://graph.microsoft.com/Machine.ReadWrite.All",
      ],
      requiresApproval: true,
      maxTokensPerRequest: 4000,
      rateLimitPerMinute: 30,
    };
    super(config, authManager, "defender");
  }

  override getTools(): MCPTool[] {
    return [
      {
        name: "defender_list_alerts",
        description: "List Microsoft Defender alerts",
        inputSchema: {
          type: "object",
          properties: {
            filter: { type: "string", description: "OData filter (e.g., 'severity eq \"High\"')" },
            orderBy: { type: "string", default: "createdDateTime desc" },
            top: { type: "number", default: 50, maximum: 1000 },
            status: { type: "string", enum: ["New", "InProgress", "Resolved"] },
            severity: { type: "string", enum: ["Informational", "Low", "Medium", "High"] },
          },
        },
      },
      {
        name: "defender_get_alert",
        description: "Get detailed Defender alert information",
        inputSchema: {
          type: "object",
          properties: {
            alertId: { type: "string" },
          },
          required: ["alertId"],
        },
      },
      {
        name: "defender_update_alert",
        description: "Update alert status/assignment (requires approval)",
        inputSchema: {
          type: "object",
          properties: {
            alertId: { type: "string" },
            status: { type: "string", enum: ["New", "InProgress", "Resolved"] },
            assignedTo: { type: "string", description: "User principal name" },
            classification: { type: "string", enum: ["TruePositive", "FalsePositive", "InformationalExpectedActivity"] },
            determination: { type: "string", enum: ["Malware", "Phishing", "UnwantedSoftware", "Other"] },
            comment: { type: "string" },
            previewOnly: { type: "boolean", default: true },
          },
          required: ["alertId"],
        },
      },
      {
        name: "defender_list_devices",
        description: "List devices in Defender for Endpoint",
        inputSchema: {
          type: "object",
          properties: {
            filter: { type: "string" },
            top: { type: "number", default: 100 },
            riskLevel: { type: "string", enum: ["Low", "Medium", "High"] },
          },
        },
      },
      {
        name: "defender_get_device",
        description: "Get device details from Defender",
        inputSchema: {
          type: "object",
          properties: {
            deviceId: { type: "string" },
          },
          required: ["deviceId"],
        },
      },
      {
        name: "defender_run_antivirus_scan",
        description: "Run antivirus scan on device (requires approval)",
        inputSchema: {
          type: "object",
          properties: {
            deviceId: { type: "string" },
            scanType: { type: "string", enum: ["Quick", "Full"], default: "Quick" },
            previewOnly: { type: "boolean", default: true },
          },
          required: ["deviceId"],
        },
      },
      {
        name: "defender_isolate_device",
        description: "Isolate device from network (requires approval)",
        inputSchema: {
          type: "object",
          properties: {
            deviceId: { type: "string" },
            isolationType: { type: "string", enum: ["Full", "Selective"], default: "Full" },
            previewOnly: { type: "boolean", default: true },
          },
          required: ["deviceId"],
        },
      },
      {
        name: "defender_release_device",
        description: "Release device from isolation (requires approval)",
        inputSchema: {
          type: "object",
          properties: {
            deviceId: { type: "string" },
            previewOnly: { type: "boolean", default: true },
          },
          required: ["deviceId"],
        },
      },
      {
        name: "defender_list_vulnerabilities",
        description: "List vulnerabilities found by Defender Vulnerability Management",
        inputSchema: {
          type: "object",
          properties: {
            filter: { type: "string" },
            top: { type: "number", default: 100 },
            severity: { type: "string", enum: ["Critical", "High", "Medium", "Low"] },
          },
        },
      },
      {
        name: "defender_get_vulnerability",
        description: "Get vulnerability details",
        inputSchema: {
          type: "object",
          properties: {
            vulnerabilityId: { type: "string" },
          },
          required: ["vulnerabilityId"],
        },
      },
      {
        name: "defender_list_threat_indicators",
        description: "List threat indicators (IOCs)",
        inputSchema: {
          type: "object",
          properties: {
            filter: { type: "string" },
            top: { type: "number", default: 100 },
          },
        },
      },
      {
        name: "defender_create_threat_indicator",
        description: "Create threat indicator (IOC) - requires approval",
        inputSchema: {
          type: "object",
          properties: {
            indicatorType: { type: "string", enum: ["FileHash", "IpAddress", "Url", "DomainName", "CertificateThumbprint"] },
            value: { type: "string" },
            action: { type: "string", enum: ["Alert", "AlertAndBlock", "Allow"] },
            title: { type: "string" },
            description: { type: "string" },
            severity: { type: "string", enum: ["Informational", "Low", "Medium", "High"] },
            expirationDateTime: { type: "string", description: "ISO 8601" },
            previewOnly: { type: "boolean", default: true },
          },
          required: ["indicatorType", "value", "action", "title"],
        },
      },
      {
        name: "defender_get_score",
        description: "Get Microsoft Secure Score for the tenant",
        inputSchema: { type: "object", properties: {} },
      },
      {
        name: "defender_list_incidents",
        description: "List Defender incidents",
        inputSchema: {
          type: "object",
          properties: {
            filter: { type: "string" },
            top: { type: "number", default: 50 },
            status: { type: "string", enum: ["Active", "Resolved"] },
          },
        },
      },
    ];
  }

  async executeTool(name: string, args: Record<string, unknown>): Promise<ToolResult> {
    const getArg = (key: string) => args[key];

    switch (name) {
      case "defender_list_alerts":
        return this.executeWithGuardrails(name, args, () => this.listAlerts(getArg("filter") as string, getArg("orderBy") as string, getArg("top") as number, getArg("status") as string, getArg("severity") as string));
      case "defender_get_alert":
        return this.executeWithGuardrails(name, args, () => this.getAlert(getArg("alertId") as string));
      case "defender_update_alert":
        return this.executeWithGuardrails(name, args, () => this.updateAlert(getArg("alertId") as string, getArg("status") as string, getArg("assignedTo") as string, getArg("classification") as string, getArg("determination") as string, getArg("comment") as string, getArg("previewOnly") as boolean));
      case "defender_list_devices":
        return this.executeWithGuardrails(name, args, () => this.listDevices(getArg("filter") as string, getArg("top") as number, getArg("riskLevel") as string));
      case "defender_get_device":
        return this.executeWithGuardrails(name, args, () => this.getDevice(getArg("deviceId") as string));
      case "defender_run_antivirus_scan":
        return this.executeWithGuardrails(name, args, () => this.runScan(getArg("deviceId") as string, getArg("scanType") as string, getArg("previewOnly") as boolean));
      case "defender_isolate_device":
        return this.executeWithGuardrails(name, args, () => this.isolateDevice(getArg("deviceId") as string, getArg("isolationType") as string, getArg("previewOnly") as boolean));
      case "defender_release_device":
        return this.executeWithGuardrails(name, args, () => this.releaseDevice(getArg("deviceId") as string, getArg("previewOnly") as boolean));
      case "defender_list_vulnerabilities":
        return this.executeWithGuardrails(name, args, () => this.listVulnerabilities(getArg("filter") as string, getArg("top") as number, getArg("severity") as string));
      case "defender_get_vulnerability":
        return this.executeWithGuardrails(name, args, () => this.getVulnerability(getArg("vulnerabilityId") as string));
      case "defender_list_threat_indicators":
        return this.executeWithGuardrails(name, args, () => this.listIndicators(getArg("filter") as string, getArg("top") as number));
      case "defender_create_threat_indicator":
        return this.executeWithGuardrails(name, args, () => this.createIndicator(getArg("indicatorType") as string, getArg("value") as string, getArg("action") as string, getArg("title") as string, getArg("description") as string, getArg("severity") as string, getArg("expirationDateTime") as string, getArg("previewOnly") as boolean));
      case "defender_get_score":
        return this.executeWithGuardrails(name, args, () => this.getScore());
      case "defender_list_incidents":
        return this.executeWithGuardrails(name, args, () => this.listIncidents(getArg("filter") as string, getArg("top") as number, getArg("status") as string));
      default:
        return {
          success: false,
          error: `Unknown tool: ${name}`,
          auditEntry: {
            correlationId: createCorrelationId(),
            userIdHash: hashUserId(createCorrelationId()),
            tool: `defender.${name}`,
            action: name,
            result: "error",
            durationMs: 0,
            agentType: "defender",
          },
        };
    }
  }

  private async listAlerts(filter?: string, orderBy: string = "createdDateTime desc", top: number = 50, status?: string, severity?: string): Promise<any> {
    const params = new URLSearchParams({ $top: top.toString(), $orderby: orderBy });
    const filters = [];
    if (status) filters.push(`status eq '${status}'`);
    if (severity) filters.push(`severity eq '${severity}'`);
    if (filter) filters.push(filter);
    if (filters.length > 0) params.append("$filter", filters.join(" and "));
    return this.graphGet(`/security/alerts_v2?${params}`);
  }

  private async getAlert(alertId: string): Promise<any> {
    return this.graphGet(`/security/alerts_v2/${alertId}`);
  }

  private async updateAlert(alertId: string, status?: string, assignedTo?: string, classification?: string, determination?: string, comment?: string, previewOnly: boolean = true): Promise<any> {
    const updates: any = {};
    if (status) updates.status = status;
    if (assignedTo) updates.assignedTo = assignedTo;
    if (classification) updates.classification = classification;
    if (determination) updates.determination = determination;
    if (comment) updates.comments = [{ comment }];
    if (previewOnly) return { preview: true, alertId, updates };
    return this.graphPatch(`/security/alerts_v2/${alertId}`, updates);
  }

  private async listDevices(filter?: string, top: number = 100, riskLevel?: string): Promise<any> {
    const params = new URLSearchParams({ $top: top.toString() });
    const filters = [];
    if (riskLevel) filters.push(`riskLevel eq '${riskLevel}'`);
    if (filter) filters.push(filter);
    if (filters.length > 0) params.append("$filter", filters.join(" and "));
    return this.graphGet(`/api/machines?${params}`);
  }

  private async getDevice(deviceId: string): Promise<any> {
    return this.graphGet(`/api/machines/${deviceId}`);
  }

  private async runScan(deviceId: string, scanType: string = "Quick", previewOnly: boolean = true): Promise<any> {
    if (previewOnly) return { preview: true, action: "runScan", deviceId, scanType };
    return this.graphPost(`/api/machines/${deviceId}/runAntiVirusScan`, { ScanType: scanType });
  }

  private async isolateDevice(deviceId: string, isolationType: string = "Full", previewOnly: boolean = true): Promise<any> {
    if (previewOnly) return { preview: true, action: "isolate", deviceId, isolationType };
    return this.graphPost(`/api/machines/${deviceId}/isolate`, { IsolationType: isolationType, Comment: "Isolated via MCP" });
  }

  private async releaseDevice(deviceId: string, previewOnly: boolean = true): Promise<any> {
    if (previewOnly) return { preview: true, action: "release", deviceId };
    return this.graphPost(`/api/machines/${deviceId}/unisolate`, { Comment: "Released via MCP" });
  }

  private async listVulnerabilities(filter?: string, top: number = 100, severity?: string): Promise<any> {
    const params = new URLSearchParams({ $top: top.toString() });
    const filters = [];
    if (severity) filters.push(`severity eq '${severity}'`);
    if (filter) filters.push(filter);
    if (filters.length > 0) params.append("$filter", filters.join(" and "));
    return this.graphGet(`/api/vulnerabilities?${params}`);
  }

  private async getVulnerability(vulnerabilityId: string): Promise<any> {
    return this.graphGet(`/api/vulnerabilities/${vulnerabilityId}`);
  }

  private async listIndicators(filter?: string, top: number = 100): Promise<any> {
    const params = new URLSearchParams({ $top: top.toString() });
    if (filter) params.append("$filter", filter);
    return this.graphGet(`/api/indicators?${params}`);
  }

  private async createIndicator(indicatorType: string, value: string, action: string, title: string, description: string, severity: string, expirationDateTime: string, previewOnly: boolean): Promise<any> {
    const indicator = { indicatorType, value, action, title, description, severity, expirationDateTime };
    if (previewOnly) return { preview: true, indicator };
    return this.graphPost("/api/indicators", indicator);
  }

  private async getScore(): Promise<any> {
    return this.graphGet("/security/secureScores");
  }

  private async listIncidents(filter?: string, top: number = 50, status?: string): Promise<any> {
    const params = new URLSearchParams({ $top: top.toString() });
    const filters = [];
    if (status) filters.push(`status eq '${status}'`);
    if (filter) filters.push(filter);
    if (filters.length > 0) params.append("$filter", filters.join(" and "));
    return this.graphGet(`/security/incidents?${params}`);
  }

  private async graphGet(path: string): Promise<any> {
    const token = await this.authManager.getToken("defender");
    const baseUrl = path.startsWith("/security") || path.startsWith("/api") ? "https://graph.microsoft.com/beta" : "https://api.securitycenter.microsoft.com";
    const response = await fetch(`${baseUrl}${path}`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!response.ok) throw new Error(`Defender GET failed: ${response.status} ${response.statusText}`);
    return response.json();
  }

  private async graphPost(path: string, body: any): Promise<any> {
    const token = await this.authManager.getToken("defender");
    const baseUrl = path.startsWith("/security") ? "https://graph.microsoft.com/beta" : "https://api.securitycenter.microsoft.com";
    const response = await fetch(`${baseUrl}${path}`, {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    if (!response.ok) throw new Error(`Defender POST failed: ${response.status} ${response.statusText}`);
    return response.json();
  }

  private async graphPatch(path: string, body: any): Promise<any> {
    const token = await this.authManager.getToken("defender");
    const response = await fetch(`https://graph.microsoft.com/beta${path}`, {
      method: "PATCH",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    if (!response.ok) throw new Error(`Defender PATCH failed: ${response.status} ${response.statusText}`);
    return response.json();
  }
}