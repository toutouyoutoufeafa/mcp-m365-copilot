import { MCPTool, ToolResult } from "../mcp/types.js";
import { BaseAgent } from "./baseAgent.js";
import { EntraAuthManager } from "../auth/entraAuth.js";
import { logger, createCorrelationId, hashUserId } from "../utils/logger.js";
import { sanitizeText } from "../dlp/secretScanner.js";

export class AutopilotAgent extends BaseAgent {
  constructor(authManager: EntraAuthManager) {
    const config = {
      type: "autopilot" as const,
      name: "Autopilot Agent",
      scopes: [
        "https://graph.microsoft.com/DeviceManagementManagedDevices.Read.All",
        "https://graph.microsoft.com/DeviceManagementManagedDevices.ReadWrite.All",
        "https://graph.microsoft.com/DeviceManagementServiceConfig.Read.All",
        "https://graph.microsoft.com/DeviceManagementServiceConfig.ReadWrite.All",
        "https://graph.microsoft.com/DeviceManagementConfiguration.Read.All",
        "https://graph.microsoft.com/DeviceManagementConfiguration.ReadWrite.All",
      ],
      requiresApproval: true,
      maxTokensPerRequest: 4000,
      rateLimitPerMinute: 30,
    };
    super(config, authManager, "autopilot");
  }

  override getTools(): MCPTool[] {
    return [
      {
        name: "autopilot_list_devices",
        description: "List Autopilot registered devices",
        inputSchema: {
          type: "object",
          properties: {
            filter: { type: "string", description: "OData filter (e.g., 'serialNumber eq \"ABC123\"')" },
            orderBy: { type: "string", description: "Order by field (e.g., 'enrollmentDateTime desc')" },
            top: { type: "number", default: 100, maximum: 1000 },
          },
        },
      },
      {
        name: "autopilot_get_device",
        description: "Get detailed Autopilot device information",
        inputSchema: {
          type: "object",
          properties: {
            deviceId: { type: "string", description: "Autopilot device ID (id)" },
          },
          required: ["deviceId"],
        },
      },
      {
        name: "autopilot_import_devices",
        description: "Import devices for Autopilot registration (requires approval)",
        inputSchema: {
          type: "object",
          properties: {
            devices: {
              type: "array",
              items: {
                type: "object",
                properties: {
                  serialNumber: { type: "string" },
                  hardwareHash: { type: "string" },
                  productKey: { type: "string" },
                  groupTag: { type: "string" },
                  assignedUserPrincipalName: { type: "string" },
                },
                required: ["serialNumber", "hardwareHash"],
              },
            },
            previewOnly: { type: "boolean", default: true },
          },
          required: ["devices"],
        },
      },
      {
        name: "autopilot_delete_device",
        description: "Delete Autopilot device registration (requires approval)",
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
        name: "autopilot_list_profiles",
        description: "List Autopilot deployment profiles",
        inputSchema: {
          type: "object",
          properties: {
            top: { type: "number", default: 50 },
          },
        },
      },
      {
        name: "autopilot_get_profile",
        description: "Get Autopilot deployment profile details",
        inputSchema: {
          type: "object",
          properties: {
            profileId: { type: "string" },
          },
          required: ["profileId"],
        },
      },
      {
        name: "autopilot_create_profile",
        description: "Create Autopilot deployment profile (requires approval)",
        inputSchema: {
          type: "object",
          properties: {
            displayName: { type: "string" },
            description: { type: "string" },
            language: { type: "string", default: "en-US" },
            locale: { type: "string", default: "en-US" },
            deviceNameTemplate: { type: "string", description: "e.g., 'PC-%SERIAL%'" },
            joinType: { type: "string", enum: ["azureADJoin", "hybridAzureADJoin"], default: "azureADJoin" },
            oobeSettings: { type: "object" },
            previewOnly: { type: "boolean", default: true },
          },
          required: ["displayName"],
        },
      },
      {
        name: "autopilot_list_enrollment_status",
        description: "List device enrollment status for Autopilot",
        inputSchema: {
          type: "object",
          properties: {
            deviceId: { type: "string" },
            top: { type: "number", default: 50 },
          },
        },
      },
      {
        name: "autopilot_reset_device",
        description: "Reset device via Autopilot (wipe/reprovision) - requires approval",
        inputSchema: {
          type: "object",
          properties: {
            deviceId: { type: "string" },
            keepEnrollmentData: { type: "boolean", default: false },
            previewOnly: { type: "boolean", default: true },
          },
          required: ["deviceId"],
        },
      },
      {
        name: "autopilot_get_hardware_hash",
        description: "Get hardware hash for a device (for manual registration)",
        inputSchema: {
          type: "object",
          properties: {
            deviceId: { type: "string" },
          },
          required: ["deviceId"],
        },
      },
    ];
  }

  async executeTool(name: string, args: Record<string, unknown>): Promise<ToolResult> {
    const getArg = (key: string) => args[key];

    switch (name) {
      case "autopilot_list_devices":
        return this.executeWithGuardrails(name, args, () => this.listDevices(getArg("filter") as string, getArg("orderBy") as string, getArg("top") as number));
      case "autopilot_get_device":
        return this.executeWithGuardrails(name, args, () => this.getDevice(getArg("deviceId") as string));
      case "autopilot_import_devices":
        return this.executeWithGuardrails(name, args, () => this.importDevices(getArg("devices") as any[], getArg("previewOnly") as boolean));
      case "autopilot_delete_device":
        return this.executeWithGuardrails(name, args, () => this.deleteDevice(getArg("deviceId") as string, getArg("previewOnly") as boolean));
      case "autopilot_list_profiles":
        return this.executeWithGuardrails(name, args, () => this.listProfiles(getArg("top") as number));
      case "autopilot_get_profile":
        return this.executeWithGuardrails(name, args, () => this.getProfile(getArg("profileId") as string));
      case "autopilot_create_profile":
        return this.executeWithGuardrails(name, args, () => this.createProfile(getArg("displayName") as string, getArg("description") as string, getArg("language") as string, getArg("locale") as string, getArg("deviceNameTemplate") as string, getArg("joinType") as string, getArg("oobeSettings") as any, getArg("previewOnly") as boolean));
      case "autopilot_list_enrollment_status":
        return this.executeWithGuardrails(name, args, () => this.listEnrollmentStatus(getArg("deviceId") as string, getArg("top") as number));
      case "autopilot_reset_device":
        return this.executeWithGuardrails(name, args, () => this.resetDevice(getArg("deviceId") as string, getArg("keepEnrollmentData") as boolean, getArg("previewOnly") as boolean));
      case "autopilot_get_hardware_hash":
        return this.executeWithGuardrails(name, args, () => this.getHardwareHash(getArg("deviceId") as string));
      default:
        return {
          success: false,
          error: `Unknown tool: ${name}`,
          auditEntry: {
            correlationId: createCorrelationId(),
            userIdHash: hashUserId(createCorrelationId()),
            tool: `autopilot.${name}`,
            action: name,
            result: "error",
            durationMs: 0,
            agentType: "autopilot",
          },
        };
    }
  }

  private async listDevices(filter?: string, orderBy?: string, top: number = 100): Promise<any> {
    const params = new URLSearchParams({ $top: top.toString() });
    if (filter) params.append("$filter", filter);
    if (orderBy) params.append("$orderby", orderBy);
    return this.graphGet(`/deviceManagement/windowsAutopilotDeviceIdentities?${params}`);
  }

  private async getDevice(deviceId: string): Promise<any> {
    return this.graphGet(`/deviceManagement/windowsAutopilotDeviceIdentities/${deviceId}`);
  }

  private async importDevices(devices: any[], previewOnly: boolean = true): Promise<any> {
    if (previewOnly) {
      return { preview: true, devices: devices.map(d => ({ serialNumber: d.serialNumber, groupTag: d.groupTag })) };
    }
    return this.graphPost("/deviceManagement/windowsAutopilotDeviceIdentities/import", { devices });
  }

  private async deleteDevice(deviceId: string, previewOnly: boolean = true): Promise<any> {
    if (previewOnly) return { preview: true, action: "delete", deviceId };
    return this.graphDelete(`/deviceManagement/windowsAutopilotDeviceIdentities/${deviceId}`);
  }

  private async listProfiles(top: number = 50): Promise<any> {
    return this.graphGet(`/deviceManagement/windowsAutopilotDeploymentProfiles?$top=${top}`);
  }

  private async getProfile(profileId: string): Promise<any> {
    return this.graphGet(`/deviceManagement/windowsAutopilotDeploymentProfiles/${profileId}`);
  }

  private async createProfile(displayName: string, description: string, language: string, locale: string, deviceNameTemplate: string, joinType: string, oobeSettings: any, previewOnly: boolean): Promise<any> {
    const profile = {
      displayName,
      description,
      language,
      locale,
      deviceNameTemplate,
      joinType,
      outOfBoxExperienceSettings: oobeSettings || {
        hidePrivacySettings: true,
        hideEULA: true,
        userType: "standard",
        deviceUsageType: "shared",
        skipKeyboardSelectionPage: true,
      },
    };
    if (previewOnly) return { preview: true, profile };
    return this.graphPost("/deviceManagement/windowsAutopilotDeploymentProfiles", profile);
  }

  private async listEnrollmentStatus(deviceId?: string, top: number = 50): Promise<any> {
    const params = new URLSearchParams({ $top: top.toString() });
    if (deviceId) params.append("$filter", `deviceId eq '${deviceId}'`);
    return this.graphGet(`/deviceManagement/windowsAutopilotDeviceIdentities?${params}&$select=id,serialNumber,enrollmentState,lastContactedDateTime`);
  }

  private async resetDevice(deviceId: string, keepEnrollmentData: boolean, previewOnly: boolean): Promise<any> {
    if (previewOnly) return { preview: true, action: "reset", deviceId, keepEnrollmentData };
    return this.graphPost(`/deviceManagement/windowsAutopilotDeviceIdentities/${deviceId}/reset`, { keepEnrollmentData });
  }

  private async getHardwareHash(deviceId: string): Promise<any> {
    return this.graphGet(`/deviceManagement/windowsAutopilotDeviceIdentities/${deviceId}?$select=hardwareHash`);
  }

  private async graphGet(path: string): Promise<any> {
    const token = await this.authManager.getToken("autopilot");
    const response = await fetch(`https://graph.microsoft.com/beta${path}`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!response.ok) throw new Error(`Graph GET failed: ${response.status} ${response.statusText}`);
    return response.json();
  }

  private async graphPost(path: string, body: any): Promise<any> {
    const token = await this.authManager.getToken("autopilot");
    const response = await fetch(`https://graph.microsoft.com/beta${path}`, {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    if (!response.ok) throw new Error(`Graph POST failed: ${response.status} ${response.statusText}`);
    return response.json();
  }

  private async graphDelete(path: string): Promise<any> {
    const token = await this.authManager.getToken("autopilot");
    const response = await fetch(`https://graph.microsoft.com/beta${path}`, {
      method: "DELETE",
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!response.ok) throw new Error(`Graph DELETE failed: ${response.status} ${response.statusText}`);
    return { success: true };
  }
}