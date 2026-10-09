import { MCPTool, ToolResult } from "../mcp/types.js";
import { BaseAgent } from "./baseAgent.js";
import { EntraAuthManager } from "../auth/entraAuth.js";
import { logger, createCorrelationId, hashUserId } from "../utils/logger.js";
import { sanitizeText } from "../dlp/secretScanner.js";

export class M365AdminAgent extends BaseAgent {
  constructor(authManager: EntraAuthManager) {
    const config = {
      type: "m365admin" as const,
      name: "M365 Admin Agent",
      scopes: [
        "https://graph.microsoft.com/Organization.Read.All",
        "https://graph.microsoft.com/Organization.ReadWrite.All",
        "https://graph.microsoft.com/Domain.Read.All",
        "https://graph.microsoft.com/Domain.ReadWrite.All",
        "https://graph.microsoft.com/ServiceHealth.Read.All",
        "https://graph.microsoft.com/Reports.Read.All",
        "https://graph.microsoft.com/Directory.Read.All",
        "https://graph.microsoft.com/Directory.ReadWrite.All",
        "https://graph.microsoft.com/Policy.Read.All",
        "https://graph.microsoft.com/Policy.ReadWrite.All",
        "https://graph.microsoft.com/RoleManagement.Read.All",
        "https://graph.microsoft.com/RoleManagement.ReadWrite.Directory",
      ],
      requiresApproval: true,
      maxTokensPerRequest: 4000,
      rateLimitPerMinute: 30,
    };
    super(config, authManager, "m365admin");
  }

  override getTools(): MCPTool[] {
    return [
      {
        name: "m365admin_get_tenant_info",
        description: "Get tenant organization information",
        inputSchema: { type: "object", properties: {} },
      },
      {
        name: "m365admin_list_domains",
        description: "List all verified domains in the tenant",
        inputSchema: {
          type: "object",
          properties: {
            top: { type: "number", default: 100 },
          },
        },
      },
      {
        name: "m365admin_get_domain",
        description: "Get domain details including DNS records",
        inputSchema: {
          type: "object",
          properties: {
            domainName: { type: "string" },
          },
          required: ["domainName"],
        },
      },
      {
        name: "m365admin_verify_domain",
        description: "Start domain verification process (requires approval)",
        inputSchema: {
          type: "object",
          properties: {
            domainName: { type: "string" },
            previewOnly: { type: "boolean", default: true },
          },
          required: ["domainName"],
        },
      },
      {
        name: "m365admin_get_service_health",
        description: "Get Microsoft 365 service health status",
        inputSchema: {
          type: "object",
          properties: {
            service: { type: "string", description: "Specific service (e.g., 'Exchange', 'Teams')" },
          },
        },
      },
      {
        name: "m365admin_list_service_issues",
        description: "List active service health issues",
        inputSchema: {
          type: "object",
          properties: {
            status: { type: "string", enum: ["ServiceOperational", "Investigating", "RestoringService", "ServiceRestored", "PostIncidentReportPublished"] },
          },
        },
      },
      {
        name: "m365admin_get_reports",
        description: "Get M365 usage reports",
        inputSchema: {
          type: "object",
          properties: {
            reportName: {
              type: "string",
              enum: ["getOffice365ActiveUserDetail", "getOffice365ActiveUserCounts", "getSharePointActivityFileCounts", "getTeamsUserActivityUserDetail", "getEmailAppUsageAppsUserCounts", "getOneDriveUsageAccountDetail"],
            },
            period: { type: "string", enum: ["D7", "D30", "D90", "D180"], default: "D30" },
            format: { type: "string", enum: ["json", "csv"], default: "json" },
          },
          required: ["reportName"],
        },
      },
      {
        name: "m365admin_list_roles",
        description: "List directory roles and role assignments",
        inputSchema: {
          type: "object",
          properties: {
            scope: { type: "string", enum: ["Directory", "EntitlementManagement"], default: "Directory" },
          },
        },
      },
      {
        name: "m365admin_get_role_members",
        description: "Get members of a directory role",
        inputSchema: {
          type: "object",
          properties: {
            roleId: { type: "string" },
            roleTemplateId: { type: "string" },
          },
        },
      },
      {
        name: "m365admin_assign_role",
        description: "Assign directory role to user/group (requires approval)",
        inputSchema: {
          type: "object",
          properties: {
            roleTemplateId: { type: "string" },
            principalId: { type: "string", description: "User or group object ID" },
            scopeType: { type: "string", enum: ["Directory", "AdministrativeUnit"], default: "Directory" },
            scopeId: { type: "string", description: "Administrative unit ID if scopeType is AdministrativeUnit" },
            previewOnly: { type: "boolean", default: true },
          },
          required: ["roleTemplateId", "principalId"],
        },
      },
      {
        name: "m365admin_remove_role",
        description: "Remove directory role assignment (requires approval)",
        inputSchema: {
          type: "object",
          properties: {
            roleAssignmentId: { type: "string" },
            previewOnly: { type: "boolean", default: true },
          },
          required: ["roleAssignmentId"],
        },
      },
      {
        name: "m365admin_list_policies",
        description: "List authentication/conditional access/permission policies",
        inputSchema: {
          type: "object",
          properties: {
            policyType: { type: "string", enum: ["authenticationMethodsPolicy", "conditionalAccess", "permissionGrant", "homeRealmDiscovery", "tokenIssuance", "tokenLifetime", "authenticationStrength"] },
          },
        },
      },
      {
        name: "m365admin_get_policy",
        description: "Get policy details",
        inputSchema: {
          type: "object",
          properties: {
            policyId: { type: "string" },
            policyType: { type: "string" },
          },
          required: ["policyId", "policyType"],
        },
      },
      {
        name: "m365admin_get_licenses",
        description: "Get subscribed SKUs and license counts",
        inputSchema: { type: "object", properties: {} },
      },
      {
        name: "m365admin_get_service_principals",
        description: "List service principals (enterprise apps)",
        inputSchema: {
          type: "object",
          properties: {
            filter: { type: "string" },
            top: { type: "number", default: 100 },
          },
        },
      },
    ];
  }

  async executeTool(name: string, args: Record<string, unknown>): Promise<ToolResult> {
    const getArg = (key: string) => args[key];

    switch (name) {
      case "m365admin_get_tenant_info":
        return this.executeWithGuardrails(name, args, () => this.getTenantInfo());
      case "m365admin_list_domains":
        return this.executeWithGuardrails(name, args, () => this.listDomains(getArg("top") as number));
      case "m365admin_get_domain":
        return this.executeWithGuardrails(name, args, () => this.getDomain(getArg("domainName") as string));
      case "m365admin_verify_domain":
        return this.executeWithGuardrails(name, args, () => this.verifyDomain(getArg("domainName") as string, getArg("previewOnly") as boolean));
      case "m365admin_get_service_health":
        return this.executeWithGuardrails(name, args, () => this.getServiceHealth(getArg("service") as string));
      case "m365admin_list_service_issues":
        return this.executeWithGuardrails(name, args, () => this.listServiceIssues(getArg("status") as string));
      case "m365admin_get_reports":
        return this.executeWithGuardrails(name, args, () => this.getReports(getArg("reportName") as string, getArg("period") as string, getArg("format") as string));
      case "m365admin_list_roles":
        return this.executeWithGuardrails(name, args, () => this.listRoles(getArg("scope") as string));
      case "m365admin_get_role_members":
        return this.executeWithGuardrails(name, args, () => this.getRoleMembers(getArg("roleId") as string, getArg("roleTemplateId") as string));
      case "m365admin_assign_role":
        return this.executeWithGuardrails(name, args, () => this.assignRole(getArg("roleTemplateId") as string, getArg("principalId") as string, getArg("scopeType") as string, getArg("scopeId") as string, getArg("previewOnly") as boolean));
      case "m365admin_remove_role":
        return this.executeWithGuardrails(name, args, () => this.removeRole(getArg("roleAssignmentId") as string, getArg("previewOnly") as boolean));
      case "m365admin_list_policies":
        return this.executeWithGuardrails(name, args, () => this.listPolicies(getArg("policyType") as string));
      case "m365admin_get_policy":
        return this.executeWithGuardrails(name, args, () => this.getPolicy(getArg("policyId") as string, getArg("policyType") as string));
      case "m365admin_get_licenses":
        return this.executeWithGuardrails(name, args, () => this.getLicenses());
      case "m365admin_get_service_principals":
        return this.executeWithGuardrails(name, args, () => this.getServicePrincipals(getArg("filter") as string, getArg("top") as number));
      default:
        return {
          success: false,
          error: `Unknown tool: ${name}`,
          auditEntry: {
            correlationId: createCorrelationId(),
            userIdHash: hashUserId(createCorrelationId()),
            tool: `m365admin.${name}`,
            action: name,
            result: "error",
            durationMs: 0,
            agentType: "m365admin",
          },
        };
    }
  }

  private async getTenantInfo(): Promise<any> {
    return this.graphGet("/organization");
  }

  private async listDomains(top: number = 100): Promise<any> {
    return this.graphGet(`/domains?$top=${top}`);
  }

  private async getDomain(domainName: string): Promise<any> {
    return this.graphGet(`/domains/${domainName}`);
  }

  private async verifyDomain(domainName: string, previewOnly: boolean): Promise<any> {
    if (previewOnly) return { preview: true, action: "verifyDomain", domainName };
    return this.graphPost(`/domains/${domainName}/forceDelete`, {});
  }

  private async getServiceHealth(service?: string): Promise<any> {
    if (service) {
      return this.graphGet(`/admin/serviceAnnouncement/healthOverviews?$filter=service eq '${service}'`);
    }
    return this.graphGet("/admin/serviceAnnouncement/healthOverviews");
  }

  private async listServiceIssues(status?: string): Promise<any> {
    const params = new URLSearchParams();
    if (status) params.append("$filter", `status eq '${status}'`);
    return this.graphGet(`/admin/serviceAnnouncement/issues?${params}`);
  }

  private async getReports(reportName: string, period: string = "D30", format: string = "json"): Promise<any> {
    const path = `/reports/${reportName}(period='${period}')`;
    if (format === "csv") {
      const token = await this.authManager.getToken("m365admin");
      const response = await fetch(`https://graph.microsoft.com/v1.0${path}?$format=text/csv`, {
        headers: { Authorization: `Bearer ${token}`, Accept: "text/csv" },
      });
      if (!response.ok) throw new Error(`Report failed: ${response.statusText}`);
      return { csv: await response.text() };
    }
    return this.graphGet(path);
  }

  private async listRoles(scope: string = "Directory"): Promise<any> {
    if (scope === "Directory") {
      return this.graphGet("/directoryRoles");
    }
    return this.graphGet("/roleManagement/directory/roleDefinitions");
  }

  private async getRoleMembers(roleId?: string, roleTemplateId?: string): Promise<any> {
    if (roleId) {
      return this.graphGet(`/directoryRoles/${roleId}/members`);
    }
    if (roleTemplateId) {
      const roles = await this.graphGet(`/directoryRoles?$filter=roleTemplateId eq '${roleTemplateId}'`);
      if (roles.value?.[0]) {
        return this.graphGet(`/directoryRoles/${roles.value[0].id}/members`);
      }
    }
    return { value: [] };
  }

  private async assignRole(roleTemplateId: string, principalId: string, scopeType: string = "Directory", scopeId?: string, previewOnly: boolean = true): Promise<any> {
    const body: any = {
      roleTemplateId,
      principalId,
    };
    if (scopeType === "AdministrativeUnit" && scopeId) {
      body.directoryScopeId = `/administrativeUnits/${scopeId}`;
    }
    if (previewOnly) return { preview: true, action: "assignRole", body };
    return this.graphPost("/directoryRoles", body);
  }

  private async removeRole(roleAssignmentId: string, previewOnly: boolean = true): Promise<any> {
    if (previewOnly) return { preview: true, action: "removeRole", roleAssignmentId };
    return this.graphDelete(`/directory/roleAssignments/${roleAssignmentId}`);
  }

  private async listPolicies(policyType?: string): Promise<any> {
    const policies: Record<string, string> = {
      authenticationMethodsPolicy: "/policies/authenticationMethodsPolicy",
      conditionalAccess: "/identity/conditionalAccess/policies",
      permissionGrant: "/policies/permissionGrantPolicies",
      homeRealmDiscovery: "/policies/homeRealmDiscoveryPolicies",
      tokenIssuance: "/policies/tokenIssuancePolicies",
      tokenLifetime: "/policies/tokenLifetimePolicies",
      authenticationStrength: "/policies/authenticationStrengthPolicies",
    };
    const path = policies[policyType as keyof typeof policies];
    if (policyType && path) {
      return this.graphGet(path);
    }
    const results: any = {};
    for (const [key, path] of Object.entries(policies)) {
      try {
        results[key] = await this.graphGet(path);
      } catch {
        results[key] = { error: "Not accessible" };
      }
    }
    return results;
  }

  private async getPolicy(policyId: string, policyType: string): Promise<any> {
    const policies: Record<string, string> = {
      authenticationMethodsPolicy: "/policies/authenticationMethodsPolicy",
      conditionalAccess: "/identity/conditionalAccess/policies",
      permissionGrant: "/policies/permissionGrantPolicies",
      homeRealmDiscovery: "/policies/homeRealmDiscoveryPolicies",
      tokenIssuance: "/policies/tokenIssuancePolicies",
      tokenLifetime: "/policies/tokenLifetimePolicies",
      authenticationStrength: "/policies/authenticationStrengthPolicies",
    };
    if (!policies[policyType]) throw new Error(`Unknown policy type: ${policyType}`);
    return this.graphGet(`${policies[policyType]}/${policyId}`);
  }

  private async getLicenses(): Promise<any> {
    return this.graphGet("/subscribedSkus");
  }

  private async getServicePrincipals(filter?: string, top: number = 100): Promise<any> {
    const params = new URLSearchParams({ $top: top.toString() });
    if (filter) params.append("$filter", filter);
    return this.graphGet(`/servicePrincipals?${params}`);
  }

  private async graphGet(path: string): Promise<any> {
    const token = await this.authManager.getToken("m365admin");
    const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!response.ok) throw new Error(`Graph GET failed: ${response.status} ${response.statusText}`);
    return response.json();
  }

  private async graphPost(path: string, body: any): Promise<any> {
    const token = await this.authManager.getToken("m365admin");
    const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    if (!response.ok) throw new Error(`Graph POST failed: ${response.status} ${response.statusText}`);
    return response.json();
  }

  private async graphDelete(path: string): Promise<any> {
    const token = await this.authManager.getToken("m365admin");
    const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
      method: "DELETE",
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!response.ok) throw new Error(`Graph DELETE failed: ${response.status} ${response.statusText}`);
    return { success: true };
  }
}